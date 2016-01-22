"""
Module for generating PPTX files from templates and excel files. 
"""


from __future__ import division, unicode_literals

import copy
import distutils.dir_util
import itertools
import os
import shutil
import sys
import tempfile
import warnings
import zipfile

from lxml import etree
import xlrd


def zip_pres_dir(src_dir, dst):
    """ Zipes ooxml directory to zip.
    :param src_dir: dir to zip
    :param dst: dir where to save
    """
    zf = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for root, dirs, files in os.walk(src_dir):
        for filename in files:
            full_path = os.path.join(root, filename)
            zf.write(full_path, os.path.relpath(full_path, src_dir))
    zf.close()


def unzip_pres(src):
    """ Unzipes template.pptx.
    :param src: .pptx template to unzip
    :return: path to directory with office xml files
    """
    tmp_dir = tempfile.mkdtemp()
    zfile = zipfile.ZipFile(src)
    zfile.extractall(tmp_dir)
    zfile.close()
    return tmp_dir


def C(start, end=None):
    """ Returns excel cells cpecified between [start; end] or only at start position.
    :param start: first cell
    :param end: last cell
    :return: cells to return
    """
    start = get_indices_from_name(start)
    if end is None:
        return [xlrd.cellname(*start)]

    end = get_indices_from_name(end)
    start, end = min(start, end), max(start, end)

    return [
        xlrd.cellname(row, col)
        for row in range(start[0], end[0]+1)
        for col in range(start[1], end[1]+1)
    ]


def col_to_num(col_str):
    """ Convert base26 column string to number. """
    col_num = 0
    for expn, char in enumerate(reversed(col_str)):
        col_num += (ord(char) - ord('A') + 1) * (26 ** expn)

    return col_num


def get_indices_from_name(name):
    """ Converts cell's name(A1, B2 etc) to number index.
    :param name: cell's name(A1, B2 etc)
    :return: cell's indexes
    """
    for i, c in enumerate(name):
        if not c.isalpha():
            break
    col = col_to_num(name[:i].upper())
    row = int(name[i:])

    return row-1, col-1


def alpha_range(start, stop):
    """ Returns chars between start char and stop char(A,D -> A,B,C,D).
    :param start: start char
    :param stop: stop char
    :return: list of chars
    """
    return [chr(x) for x in range(ord(start), ord(stop)+1)]


def remove(e):
    ''' Removes parent element from xml.
    :param e: ooxml element
    '''
    e.getparent().remove(e)


def get_or_create_child(parent, name):
    ''' Gets or creates xml element child.
    :param parent: xml element name
    :param name: xml parent's child name
    :return: xml parent's child
    '''
    e = parent.find(name)
    return e if e is not None else etree.SubElement(parent, name)


class XMLModifier(object):
    """
    Class helper to modify ooxml documents.
    """

    NS = {
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
        'cs': 'http://schemas.microsoft.com/office/drawing/2012/chartStyle',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    }

    def __init__(self, xml_file):
        """ Initialize with ooxml file.
        :param xml_file: ooxml file
        """
        self.xml_file = xml_file
        self.tree = etree.parse(xml_file)

    def xpath(self, xpath, e=None, error=True):
        """ Returns element via it's xpath.
        :param xpath: element's xpath
        :param e: ooxml element
        :param error: errors by default
        :return: xml elemet
        """
        elements = (self.tree if e is None else e).xpath(xpath, namespaces=self.NS)
        if not elements and error:
            raise ValueError('There is no elements for this XPath %s.' % xpath)
        return elements

    def write(self):
        """ Writes changes to xml. """
        self.tree.write(self.xml_file)


class ChartFiller(XMLModifier):
    """
    Class fills xml files
    """
    def __init__(self, generator, path):
        super(ChartFiller, self).__init__(path)
        self.generator = generator

    def fill_data(self, data, conv=None):
        """ Fills template from excel data.
        :param data: excel data
        """
        for ser_index, ser_data in enumerate(data):
            self.fill_series(ser_data, ser_index, conv=conv)

    def fill_series(self, data, n, conv=None):
        """ Fills template from excel values.
        :param data: excel values
        :param n: numbers
        """
        numCache = self.xpath('//c:ser[c:idx[@val="%s"]]//c:numCache' % n)[0]
        for e in self.xpath('c:pt', numCache):
            remove(e)
        for pt_index, pt_val in enumerate(data):
            if conv:
                pt_val = conv(pt_val)
            pt_e = etree.SubElement(numCache, '{%s}pt' % self.NS['c'])
            pt_e.set('idx', str(pt_index))
            val_e = etree.SubElement(pt_e, '{%s}v' % self.NS['c'])
            val_e.text = str(pt_val)

    def fill_from_cells(self, cells, conv=None):
        """ Fills chart with values from excel.
        :param cells: excel cells
        """
        data = [
            [self.generator.get_cell(c) for c in ser]
            for ser in cells
        ]
        self.fill_data(data, conv=conv)

    @property
    def x(self):
        return float(self.xpath('//c:x')[0].get('val'))

    @property
    def w(self):
        return float(self.xpath('//c:w')[0].get('val'))

    @property
    def y(self):
        return float(self.xpath('//c:y')[0].get('val'))

    @property
    def h(self):
        return float(self.xpath('//c:h')[0].get('val'))

    def set_axis_max(self, val):
        """ Sets axis maximum.
        :param val: axis maximum
        """
        e = get_or_create_child(self.xpath('//c:valAx/c:scaling')[0], '{%s}max' % self.NS['c'])
        if val is None:
            remove(e)
        else:
            e.set('val', val)

    def set_axis_min(self, val):
        """ Sets axis minimum.
        :param val: axis minimum
        """
        e = get_or_create_child(self.xpath('//c:valAx/c:scaling')[0], '{%s}min' % self.NS['c'])
        if val is None:
            remove(e)
        else:
            e.set('val', val)

    def set_major_unit(self, val):
        """ Sets majorUnit.
        :param val: majorUnit value
        """
        e = get_or_create_child(self.xpath('//c:valAx')[0], '{%s}majorUnit' % self.NS['c'])
        if val is None:
            remove(e)
        else:
            e.set('val', val)


class PPTXGenerator(XMLModifier):
    """ Generates PPTX files from excel data and template """

    template_shape_names = ()
    separate_charts = None

    slide_number = 1
    sheet_name = None
    separate_charts_dir = 'separate-charts'

    def __init__(self, template, excel, dst='output.pptx', fill_empty=False, clean=False):
        self.fill_empty = fill_empty
        self.dst = dst
        self.clean = clean

        workbook = xlrd.open_workbook(excel)
        self.worksheet = None
        if self.sheet_name:
            try:
                self.worksheet = workbook.sheet_by_name(self.sheet_name)
            except xlrd.XLRDError:
                pass
        if not self.worksheet:
            self.worksheet = workbook.sheet_by_index(0)

        self.tmp_dir = unzip_pres(template)

        try:
            self.generate()
        finally:
            shutil.rmtree(self.tmp_dir)

    def generate(self):
        """ Generates presentation from template and excel values. """
        super(PPTXGenerator, self).__init__(self.get_slide_path())

        self.tree = etree.parse(self.get_slide_path())
        self.last_id = max(int(e.get('id')) for e in self.xpath('//*[@id]'))
        self.shapes = self.xpath('//p:spTree')[0]
        if self.clean:
            for x in ['p:sp', 'p:pic', 'p:cxnSp']:
                for e in self.xpath('//p:spTree/%s[not(*//p:cNvPr[@title])]' % x, error=False):
                    remove(e)

        self.template_shapes = {}
        not_found = []
        for name in self.template_shape_names:
            elements = self.xpath('//p:spTree//*[*/p:cNvPr[@title="%s"]]' % name, error=False)
            if len(elements) == 0:
                not_found.append(name)
                continue
            if len(elements) > 1:
                warnings.warn(
                    '%s elements with title `%s` were found. '
                    'First will be used as template, others will be removed.' % (len(elements), name)
                )
            self.template_shapes[name] = elements[0]
            for e in elements:
                remove(e)

        if not_found:
            print "Elements with these titles are required but weren't found: %s." % not_found
            sys.exit(1)

        self.shapes = {
            self.xpath('p:nvSpPr/p:cNvPr', e)[0].get('title'): e
            for e in self.xpath('//p:sp')
        }
        self.fill_values()

        self.write()

        zip_pres_dir(self.tmp_dir, self.dst)

        def get_sldSz(root_dir):
            """ Returns slide size.
            :param path: template path
            :return: slide size
            """
             return etree.parse(
                os.path.join(root_dir, 'ppt', 'presentation.xml')
            ).xpath('//p:sldSz', namespaces=self.NS)[0]

        def get_proportions(sldSz):
            """ Returns slide proportions.
            :param sldSz: slide size
            :return: slide proportions
            """
            return int(sldSz.get('cx'))/int(sldSz.get('cy'))

        def set_slide_sizes(root_dir, e_w, e_h):
            """ Returns slide sizes.
            :param root_dir: template path
            :param e_w: slide width
            :param e_h: slide height
            :return: slide width and height
            """
            path = os.path.join(root_dir, 'ppt', 'presentation.xml')
            pres_tree = etree.parse(path)
            sldSz = pres_tree.xpath('//p:sldSz', namespaces=self.NS)[0]

            proportions = e_w/e_h
            if proportions < main_proportions:
                e_w = int(round(e_h*main_proportions, 0))
            else:
                e_h = int(round(e_w/main_proportions, 0))

            sldSz.set('cx', str(e_w))
            sldSz.set('cy', str(e_h))
            pres_tree.write(path)

            return e_w, e_h

        sldSz = get_sldSz(self.tmp_dir)
        main_proportions = get_proportions(sldSz)

        for element_name, path_name in (self.separate_charts or {}).items():
            e = copy.deepcopy(self.E(element_name))
            e_w, e_h = self.get_element_sizes(e)
            nvGrpSpPr = copy.deepcopy(self.xpath('//p:spTree/p:nvGrpSpPr')[0])
            grpSpPr = copy.deepcopy(self.xpath('//p:spTree/p:grpSpPr')[0])
            separate_tmp_dir = tempfile.mkdtemp()
            try:
                distutils.dir_util.copy_tree(self.tmp_dir, separate_tmp_dir)
                slide_w, slide_h = set_slide_sizes(separate_tmp_dir, e_w, e_h)
                self.set_element_pos(e, (slide_w-e_w)/2, (slide_h-e_h)/2)
                tree = etree.parse(self.get_slide_path(path=separate_tmp_dir))
                spTree = tree.xpath('//p:spTree', namespaces=self.NS)[0]
                spTree.clear()
                spTree.append(nvGrpSpPr)
                spTree.append(grpSpPr)
                spTree.append(e)
                tree.write(self.get_slide_path(path=separate_tmp_dir))

                try:
                    os.mkdir(self.separate_charts_dir)
                except OSError:
                    pass
                zf = zipfile.ZipFile(os.path.join(self.separate_charts_dir, path_name), "w", zipfile.ZIP_DEFLATED)
                for root, dirs, files in os.walk(separate_tmp_dir):
                    for filename in files:
                        full_path = os.path.join(root, filename)
                        zf.write(full_path, os.path.relpath(full_path, separate_tmp_dir))
                zf.close()
            finally:
                shutil.rmtree(separate_tmp_dir)

    def get_relations(self):
        """ Returns slide relations.
        :return: dictionary of relations
        """
        path = os.path.join(self.tmp_dir, 'ppt', 'slides', '_rels', 'slide%s.xml.rels' % self.slide_number)
        rels = etree.parse(path).getroot()
        return {
            r.get('Id'): os.path.abspath(os.path.join(self.get_slide_path(), '..', r.get('Target'))) for r in rels
        }

    def get_elements_by_title(self, title, error=True):
        """ Returns list of elements choosed by title.
        :param title: ooxml element's title
        :param error: errors, True by default
        :return: list of elements
        """
        return self.xpath('//p:spTree//*[*/p:cNvPr[@title="%s"]]' % title, error=error)

    def get_element_by_title(self, title):
        """ Returns first element choosed by title.
        :param title: ooxml element title
        :return: first ooxml element
        """
        xpath = '//p:spTree//*[*/p:cNvPr[@title="%s"]]' % title
        elements = self.xpath(xpath)
        if len(elements) > 1:
            warnings.warn('The only element is expected at XPath `%s`, %s found.' % (xpath, len(elements)))
            for e in elements:
                print etree.tostring(e)
        return elements[0]
    E = get_element_by_title

    def get_slide_path(self, path=None):
        """ Returns slide path.
        :param path: slide path, by default None
        :return: slide path
        """
        return os.path.join(path or self.tmp_dir, 'ppt', 'slides', 'slide%s.xml' % self.slide_number)

    def with_comma(self, v):
        """ Adds comma if value is 4 digits or more.
        :param v: value
        :return: value with comma (321,234; 324,4 etc)
        """
        s = self.format_float(v, prec=0)
        if len(s) > 3:
            s = s[:-3]+','+s[-3:]
        return s

    def format_float(self, val, prec=1, strip=True, maximum=None, include_sign=False):
        """ Format float for template.
        :param val: value
        :return: formatted float
        """
        if maximum and val > maximum:
            val = maximum
        formatted = str(round(val, prec))

        sign = ''
        if include_sign and val > 0:
            sign = '+'
        return sign+(formatted.rstrip('0').rstrip('.') if strip else formatted)

    def format_percent(self, val, prec=None, strip=None, include_sign=None):
        """ Formats percent.
        :param val: value
        """
        kwargs = {}
        if prec is not None:
            kwargs['prec'] = prec
        if strip is not None:
            kwargs['strip'] = strip
        if include_sign is not None:
            kwargs['include_sign'] = include_sign
        return self.format_float(val*100, **kwargs)

    def format_money(self, val, prec=None, strip=None, include_sign=None):
        """ Formats money value for template.
        :param val: value
        :return: formated string
        """
        kwargs = {}
        if prec is not None:
            kwargs['prec'] = prec
        if strip is not None:
            kwargs['strip'] = strip

        return (('+' if include_sign else '') if val > 0 else '-') + '$' + self.format_float(abs(val), **kwargs)

    def get_simple_fillers(self):
        return {}

    def get_lines_fillers(self):
        return []

    def fill_values(self):
        """
        Fills values from excel to template elements.
        """
        filled_cells = set()
        duplicated = set()

        for conv, cells in self.get_simple_fillers().iteritems():
            for cell in itertools.chain.from_iterable(cells):
                (duplicated if cell in filled_cells else filled_cells).add(cell)
                val = self.get_cell(cell)
                try:
                    val = conv(val)
                except Exception as e:
                    warnings.warn(
                        'Could not prepare value `%s` from cell `%s`.\nError was: %s' % (
                            val,
                            cell,
                            e,
                        )
                    )
                    val = 'INVALID VALUE'
                try:
                    self.set_text(cell, val)
                except Exception as e:
                    warnings.warn(
                        'Could not set text for element %s.\nError was: %s' % (
                            cell,
                            e,
                        )
                    )

        for max_chars, cells in self.get_lines_fillers():
            for cell in itertools.chain.from_iterable(cells):
                (duplicated if cell in filled_cells else filled_cells).add(cell)
                val = self.get_cell(cell)
                lines = [word[:max_chars] for word in val.split()]
                self.set_element_text_lines(self.E(cell), lines)

        if duplicated:
            warnings.warn('These cell were filled several times:\n%s' % duplicated)

    def get_chart(self, path):
        """ Returns chart by path.
        :param path: chart path
        :return: Chart Filler
        """
        return ChartFiller(self, path)

    def get_chart_by_title(self, title):
        """ Returns chart by title.
        :param title: chart title
        :return: chart
        """
        e = self.get_element_by_title(title)
        chart = self.xpath('.//c:chart', e)[0]
        id = chart.get('{%s}id' % self.NS['r'])
        return self.get_chart(self.get_relations()[id])

    def get_chart_path(self, chart):
        """ Returns chart path.
        :param chart: chart name
        :return: chart path
        """
        return os.path.join(self.tmp_dir, 'ppt', 'charts', 'chart%s.xml' % chart)

    def fill_chart(self, chart, data, conv=None):
        """ Fills chart with data.
        :param chart: chart name
        :param data: data from excel
        """
        path = self.get_chart_path(chart)
        chart = self.get_chart(path)
        chart.fill_data(data, conv=conv)
        chart.write()

    def get_id(self):
        """ Generates unique id.
        :return: unique id
        """
        self.last_id += 1
        return self.last_id

    def get_cell(self, name):
        """ Returns cell specified by name.
        :param name: excel cell name (A1, B4 etc)
        :return: cell
        """
        row, col = get_indices_from_name(name)
        return self.worksheet.row_values(row)[col]

    def set_text(self, name, text):
        """ Set text on shape.
        :param name: shape's name
        :param text: text
        """
        shape = self.shapes.get(name)
        if shape is None:
            warnings.warn("Tried to set text for %s, but the element wasn't found." % name, stacklevel=2)
            return

        self.set_element_text(shape, text)

    def set_element_text(self, shape, text):
        """ Sets element's text.
        :param shape: element's shape
        :param text: text
        """
        elements = self.xpath('.//a:t', shape)
        elements[0].text = '' if self.fill_empty else text

        for e in elements[1:]:
            remove(e.getparent())

    def set_element_text_color(self, shape, color):
        """ Sets element text color.
        :param shape: shape
        :param color: hex color
        """
        solidFill = self.xpath('.//a:solidFill', shape)[0]
        solidFill.clear()
        if etree.iselement(color):
            solidFill.append(color)
        else:
            srgbClr = etree.SubElement(solidFill, "{%s}srgbClr" % self.NS['a'])
            srgbClr.set('val', color)

    def get_element_fill_color(self, shape):
        """ Returns element's color.
        :param shape: element's shape
        :return: hex color
        """
        return self.xpath('.//a:solidFill/a:srgbClr', shape)[0].get('val')

    def set_element_text_direction(self, shape, direction):
        """ Sets text direction.
        :param shape: element's shape
        :param direction: text dirction
        """
        bodyPr = self.xpath('.//a:bodyPr', shape)[0]
        if direction:
            bodyPr.set('vert', direction)
        else:
            bodyPr.attrib.pop('vert', None)

    def set_element_text_alignment(self, shape, algn):
        """ Sets element's alignment.
        :param shape: element's shape
        :param algn: element's alignment
        """
        pPr = self.xpath('.//a:pPr', shape)[0]
        if algn:
            pPr.set('algn', algn)
        else:
            pPr.attrib.pop('algn', None)

    def set_element_text_size(self, shape, size):
        """ Sets element's text size.
        :param shape: element's shape
        :param size: element's text size
        """
        rPr = self.xpath('.//a:rPr', shape)[0]
        rPr.set('sz', str(int(round(float(size), 0)*100)))

    def set_element_text_lines(self, shape, lines):
        """ Sets textBox for text.
        :param shape: shape type
        :param lines: number of text lines
        """
        txBody = self.xpath('.//p:txBody', shape)[0]

        ps = self.xpath('.//a:p', txBody)
        first_p = ps[0]
        for r in self.xpath('.//a:r', first_p)[1:]:
            remove(r)

        for p in ps:
            remove(p)

        for line in lines:
            p = copy.deepcopy(first_p)
            t = self.xpath('.//a:t', p)[0]
            t.text = '' if self.fill_empty else line
            txBody.append(p)

    def add_shape(self, e):
        """ Adds new element to xml.
        :param e: ooxml element
        """
        cNvPr = self.xpath('.//p:cNvPr', e)[0]
        i = self.get_id()
        cNvPr.set('id', str(i))
        cNvPr.set('name', '%s %s' % (cNvPr.get('name'), i))
        self.xpath('//p:spTree')[0].append(e)

    def set_element_pos(self, e, x, y):
        """ Function sets element (x,y) position.
        :param e: ooxml element
        :param x: element's x coordinate
        :param y: element's y coordinate
        """
        off = self.xpath('.//a:xfrm/a:off', e)[0]
        if x is not None:
            off.attrib['x'] = str(int(x))
        if y is not None:
            off.attrib['y'] = str(int(y))

    def mod_element_pos(self, e, x, y):
        """ Function modifies existing element (x,y) position.
        :param e: ooxml element
        :param x: element's x coordinate
        :param y: element's y coordinate
        """
        off = self.xpath('.//a:xfrm/a:off', e)[0]
        x_old = off.get('x')
        y_old = off.get('y')

        x_new = int(x_old)+x if x is not None else None
        y_new = int(y_old)+y if y is not None else None

        self.set_element_pos(e, x_new, y_new)

    def set_element_size(self, e, w, h):
        """ Function sets element's size(cx, cy).
        :param e: ooxml element
        :param w: element's width
        :param h: element's height
        """
        ext = self.xpath('.//a:xfrm/a:ext', e)[0]
        if w is not None:
            ext.attrib['cx'] = str(int(w))
        if h is not None:
            ext.attrib['cy'] = str(int(h))

    def clone_template(self, name):
        """ Function clones template.
        :param name: template name
        :return: template copy
        """
        return copy.deepcopy(self.template_shapes[name])

    def add_line(self, x0, y0, x1, y1, template_name, shapes=None):
        """ Function creates line.
        :param x0: line's x beginning
        :param y0: line's y beginning
        :param x1: line's x end
        :param y1: line's y end
        :param template_name: specify template name where to create line
        """
        flipH = flipV = False
        if x0 > x1:
            x0, x1 = x1, x0
            flipH = True
        if y0 > y1:
            y0, y1 = y1, y0
            flipV = True

        x0, y0, x1, y1 = [int(round(v, 0)) for v in [x0, y0, x1, y1]]

        e = self.clone_template(template_name)
        xfrm = self.xpath('.//a:xfrm', e)[0]
        ext = self.xpath('a:ext', xfrm)[0]
        self.xpath('.//p:cNvCxnSpPr', e)[0].clear()

        if flipH:
            xfrm.attrib['flipH'] = '1'
        elif 'flipH' in xfrm.attrib:
            del xfrm.attrib['flipH']

        if flipV:
            xfrm.attrib['flipV'] = '1'
        elif 'flipV' in xfrm.attrib:
            del xfrm.attrib['flipV']

        ext.attrib['cx'] = str(x1-x0)
        ext.attrib['cy'] = str(y1-y0)

        self.set_element_pos(e, x0, y0)
        self.add_shape(e) if shapes is None else shapes.append(e)

    def add_circle(self, x, y, template_name, shapes=None):
        """ Add's circle to tamplate.
        :param x: circle's x
        :param y: circle's y
        :param template_name: specify template name where to create circle
        """
        e = self.clone_template(template_name)
        ext = self.xpath('.//a:ext', e)[0]

        x -= int(ext.get('cx'))/2
        y -= int(ext.get('cy'))/2

        self.set_element_pos(e, x, y)
        self.add_shape(e) if shapes is None else shapes.append(e)

    def get_shape_coords(self, name):
        """ Returns elements coordinates
        :param name: shape name
        :return: returns element's
        """
        return self.get_element_coords(self.shapes[name])

    def get_element_sizes(self, e):
        """ Returns element size
        :param e: element
        :return: (width, height)
        """
        coords = self.get_element_coords(e)
        return coords[2]-coords[0], coords[3]-coords[1]

    def get_element_coords(self, e):
        """ Returns element's coordinates.
        :param e: ooxml element
        :return: x0, yo, x1, y1
        """
        e = self.xpath('.//a:xfrm|.//p:xfrm', e)[0]
        off = self.xpath('a:off', e)[0]
        ext = self.xpath('a:ext', e)[0]

        x0 = int(off.get('x'))
        y0 = int(off.get('y'))
        x1 = x0+int(ext.get('cx'))
        y1 = y0+int(ext.get('cy'))

        return x0, y0, x1, y1

    def get_plot_area_coords_from_chart(self, chart_name):
        """ Gets element's coordinates.
        :param chart_name: element's name
        :return: (x0, y0, x1, y1)
        """
        chart_e = self.get_element_by_title(chart_name)
        coords = self.get_element_coords(chart_e)
        chart = self.get_chart_by_title(chart_name)
        chart_w = coords[2]-coords[0]
        chart_h = coords[3]-coords[1]

        x0 = coords[0]+chart_w*chart.x
        y0 = coords[1]+chart_h*chart.y
        x1 = x0+chart_w*chart.w
        y1 = y0+chart_h*chart.h

        return tuple(int(round(x, 0)) for x in [x0, y0, x1, y1])

    def replace_pic_color(self, pic_e, color):
        """ Function replace element's color.
        :param pic_e: ooxml element
        :param color: hex color
        """
        blib_e = self.xpath('.//a:blip', pic_e)[0]
        clrRepl = etree.Element('{%s}clrRepl' % self.NS['a'])
        blib_e.insert(0, clrRepl)
        etree.SubElement(clrRepl, '{%s}srgbClr' % self.NS['a']).set('val', color)

    def get_element_font_size(self, e):
        """ Returns element's front size
        :param e: ooxml element
        :return: element's front size
        """
        return int(self.xpath('.//a:rPr', e)[0].get('sz'))

    def set_element_font_size(self, e, size):
        """ Sets element's front size.
        :param e: ooxml element
        :param size: element's front size
        """
        self.xpath('.//a:rPr', e)[0].set('sz', str(int(round(size, 0))))

    def set_element_rotation(self, e, angle):
        """ Sets element's rotation angle.
        :param e: ooxml element
        :param angle: rotation angle
        """
        xfrm = self.xpath('.//a:xfrm', e)[0]
        xfrm.set('rot', str(int(round(angle, 0))))

    def set_element_flipv(self, e, flag):
        """ Sets element vertical flip.
        :param e: ooxml element
        :param flag: rotation True or False
        """
        xfrm = self.xpath('.//a:xfrm', e)[0]
        if flag:
            xfrm.set('flipV', '1')
        elif 'flipV' in xfrm.attrib:
            del xfrm.attrib['flipV']

    def set_element_fliph(self, e, flag):
        """ Sets element horizontal flip.
        :param e: ooxml element
        :param flag: flip True or False
        """
        xfrm = self.xpath('.//a:xfrm', e)[0]
        if flag:
            xfrm.set('flipH', '1')
        elif 'flipH' in xfrm.attrib:
            del xfrm.attrib['flipH']


class CMDHandler(object):
    """
    Class that handles command line options and starts template generation.
    """
    def __init__(self, dashboard_class, template_path=None, dst=None):
        if len(sys.argv) <= 1:
            print "Data file is not specified!"
            sys.exit(1)

        dashboard_class(
            template_path or os.path.join(os.path.dirname(sys.argv[0]), 'template.pptx'),
            sys.argv[-1],
            fill_empty='--fill-empty' in sys.argv,
            clean='--clean' in sys.argv,
            dst=dst or 'output.pptx',
        )

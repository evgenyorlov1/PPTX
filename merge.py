#!/usr/bin/env python

from __future__ import unicode_literals

import argparse
import copy
import shutil
import os
from os import path

from lxml import etree


from data2ppt import XMLModifier, zip_pres_dir, unzip_pres

OUTFILE = 'output.pptx'


class PresentationMerger(object):

    def __init__(self, presentations):
        self.presentations = presentations

    def merge(self):
        self.dirs = [unzip_pres(pres) for pres in self.presentations]
        try:
            self.main_dir = self.dirs[0]

            self.content_types = XMLModifier(path.join(self.main_dir, '[Content_Types].xml'))
            self.default_content_types = {
                e.get('Extension') for e in self.content_types.tree.getroot().iter(tag='{*}Default')
            }
            self.pres_rels = XMLModifier(path.join(self.main_dir, 'ppt', '_rels', 'presentation.xml.rels'))
            self.pres = XMLModifier(path.join(self.main_dir, 'ppt', 'presentation.xml'))
            self.last_slide_id = int(self.pres.xpath('//p:sldId')[0].get('id'))

            for i, d in enumerate(self.dirs[1:], 2):
                self.merge_presentation_dir(d, i)

            self.content_types.write()
            self.pres_rels.write()
            self.pres.write()

            zip_pres_dir(self.main_dir, OUTFILE)
        finally:
            for d in self.dirs:
                shutil.rmtree(d)

    def add_content_type(self, path, content_type):
        etree.SubElement(
            self.content_types.tree.getroot(),
            'Override',
            PartName=path,
            ContentType=content_type,
        )

    def add_default_content_type(self, extension, content_type):
        if extension in self.default_content_types:
            return

        self.default_content_types.add(extension)
        etree.SubElement(
            self.content_types.tree.getroot(),
            'Default',
            Extension=extension,
            ContentType=content_type,
        )

    def merge_presentation_dir(self, d, n):
        slide_number = 1

        self.add_content_type(
            path.join('/', 'ppt', 'slides', '%sslide%s.xml' % (n, slide_number)),
            "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
        )
        self.last_slide_id += 1
        sldId = etree.SubElement(
            self.pres.xpath('//p:sldIdLst')[0],
            '{%s}sldId' % self.pres.NS['p'],
            id=str(self.last_slide_id)
        )
        sldId.attrib['{%s}id' % self.pres.NS['r']] = "rId%s" % (n*100)
        etree.SubElement(
            self.pres_rels.tree.getroot(),
            'Relationship',
            Id="rId%s" % (n*100),
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
            Target=path.join('slides', '%sslide%s.xml' % (n, slide_number)),
        )

        content_types = etree.parse(path.join(d, '[Content_Types].xml'))
        for e in content_types.getroot().iter(tag='{*}Default'):
            self.add_default_content_type(e.get('Extension'), e.get('ContentType'))
        self.merge_related(
            path.join('ppt', 'slides', 'slide%s.xml' % slide_number),
            d,
            n,
            {},
            content_types,
        )

    def merge_related(self, path_, d, n, merged, content_types):
        path_ = path.abspath(path.join('/', path_))
        if path_ in merged:
            return

        head, tail = path.split(path_)
        new_path = path.join(head, '%s%s' % (n, tail))

        merged[path_] = new_path


        try:
            os.makedirs(path.dirname(path.join(self.main_dir, path.relpath(new_path, '/'))))
        except OSError:
            pass
        shutil.copy(
            path.join(d, path.relpath(path_, '/')),
            path.join(self.main_dir, path.relpath(new_path, '/')),
        )

        rels_path = path.relpath(path.join(head, '_rels', '%s.rels' % tail), '/')
        if path.exists(path.join(d, rels_path)):
            new_rels_path = path.relpath(path.join(head, '_rels', '%s%s.rels' % (n, tail)), '/')
            non_follow_types = [
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster',
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject',
            ]
            rels = etree.parse(path.join(d, rels_path))
            for e in rels.getroot():
                if e.get('Type') in non_follow_types:
                    continue
                self.merge_related(path.join(path_, '..', e.get('Target')), d, n, merged, content_types)

            for e in rels.getroot():
                if e.get('Type') in non_follow_types:
                    continue
                new_abs_path = path.join('/', merged[path.normpath(path.join(path_, '..', e.get('Target')))])
                ext = e.get('Target').rsplit('.', 1)[-1]
                if ext == 'xml' or ext not in self.default_content_types:
                    for c in content_types.getroot():
                        if c.get('PartName') == path.abspath(path.join('/', path_, '..', e.get('Target'))):
                            if not any(c.get('PartName') == new_abs_path for c in self.content_types.tree.getroot()):
                                c = copy.deepcopy(c)
                                c.set('PartName', new_abs_path)
                                self.content_types.tree.getroot().append(c)
                            break
                    else:
                        raise Exception
                e.set(
                    'Target',
                    new_abs_path,
                )
            try:
                os.makedirs(path.dirname(path.join(self.main_dir, new_rels_path)))
            except OSError:
                pass
            rels.write(path.join(self.main_dir, new_rels_path))
            path.join(self.main_dir, new_rels_path)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('pptx-file', nargs='+')
    args = parser.parse_args()

    merger = PresentationMerger(vars(args)['pptx-file'])
    merger.merge()

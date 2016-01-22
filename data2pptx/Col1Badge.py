from lxml import etree
import os


namespace = {'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
           'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'}


def Col1Badge(excel):
    title = 'Col 1 Badge'
    path = os.path.join(os.path.split(os.path.abspath(__file__))[0], 'ppt/slides/slide1.xml') #depends on dashboard
    tree = etree.parse(path)
    element = tree.xpath('//*[@title="Col 1 Badge"]', namespaces=namespace)
    print element[0].attrib
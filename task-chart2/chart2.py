from lxml import etree
import os


namespace = {'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
           'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p':'http://schemas.openxmlformats.org/presentationml/2006/main'}


def chart2(excel):

    B16 = round(float(excel[15][1])*100, 1) #Line 1 Square
    D16 = round(float(excel[15][3])*100, 1) #Line 1 Circle
    B17 = round(float(excel[16][1])*100, 1) #Line 2 Square
    D17 = round(float(excel[16][3])*100, 1) #Line 2 Circle
    B18 = round(float(excel[17][1])*100, 1) #Line 3 Square
    D18 = round(float(excel[17][3])*100, 1) #Line 3 Circle
    B19 = round(float(excel[18][1])*100, 1) #Line 4 Square
    D19 = round(float(excel[18][3])*100, 1) #Line 4 Circle
    B20 = round(float(excel[19][1])*100, 1) #Line 5 Square
    D20 = round(float(excel[19][3])*100, 1) #Line 5 Circle


    path = os.path.join(os.path.split(os.path.abspath(__file__))[0], 'ppt/slides/slide1.xml')
    tree = etree.parse(path)

    LeftStraightConnector = tree.xpath('/p:sld/p:cSld/p:spTree/p:cxnSp[2]/p:spPr/a:xfrm/a:off', namespaces=namespace)
    Left = int(LeftStraightConnector[0].get('x')) #left corner coordinates
    RightStraightConnector = tree.xpath('/p:sld/p:cSld/p:spTree/p:cxnSp[6]/p:spPr/a:xfrm/a:off', namespaces=namespace)
    Right = int(RightStraightConnector[0].get('x')) #right corner coordinates
    zeroStraightConnector = tree.xpath('/p:sld/p:cSld/p:spTree/p:cxnSp[7]/p:spPr/a:xfrm/a:off', namespaces=namespace)
    Zero = int(zeroStraightConnector[0].get('x')) #zero coordinates
    leftStep = (Zero - Left)/10 #left part step [0; 10]
    rightStep = (Right - Zero)/15 #right part step [0; 15]
    SquareWidth = int(tree.xpath('/p:sld/p:cSld/p:spTree/p:grpSp[1]/p:grpSpPr/a:xfrm/a:ext', namespaces=namespace)[0].get('cx'))/2 #width of square element
    CircleWidth = int(tree.xpath('/p:sld/p:cSld/p:spTree/p:sp[15]/p:spPr/a:xfrm/a:ext', namespaces=namespace)[0].get('cx'))/2 #width of circle element

    Line1Square = tree.xpath('//*[@title="Line 1 Square"]', namespaces=namespace)
    Square1 = Line1Square[0].getparent().getparent().xpath('.//p:grpSpPr/a:xfrm/a:off', namespaces=namespace)[0]
    if(B16 > 0):
        Square1.attrib['x'] = unicode(Zero + int(rightStep*B16) - SquareWidth)
    else:
        Square1.attrib['x'] = unicode(Zero - int(abs(leftStep*B16)) - SquareWidth)

    Line1Circle = tree.xpath('//*[@title="Line 1 Circle"]', namespaces=namespace)
    Circle1 = Line1Circle[0].getparent().getparent().xpath('.//p:spPr/a:xfrm/a:off', namespaces=namespace)[0]
    if(D16 > 0):
        Circle1.attrib['x'] = unicode(Zero + int(rightStep*D16) - CircleWidth)
    else:
        Circle1.attrib['x'] = unicode(Zero - int(abs(leftStep*D16)) - CircleWidth)

    Line2Square = tree.xpath('//*[@title="Line 2 Square"]', namespaces=namespace)
    Square2 = Line2Square[0].getparent().getparent().xpath('.//p:grpSpPr/a:xfrm/a:off', namespaces=namespace)[0]
    if(B17 > 0):
        Square2.attrib['x'] = unicode(Zero + int(rightStep*B17) - SquareWidth)
    else:
        Square2.attrib['x'] = unicode(Zero - int(abs(leftStep*B17)) - SquareWidth)

    Line2Circle = tree.xpath('//*[@title="Line 2 Circle"]', namespaces=namespace)
    Circle2 = Line2Circle[0].getparent().getparent().xpath('.//p:spPr/a:xfrm/a:off', namespaces=namespace)[0]
    if(D17 > 0):
        Circle2.attrib['x'] = unicode(Zero + int(rightStep*D17) - CircleWidth)
    else:
        Circle2.attrib['x'] = unicode(Zero - int(abs(leftStep*D17)) - CircleWidth)

    Line3Square = tree.xpath('//*[@title="Line 3 Square"]', namespaces=namespace)
    Square3 = Line3Square[0].getparent().getparent().xpath('.//p:grpSpPr/a:xfrm/a:off', namespaces=namespace)[0]
    if(B18 > 0):
        Square3.attrib['x'] = unicode(Zero + int(rightStep*B18) - SquareWidth)
    else:
        Square3.attrib['x'] = unicode(Zero - int(abs(leftStep*B18)) - SquareWidth)

    Line3Circle = tree.xpath('//*[@title="Line 3 Circle"]', namespaces=namespace)
    Circle3 = Line3Circle[0].getparent().getparent().xpath('.//p:spPr/a:xfrm/a:off', namespaces=namespace)[0]
    if(D18 > 0):
        Circle3.attrib['x'] = unicode(Zero + int(rightStep*D18) - CircleWidth)
    else:
        Circle3.attrib['x'] = unicode(Zero - int(abs(leftStep*D18)) - CircleWidth)

    Line4Square = tree.xpath('//*[@title="Line 4 Square"]', namespaces=namespace)
    Square4 = Line4Square[0].getparent().getparent().xpath('.//p:grpSpPr/a:xfrm/a:off', namespaces=namespace)[0]
    if(B19 > 0):
        Square4.attrib['x'] = unicode(Zero + int(rightStep*B19) - SquareWidth)
    else:
        Square4.attrib['x'] = unicode(Zero - int(abs(leftStep*B19)) - SquareWidth)

    Line4Circle = tree.xpath('//*[@title="Line 4 Circle"]', namespaces=namespace)
    Circle4 = Line4Circle[0].getparent().getparent().xpath('.//p:spPr/a:xfrm/a:off', namespaces=namespace)[0]
    if(D19 > 0):
        Circle4.attrib['x'] = unicode(Zero + int(rightStep*D19) - CircleWidth)
    else:
        Circle4.attrib['x'] = unicode(Zero - int(abs(leftStep*D19)) - CircleWidth)

    Line5Square = tree.xpath('//*[@title="Line 5 Square"]', namespaces=namespace)
    Square5 = Line5Square[0].getparent().getparent().xpath('.//p:grpSpPr/a:xfrm/a:off', namespaces=namespace)[0]
    if(B20 > 0):
        Square5.attrib['x'] = unicode(Zero + int(rightStep*B20) - SquareWidth)
    else:
        Square5.attrib['x'] = unicode(Zero - int(abs(leftStep*B20)) - SquareWidth)

    Line5Circle = tree.xpath('//*[@title="Line 5 Circle"]', namespaces=namespace)
    Circle5 = Line5Circle[0].getparent().getparent().xpath('.//p:spPr/a:xfrm/a:off', namespaces=namespace)[0]
    if(D20 > 0):
        Circle5.attrib['x'] = unicode(Zero + int(rightStep*D20) - CircleWidth)
    else:
        Circle5.attrib['x'] = unicode(Zero - int(abs(leftStep*D20)) - CircleWidth)

    with open(path, 'w') as file:
        xml = etree.tostring(tree, pretty_print=False)
        file.write(xml)
        file.close()
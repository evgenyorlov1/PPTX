import os
import zipfile
import shutil
import sys
import csv
import xlrd
import chart2
from lxml import etree

chart1Name = 'chart1.pptx'
chart2Name = 'chart2.pptx'
inputChart2 = 'chart2-template.pptx'


def zipFiles():
    global chart2Name
    zf = zipfile.ZipFile(chart2Name, "w")
    zf.write('[Content_Types].xml')
    for root, dirs, files in os.walk('ppt/'):
        for file in files:
            zf.write(os.path.join(root, file))
    for root, dirs, files in os.walk('_rels/'):
        for file in files:
            zf.write(os.path.join(root, file))
    for root, dirs, files in os.walk('docProps/'):
        for file in files:
            zf.write(os.path.join(root, file))
    zf.close()


def unzipFiles():
    global inputChart2
    zfile = zipfile.ZipFile(inputChart2)
    zfile.extractall()


def removeFiles():
    os.remove('[Content_Types].xml')
    shutil.rmtree('ppt/')
    shutil.rmtree('_rels/')
    shutil.rmtree('docProps/')


def excel2list(excel):
    workbook = xlrd.open_workbook(excel)
    sheet_names = list(workbook.sheet_names())
    worksheet = workbook.sheet_by_name(str(sheet_names[0])) #assume we have one sheet in workbook
    result = []
    for row in xrange(worksheet.nrows):
        result.append(
            list(x.encode('utf-8') if type(x) == type(u'') else x for x in worksheet.row_values(row))
        )
    return result


if __name__ == '__main__':
    excel = sys.argv[1]
    excelList = excel2list(excel)
    unzipFiles()
    chart2.chart2(excelList)
    zipFiles()
    removeFiles()

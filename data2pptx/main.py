import os
import zipfile
import shutil
import sys
import csv
import xlrd
import Col1Badge
import Col1FillHeight
from lxml import etree


outputFileName = 'output.pptx'
inputFileName = 'template.pptx'


def zipFiles():
    global outputFileName
    zf = zipfile.ZipFile(outputFileName, "w")
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
    global inputFileName
    zfile = zipfile.ZipFile(inputFileName)
    zfile.extractall()


def removeFiles():
    os.remove('[Content_Types].xml')
    shutil.rmtree('ppt/')
    shutil.rmtree('_rels/')
    shutil.rmtree('docProps/')


#transforms excel file 2 python list object
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
    Col1Badge.Col1Badge(excelList)              #here would be a coll to script that works with separate dashboard
    Col1FillHeight.Col1FillGeight(excelList)    #same
    zipFiles()
    removeFiles()

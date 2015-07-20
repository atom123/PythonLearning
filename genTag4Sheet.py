#!/usr/bin/python
#-*- coding:utf8 -*-
from xml.etree import ElementTree
import xlrd

def GenTag4Sheet(TableElement, SheetTagList, SheetTagValueList, SheetName):
	if ( TableElement.getchildren() ):
		if ( TableElement.tag != SheetName ):
			SheetTagList.append('_GRPSTART_' + TableElement.tag)
			SheetTagValueList.append('Required')

		for child in TableElement:
			GenTag4Sheet(child, SheetTagList, SheetTagValueList, SheetName)

		if ( TableElement.tag != SheetName ):
			SheetTagList.append('_GRPEND_' + TableElement.tag)
			SheetTagValueList.append('Required')
		
if __name__ == "__main__":

	OutFile = r'C:\Users\jeguan\Desktop\Test_2_bk.xml'
	outFile = r'C:\Users\jeguan\Desktop\test.xls'

	wb = xlrd.open_workbook(outFile)
	ws = wb.sheet_by_name('ClientAdmin-fsdb0')
	nrows = ws.nrows
	ncols = ws.ncols

	
#	for nrow in range(0, nrows):
#		for ncol in range(0, ncols):
#			ws.write(nrow, ncol, "")
#	
	root  = ElementTree.parse(OutFile).getroot()
	ListNode = root.iter('Response')
#	tables = root.findall("Response/" + "ClientAdmin-fsdb0")

	for node in ListNode:
		SheetTagList = []
		SheetTagListValue = []

		SheetTagList.append('XML_ACTION')
		SheetTagListValue.append('INSERT')

		GenTag4Sheet(root, SheetTagList, SheetTagListValue, "ClientAdmin-fsdb0")
		print(SheetTagList)
		print(SheetTagListValue)

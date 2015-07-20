#!/usr/bin/python
#-*- coding:utf-8 -*-

import os
import re
import xlrd
from xml.etree import ElementTree

###################################################################################
# Function Name: genXmlScript4FSDB(tabname, sheetname, inputWorkBook)
#
# Description:   generate xml script for FSDB.
#
# Input Value:   rdsname   --- the path name which used to store xml script for sheet
#                sheetname --- the sheet name
#                inputWorkBook --- Book object returned by xlrd.open_workbook(inputFile)
#
# Return Value:  if there is no record in the sheet, only return 
###################################################################################
def genXmlScript4FSDB(rdsname, sheetname, inputWorkBook):
		''' Module to generate the xml script for FSDB. '''

		wb = inputWorkBook
		ws = wb.sheet_by_name(sheetname)
		rowValuesList = ws.row_values(0)
		numRows = ws.nrows	# number of rows in this sheet.

		# no record exits in this sheet, so no need to go through
		if ( numRows <= 1 ):
			return 

		reCmd = re.compile(r'.+(?=-fsdb\d*)')
		tempSheetname = reCmd.findall(sheetname)

		Action = 'READ'
    	# set the root of xml script tree for request
		if ( tempSheetname[0] == 'ProtectionSiteAdmin' )
			Action = 'READALL'

		root = ElementTree.Element('Request',{'Action': Action, 'RequestId': '100000'})

		# go through all records in sheet named "ClientAdmin"
		if ( tempSheetname[0] == 'ClientAdmin' ):
			# to get the location column of "ClientName"
			for num_element in range(0, len(rowValuesList)):
				if rowValuesList[num_element] == 'ClientName':
					break
			# add a sub node named tempSheetname[0], tempSheetname is a list. 
			AttriSheet = ElementTree.Element(tempSheetname[0])
			SubAttriSheet = ElementTree.SubElement(AttriSheet, 'ClientName')

			for nrows in range(1, numRows):
				tempValue = str(ws.cell(nrows, num_element).value).strip()
				if ( tempValue != "" ):
					SubAttriSheet.text = str(tempValue)
					root.append(AttriSheet)
					tree = ElementTree.ElementTree(root)
					tree.write(sheetname + '.xml'+ str(nrows), "utf8")
					root.remove(AttriSheet)
		else:
       		# create an element tree object from the root element.
			SubAttriSheet = ElementTree.SubElement(root, tempSheetname[0])
			tree = ElementTree.ElementTree(root) 
			tree.write(sheetname + '.xml', "utf8")

if __name__ == "__main__":

		rdsname = os.environ.get('PWD')
		sheetname = "ClientAdmin-fsdb0"
		inputFile = "C:\Users\jeguan\Desktop\hft47_xlsprov_R32_58693.xls"
		inputWorkBook = xlrd.open_workbook(inputFile)

		genXmlScript4FSDB(rdsname, sheetname, inputWorkBook)


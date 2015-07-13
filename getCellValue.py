#-*- coding: utf- -*-
#!/usr/bin/python

import sys
import xlrd

def open_excel(file='file.xls'):
	try:
		wb = xlrd.open_workbook(file)
		return wb
	except Excetion,e:
		print(str(e))
#######################################################################
#	get the value of special cells.
#   the output is as follows:
#	
#	{'Administrator': 'Alcatel&01'}
#	{'Webportal': 'Administrator2.'}
#######################################################################
def excel_table_byname1(file= 'file.xls',rowNameIndex=0,sheet_name=u'Sheet'):
	wb = open_excel(file)
	ws = wb.sheet_by_name(sheet_name)
	nrows = ws.nrows # rows of the sheet

	rowValues = ws.row_values(rowNameIndex) # the value of the row.
	print(rowValues)
	#for i in range(0,len(rowValues)):
	
	Dict = {}
	for rownum in range(1,nrows):
		if str(ws.cell(rownum,1).value).strip() == '':
			continue

		Dict[str(ws.cell(rownum,1).value).strip()] = str(ws.cell(rownum,2).value).strip()

	return Dict 
##################################################################################
#	Output for list1
#	[{'Administrator': 'Alcatel&01'}, {'Webportal': 'Administrator2.'}]
##################################################################################
def excel_table_byname2(file= 'file.xls',rowNameIndex=0,sheet_name=u'Sheet'):
	wb = open_excel(file)
	ws = wb.sheet_by_name(sheet_name)
	nrows = ws.nrows # rows of the sheet

	list1 = []
	for rownum in range(1,nrows):
		if str(ws.cell(rownum,1).value).strip() == '':
			continue
		Dict = {}
		Dict[str(ws.cell(rownum,1).value).strip()] = str(ws.cell(rownum,2).value).strip()

		list1.append(Dict)

	return list1

################################################################################
# Function Name:	getLogin
#
# Decsription:		Do "Role" check to find the "ADMINISTRATOR", then check the 
#					"XML_AXCTION" to find "LOGIN" action. 
#					By "LOGIN" and "ADMINISTRATOR",	we can easily locate a cell 
#					for "ClientName" and "Password" separatly.
#
# Inputs:			filename	-	The file contained "ClientName" and "Password".
#									In our scenario, the file is CTSTemplates.xls.
#					sheet_name	-	worksheet name.	
#
# Output:			Login		-	a dict to save "ClientName" and "Password".
################################################################################
def getLogin(filename, sheet_name=u'Sheet'): 

	Login = {}			# to save usrname and passwd.
	titleColOrder = {}	# to save the title column order.	
	nrows = ws.nrows 	# rows of the sheet.
	ncols = ws.ncols 	# column of this sheet.

	wb = open_excel(filename)			# open a excel and return a Book object.
	ws = wb.sheet_by_name(sheet_name)	# worksheet got by input sheet_name.

	# the value of the rowNameIndex row.
	rowValuesList = ws.row_values(0) 

	# This is used to save the position where "XML_ACTION", "ClientName", 
	# "Password" and "Role" are in "ClientAdmin-fsdb0" of CTSTemplates.xls.
	 
	for num_element in range(0, len(rowValuesList)):

		if rowValuesList[num_element] == u'XML_ACTION':
			# Column order for "XML_ACTION"
			titleColOrder[u'XML_ACTION'] = num_element

		elif rowValuesList[num_element] == u'ClientName':
			# Column order for "ClientName"
			titleColOrder[u'ClientName'] = num_element

		elif rowValuesList[num_element] == u'Password':
			# Column order for "Password"
			 titleColOrder[u'Password'] = num_element

		elif rowValuesList[num_element] == u'Role':
			# Column order for "Role"
			titleColOrder[u'Role'] = num_element
		
	print(titleColOrder)

	try:
		# Check the "Role" column to find the "ADMINISTRATOR" role.
		for i in range(0, titleColOrder[u'Role']):
			if ws.cell(i, titleColOrder[u'Role'])).value == u'ADMINISTRATOR':
				if ws.cell(i, titleColOrder[u'XML_ACTION']).value == u'LOGIN':
					ClientName = ws.cell(i,titleColOrder[u'ClientName']).value
					Password = ws.cell(i,titleColOrder[u'Password']).value

					Login[u'ClintName'] = ClientName
					Login[u'Password'] = Password

					return Login	# Only one "ADMISTRATOR" ClientName and 
									# Password is needed. Therefore, if found,
									# just return.
		raise StopIteration
		
	except StopIteration:
		print("\n No 'ADMISTRATOR' Login are found")
		return {}

if __name__ == "__main__":

	filename = r'C:\Users\jeguan\Desktop\CTSTemplate.xls'
#01 Dict
# Test for excel_table_byname1
	#tables = excel_table_byname1(filename, 0, u'ClientAdmin-fsdb0')
	# Output for the following print() is:
	# {'Administrator': 'Alcatel&01', 'Webportal': 'Administrator2.'}
	#print(tables)
	#
	#print(tables.keys())
	#keyList = tables.keys()
	#for i in range(0,len(keyList)):
	#	usrName = keyList[i]
	#	passwd = tables[usrName]
	#	print("usrName = %s" % usrName)
	#	print("passwd = %s" % passwd)
	
#02 List
# Test for excel_table_byname2
	#tables = excel_table_byname2(filename, 0, u'ClientAdmin-fsdb0')

#[{'Administrator': 'Alcatel&01'}, {'Webportal': 'Administrator2.'}]
	#print(tables)

#{'Administrator': 'Alcatel&01'}
#{'Webportal': 'Administrator2.'}
	#for i in range(0, len(tables)):
		#print(tables[i])

#03 "Role" check
	getLogin(filename, u'ClientAdmin-fsdb0')



#-*- coding: utf- -*-
#!/usr/bin/python

import sys
import xlrd
from xml.etree import ElementTree as ET

######################################################################
#	open an excel
######################################################################
def open_excel(inputFile='file.xls'):
	try:
		wb = xlrd.open_workbook(inputFile)
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
def excel_table_byname1(inputFile= 'file.xls',rowNameIndex=0,sheet_name=u'Sheet'):
	wb = open_excel(inputFile)
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
def excel_table_byname2(inputFile= 'file.xls',rowNameIndex=0,sheet_name=u'Sheet'):
	wb = open_excel(inputFile)
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
#	Function:		GenLoginFile
#
#	Description:	Create a tree for Login file. 	
#
#					The generated Login File is:
#
#	<?xml version='1.0' encoding='utf8'?>
#	<Request Action="LOGIN" RequestID="100000">
#		<Authentication>
#			<ClientName>Administrator</ClientName>
#			<Password>Alcatel&01</Password>
#		</Authentication>
#	</Request>
#
#	Input:			Login	-	an input dict, the "ClientName" is used as the Key
#								"Password" is used as the value.
#					Flag	-	 Flag to indicate to get FSDB or GLS Login File.
#
#									1 - get FSDB Login File
#									2 - get GLS Login File
#
#	Output:			NONE				
#			
################################################################################
def GenLoginFile(Login, Flag):

	# Create a root for the tree.
	root = ET.Element("Request", {"Action": "LOGIN", "RequestID": "100000"})

	# Create a sub tree. authSubAttrib is the subroot for usrSubAttrib and 
	# passwdSubAttrib. usrSubAttrib and passwdSubAttrib are on the same level.
	authSubAttrib = ET.Element("Authentication")

	usrSubAttrib = ET.SubElement(authSubAttrib, "ClientName")
	usrSubAttrib.text = Login["ClientName"]

	passwdSubAttrib = ET.SubElement(authSubAttrib, "Password" )
	passwdSubAttrib.text = Login["Password"]

	# "root" is the root for the LoginFile tree.
	root.append(authSubAttrib)
	tree = ET.ElementTree(root)
	
	# Besides FSDB and GLS, if another elements' Login File needs to be got, just
	# add "elif" check and set the value of "LoginName".
	if Flag == 1:
		LoginName = "fsdb"

	elif Flag == 2:
		LoginName = "gls"

	# Save this tree in a file.
	tree.write("LoginFile" + LoginName + ".xml", "utf8")
		

################################################################################
#	Function Name:	getLogin
#
#	Decsription:	Do "Role" check to find the "ADMINISTRATOR", then check the 
# 					"XML_AXCTION" to find "LOGIN" action. 
# 					By "LOGIN" and "ADMINISTRATOR",	we can easily locate a cell 
# 					for "ClientName" and "Password" separatly.
#
#	Inputs:			inputWorkBook	-	The file contained "ClientName" and "Password".
# 										In our scenario, the file is CTSTemplates.xls.
#					sheet_name		-	worksheet name.	
#
#	Output:			Login			-	a dict to save "ClientName" and "Password".
#										if not found, then just return {}
################################################################################
def getLogin(inputWorkBook, sheet_name='Sheet'): 

	Login = {}			# to save usrname and passwd.
	titleColOrder = {}	# to save the title column order.	

	wb = open_excel(inputWorkBook)			# open a excel and return a Book object.
	ws = wb.sheet_by_name(sheet_name)		# worksheet got by input sheet_name.

	nrows = ws.nrows 	# rows of the sheet.
	ncols = ws.ncols 	# column of this sheet.

	# the value of the rowNameIndex row.
	rowValuesList = ws.row_values(0) 

	# This is used to save the position where "XML_ACTION", "ClientName", 
	# "Password" and "Role" are in "ClientAdmin-fsdb0" of CTSTemplates.xls.
	 
	for num_element in range(0, len(rowValuesList)):

		if rowValuesList[num_element] == 'XML_ACTION':
			# Column order for "XML_ACTION"
			titleColOrder['XML_ACTION'] = num_element

		elif rowValuesList[num_element] == 'ClientName':
			# Column order for "ClientName"
			titleColOrder['ClientName'] = num_element

		elif rowValuesList[num_element] == 'Password':
			# Column order for "Password"
			 titleColOrder['Password'] = num_element

		elif rowValuesList[num_element] == 'Role':
			# Column order for "Role"
			titleColOrder['Role'] = num_element
		
	print(titleColOrder)

	try:
		# Check the "Role" column to find the "ADMINISTRATOR" role.
		for i in range(0, titleColOrder['Role']):
			if ws.cell(i, titleColOrder['Role']).value == 'ADMINISTRATOR':
				if ws.cell(i, titleColOrder['XML_ACTION']).value == 'LOGIN':
					ClientName = str(ws.cell(i,titleColOrder['ClientName']).value).strip()
					Password = str(ws.cell(i,titleColOrder['Password']).value).strip()

					Login['ClientName'] = ClientName
					Login['Password'] = Password
					
					# Generate the Login file.
					GenLoginFile(Login, 1)

					return Login	# Only one "ADMISTRATOR" ClientName and 
									# Password is needed. Therefore, if found,
									# just return.
		raise StopIteration
		
	except StopIteration:
		print("\n No 'ADMISTRATOR' Login are found or XML_ACTION is not 'LOGIN'")
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
	getLogin(filename, 'ClientAdmin-fsdb0')



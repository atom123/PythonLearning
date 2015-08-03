#!/usr/bin/env python
#-*- coding:utf-8 -*-
#******This tool is used to retrive the xlsprov template from lab configuration.******

import os
import sys
sys.path.append('/opt/LU3P/lib/python2.6/site-packages')
import re
import xlwt
import xlrd
import getopt
import itertools
import subprocess

from time import gmtime, strftime
from xlutils.copy import copy
from xlutils.styles import Styles

from xml.etree import ElementTree
from xlrd import open_workbook
from os.path import join

DumpDir = '/export/home/lss/logs/xlsRead-dbdump'
ReadLogFile = "/export/home/lss/logs/xlsRead-dbdump/xlsRead.log"

PwdPath = os.environ.get('PWD')
ReqPath = PwdPath +'/reqdir'
ResPath = PwdPath +'/resdir'

ReadList = ["FeatureLicenseParameters", "CapacityLicenseKey", "GlobalParameters", "NGSSParameters", 
             "IMSDeviceServer", "NGSSSIPia", "DiamPort", "DiamAppIDParameters", "GatewayH248",
             "WiFiIMSParameters", "TmsiVlrParameters", "ANSI41MAPParameters", "H248DeviceServer",
             "FS5000SIPia", "H248Port", "FS5000DeviceServer", "LocalCCFConfiguration", "IMSDigitTable",
             "FS5000Translation", "SCGCapacityLicenseKey", "SCGFeatureLicenseParameters", "SS7Timer",
             "SS7DeviceServer", "SS7GTI", "SS7M3uaAs", "SS7M3uaAsp", "SS7M3uaAspSg", "SS7M3uaDpc",
             "SS7M3uaParameters", "SS7M3uaSgAtca", "SS7SccpEntity", "SS7Stack", "SS7Subsystem",
             "SS7TranslationGroup", "SccpParameters", "SS7SccpEntitySet"]

XmlDict = {'DiamMultDestProfile' : 'DiameterMultipleDestinationsProfileTable',
            'SipAggrTrustedRateLimitSet' : 'SipAggregateTrustedRateLimitSet',
            'SipAggrUntrustedRateLimitSet' : 'SipAggregateUntrustedRateLimitSet',
            'SupplementaryServiceInformation' : 'SupplementaryServiceInformationandServiceIdentityMappingTable' ,
            'ImsFeatureTagDrivenDomainSelTab' : 'ImsFeatureTagDrivenDomainSelTable' ,
            'AccessLocationtoESRNMapping' : 'AccessLocationtoESRNMappingTable' ,
            'BorderGatewayPublishedRealm' : 'BorderGatewayPublishedRealmTable' ,
            'MrfAnnouncementInterProf' : 'MrfAnnouncementInterfaceProfile'}

XlsDict = {'DiameterMultipleDestinationsProfileTable' : 'DiamMultDestProfile' ,
            'SipAggregateTrustedRateLimitSet' : 'SipAggrTrustedRateLimitSet',
            'SipAggregateUntrustedRateLimitSet' : 'SipAggrUntrustedRateLimitSet',
            'SupplementaryServiceInformationandServiceIdentityMappingTable' : 'SupplementaryServiceInformation' ,
            'ImsFeatureTagDrivenDomainSelTable' : 'ImsFeatureTagDrivenDomainSelTab',
            'AccessLocationtoESRNMappingTable' : 'AccessLocationtoESRNMapping' ,
            'BorderGatewayPublishedRealmTable' : 'BorderGatewayPublishedRealm' ,
            'MrfAnnouncementInterfaceProfile' : 'MrfAnnouncementInterProf'}

#this list contains the tables which UPDATE only.
UpdateList = ["GlobalParameters", "NGSSParameters", "DiamAppIDParameters", "LocalCCFConfiguration",
              "SS7M3uaParameters", "SipErrorTreatmentTable"]

#dynamic table list. For those tables, there are dynamic child tables which can be configured by users,
#the format of xml response of them are not fixed. So, handle them specially.
DynamicTableList = ["SDPMediaSubsPolicyTable","IcsiTable","SipLinkSetTable","SDPProfileTable",
                    "MrfAnnouncementInterProf","SCSCFProfileTable","ScscfAsAffiliationTable"]

#for table in this list, need to mark default record to UPDATE and mark the new record to CREATE.
DefulRecordList = ["SCTPProfileTable", "SCTPConnectionManagementProfile","SIPStackTimerProfile",
                   "IMSACRChargingProfileTable","HomeNetworkIdentifierTable", "AudioCodec", "VideoCodec"]

#tables which have muilt records when it was created.
CreatWithDefauRecordList = ["BGCFPaniAccessTypeTable", "BGCFPaniAccessInfoTable", "FephFlowPolicyTable",
                            "BGCFDirectAssistTable", "SipAggrUntrustedRateLimitSet", "FephAggPacketPolicyTable",
                            "OnlineChargingProfileTable", "OnlineChargingTriggerData", "FephRemoteIpPolicyTable",
                            "SipTrustedRateLimitSet", "SipUntrustedRateLimitSet", "SipAggrTrustedRateLimitSet"]

#set the font style for the Columns and the Rows for the output Excel file 
myTAGstyle = xlwt.easyxf('pattern: pattern solid, fore_colour pale_blue; font: name Palatino Linotype, bold on;')
myDATAstyle = xlwt.easyxf('pattern: pattern solid, fore_colour light_yellow; font: name Palatino Linotype;')


###############################################################
#
# Function Name: GetIP(Flag)
#
# Description:   define one function to get the CONFIG, FSDB or
#				 GLS IP address of this server.					
#
# Input Value:   Flag	---	Flag to confirm whether it is CONFIG,
#                           FSDB or GLS.
#                               0 - CONFIG(defult value)
#                               1 - FSDB
#                               2 - GLS
#
#                           Default to get the CONFIG's IP.
#
# Return Value: aimed_IP --- The IP that we get per "Flag". 
#
###############################################################
def GetIP(Flag = 0):
    ''' Module to get IP per the input Flag. The default value
        is 0, which means that the CONFIG's IP will be got.'''

    # Get CONFIG's IP
    if Flag == 0:
        aimed_Grep = "cnfg"
        aimed_Print = "CONFG IP: "

    # Get FSDB's IP
    elif Flag == 1:	
        aimed_Grep = "fsdb"
        aimed_Print = "FSDB IP: "

        # Get GLS's IP
    elif Flag == 2:	
        aimed_Grep = "gls"
        aimed_Print = "GLS IP: "

        # Reverved for extension. If another IP need to be got,
        # then, just add the "elif" branch and set the value for 
        # "aimed_Grep" and "aimed_Print".

    else:
        PrintAndSaveLog('\nInvalid input flag\n')
        return ("")

        # The aimed string starts with "aimed_Grep" and it 
        # contains "-g0" and "floating", we only get the 6th 
        # part divided by ";".
    aimed_IPCmd = 'grep ^' + aimed_Grep             \
		+ ' /var/opt/lib/sysconf/service_ip.data'   \
		+ ' | grep "\-g0" | grep floating'          \
		+ ' | cut -d ";" -f 6'

    aimed_Line= os.popen(aimed_IPCmd)
    aimed_IP = aimed_Line.readline().strip('\n')

    AIMEDIP = aimed_Print + aimed_IP 
    PrintAndSaveLog(AIMEDIP)

    return aimed_IP


################################################################################
#	Function:		GenXml4FsdbGls	
#
#	Description:	Generate the xml request for FSDB and GLS.  	
#
#	Input:			destFile  - the path to save the xml request.
#					sheetname - name of the sheet, from which the data will be 
#								retrieved to generate the destFile.
#					LoginName - the "ClientName" and "Password" are saved in 
#								this Dict. 
#
#	Output:			destFile.xml - this file is used to save the generated
#								   request.
################################################################################
def GenXml4FsdbGls(destFile, sheetname, LoginName):
    ''' Generate the xml for GLS and/or FSDB. Special treatment is needed for
		FSDB because there is "-fsdb\d" at the end of each tab benath the FSDB.
		When generating the xml request file, no "-fsdb\d" postfix should exit.
	'''
    root = ElementTree.Element("Request", {"Action": "READ", "RequestId": "100000"})

    # remove the postfix "-fsdb\d" or "-gls" from the sheetname
    reCmd = re.compile(r'.+(?=-\w)')
    tempSName = reCmd.findall(sheetname)

    Attrib = ElementTree.Element(tempSName[0])

    if tempSName[0] == "ClientAdmin":
        SubAttrib = ElementTree.SubElement(Attrib, "ClientName")
        SubAttrib.text = LoginName["ClientName"]

    # add the elemnet subelement to the end of this elements internal list of 
    # subelements.
    root.append(Attrib)
    tree = ElementTree.ElementTree(root)

    tree.write(destFile + '.xml', 'utf8')


################################################################################
#	Function:		getProvRes4FSDB	
#
#	Description:	get the provisioned response fromt the "OutFile".
#
#	Input:			OutFile  - path of the OutFile.
#					provRes4FSDBFile - file to save the retrieved response.
#					ResPath - path to save the file.
#
#	Output:			NONE
#			
################################################################################
def getProvRes4FSDB(OutFile, provRes4FSDBFile):
    '''get the provisioned value from the "OutFile" generated by xml2fsdbgls.py '''

    new_file =  provRes4FSDBFile
    newfile = open(new_file, 'wb')

    open_file = open(OutFile, 'r')
    read_file = open_file.readlines()

    # re.S means: Make the '.' special character match any character at all,
    # including a newline; without this flag, '.' will match anything except a newline.
    # '(.+?)' means: this is a lazzy match. When the fist 'Response>' is found, then
    # it will not try to match the next 'Response>'
    re_patt = re.compile(r'<Response Status="OKAY" CongLvl="LEVEL0"*(.+?)Response>', re.S)

    str1 = ""
    # save the readout lines into str1.
    for line in read_file:
        str1 = str1 + line

    result = re_patt.search(str1)
    if result != None:
        newfile.write(result.group(0))



################################################################################
#	Function:		Gen_Login_Logoff_File
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
#									1 - get FSDB Login File
#									2 - get GLS Login File
#
#					genLoginFile - 1 - generate a Login file (default)
#								   0 - generate a Logoff file
#
#					ActValue	 - Flag to indicate to generate Login or Logoff file.
#									 LOGIN - Action for LOGIN (default)
#									 LOGOFF - Action for LOGOFF
#
#	Output:			savedName	- file name for the Login or Logoff	
#			
################################################################################
def Gen_Login_Logoff_File(Login, Flag, genLoginFile=1, ActValue="LOGIN"):

    # Init the Logfile parameter, this is used when save a LOGOFF file.
    # For LOGIN file, this parameter is set to "LoginFile".
    LogFile = "LogoffFile"

    if genLoginFile == 0:
        ActValue = "LOGOFF"

    # Create a root for the tree.
    root = ElementTree.Element("Request", {"Action": ActValue, "RequestId": "100000"})

    # Create a sub tree. authSubAttrib is the subroot for usrSubAttrib and 
    # passwdSubAttrib. usrSubAttrib and passwdSubAttrib are on the same level.
    authSubAttrib = ElementTree.Element("Authentication")

    usrSubAttrib = ElementTree.SubElement(authSubAttrib, "ClientName")
    usrSubAttrib.text = Login["ClientName"]

    # For Logoff, there is no need to add the password.
    if genLoginFile:
        passwdSubAttrib = ElementTree.SubElement(authSubAttrib, "Password" )
        passwdSubAttrib.text = Login["Password"]
        LogFile = "LoginFile"

    # "root" is the root for the LoginFile tree.
    root.append(authSubAttrib)
    tree = ElementTree.ElementTree(root)

    # Besides FSDB and GLS, if another elements' Login File needs to be got, just
    # add "elif" check and set the value of "LoginName".
    if Flag == 1:
        LogName = "fsdb"

    elif Flag == 2:
        LogName = "gls"

    savedName = LogFile + LogName + ".xml"

    # Save this tree in a file.
    tree.write(savedName, "utf8")

    
    return savedName


################################################################################
#	Function Name:	getLoginLogoff
#
#	Decsription:	Do "Role" check to find the "ADMINISTRATOR", then check the 
# 					"XML_AXCTION" to find "LOGIN" action. 
# 					By "LOGIN" and "ADMINISTRATOR",	we can easily locate a cell 
# 					for "ClientName" and "Password" separatly.
#
#	Inputs:			inputWorkBook	-	The file contained "ClientName" and "Password".
# 										In our scenario, the file is CTSTemplates.xls.
#					sheet_name		-	worksheet name.	
#					Flag			-	Flag to indicate to get FSDB or GLS Login File.
#											1 - get FSDB Login File
#											2 - get GLS Login File
#					genLoginFile	-	0 - generate the login file (default)
#									-	1 - generate the logoff file
#					ActValue		-	Flag to indicate to generate Login or Logoff file.
#											LOGIN - Action for LOGIN (default)
#											LOGOFF - Action for LOGOFF
#
#	Output:			Login			-	a dict to save "ClientName" and "Password".
#										if not found, then just return {}
################################################################################
def getLoginLogoff(inputWorkBook, Flag, sheet_name='Sheet', genLoginFile=1, ActValue="LOGIN"):

    Login = {}			# to save usrname and passwd.
    titleColOrder = {}	# to save the title column order.	

    wb = inputWorkBook
    ws = wb.sheet_by_name(sheet_name)		# worksheet got by input sheet_name.

	nrows = ws.nrows	# the number of rows in this sheet.

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

    try:
        # Check the "Role" column to find the "ADMINISTRATOR" role.
        # First, loop through the "Role" column to find "ADMINISTRATOR",
        # if "ADMINISTRATOR" is found, then we can get its row in this worksheet.
        # Second, check the value of "XML_ACTION" to determin wheter its value is
        # "LOGIN".
        # Last, after the first step and second step, we can find the aimed row,
        # then with columns of "ClientName" and "Password" we can finally get
        # their value. No need to start with i = 0 for the reason the first row
		# is used to save the title name. range(1, nrows) means: 1, 2, ..., nrows-1
        for i in range(1, nrows):
            if ws.cell(i, titleColOrder['Role']).value == 'ADMINISTRATOR':
                if ws.cell(i, titleColOrder['XML_ACTION']).value == 'LOGIN':
                    ClientName = str(ws.cell(i,titleColOrder['ClientName']).value).strip()
                    Password = str(ws.cell(i,titleColOrder['Password']).value).strip()

                    Login['ClientName'] = ClientName
                    Login['Password'] = Password

                    # Generate the Login file.
                    Gen_Login_Logoff_File(Login, Flag)

                    #Generate the Logoff file.
                    Gen_Login_Logoff_File(Login, Flag, 0, "LOGOFF")

                    return Login	# Only one "ADMISTRATOR" ClientName and 
                                    # Password is needed. Therefore, if found,
                                    # just return.
        raise StopIteration

    except StopIteration:

        print("\n No 'ADMISTRATOR' Login are found or XML_ACTION is not 'LOGIN'")
        Login['ClientName'] = ""	# not found just keep Login[] to be NULL.
        Login['Password'] = "" 

        # Generate the Login file.
        Gen_Login_Logoff_File(Login, Flag)
        #Generate the Logoff file.
        Gen_Login_Logoff_File(Login, Flag, 0, "LOGOFF")

        return Login



########################################################################
# Function Name: PrintAndSaveLog(Log_Message)
# Description:   define one function to print and save log.
# Input Value:   Log_Message --- the log message that need to be printed.
# Return Value:  NULL
########################################################################
def PrintAndSaveLog(Log_Message):
        ''' Module to send log messages into /export/home/lss/logs/xlsprov.'''
        global ReadLogFile
        global DumpDir

        if (not os.path.exists(DumpDir)):
                os.mkdir(DumpDir)

        try:
                LogDir = os.path.dirname(DumpDir)
                if os.path.exists(LogDir) == False:
                    os.mkdir(LogDir)

                LogFile = open(ReadLogFile,"a")
                LogDateTime = strftime("%h %d %T ",gmtime())

                # write the log to log file.
                LogFile.write(LogDateTime)
                LogFile.write(' ')
                LogFile.write(Log_Message) 
                LogFile.write('\n')
                LogFile.close()
                print  Log_Message

        except Exception, exc:
                print exc # str(exc) is printed
                raise Exception, 'PrintAndSaveLog() failed!!!'


##############################################################################
# Function Name: SaveSheetNameFrmXls(sname)
#
# Description:   define one function to save sheet names from input Excel file.
#
# Input Value:   sname   --- the sheet name.
#
# Return Value:  NULL
##############################################################################
def SaveSheetNameFrmXls(sname):
        ''' Module to save all sheets name from the input Excel template. '''

        XlsWorkBook = open_workbook(XlsFile, 'rb')
        
        # Save all sheets to the list(sname) from input Excel file except 'Index' 
		# and 'Site Specific Data'.
        for sheet in XlsWorkBook.sheets():
			if sheet.name in ['Index', 'Site Specific Data', 'ENUMTLDTable']:
				continue
			else:
				sname.append(sheet.name)


###################################################################################
# Function Name: GenXmlScript(tabname,sheetname,NodeID)
# Description:   define one function to generate xml script for each sheet.
# Input Value:   tabname   --- the path name which used to store xml script for sheet
#                sheetname --- the sheet name
#                NodeID    --- the node ID for SIPia port
# Return Value:  NULL
###################################################################################
def GenXmlScript(tabname,sheetname,NodeID):
	''' Module to generate the xml script according to the sheetname. '''

	if sheetname in ReadList:
		action = 'READ'
	else:
		action = 'READ_ALL'

	if sheetname in XmlDict:
		sheetname = XmlDict[sheetname]
	else:
		pass

	# Set the root of xml script tree for sheetname
	root = ElementTree.Element('Request', {'Action': action})
	AttriSheet = ElementTree.Element(sheetname)

	if NodeID != 0:
		SubAttriSheet = ElementTree.SubElement(AttriSheet,'NODE')
		SubAttriSheet.text = str(NodeID)

	root.append(AttriSheet)

	# Create an element tree object from the root element.
	tree = ElementTree.ElementTree(root)

	if NodeID == 0:
		tree.write(tabname + '.xml', 'utf8')
	else:
		strnodeid = str(NodeID)
		modtabname = tabname + strnodeid;
		tree.write(modtabname + '.xml', 'utf8')


###################################################################################
# Function Name: GenTagForSheet(TableElement,SheetTagList,SheetTagValueList,SheetName)
# Description:   define one function to generate tag for each sheet.
#                Save the tag name into SheetTagList;
#                Save the value of each tag into SheetTagValueList.
# Input Value:   TableElement      --- the elemet of this table.
#                SheetTagList      --- the tag in the sheet.
#                SheetTagValueList --- the value of the tag.
#                SheetName         --- the sheet name.
# Return Value:  NULL
###################################################################################
def GenTagForSheet(TableElement,SheetTagList,SheetTagValueList,SheetName):
	''' Module to generate tag according to the sheet name. '''
	if (TableElement.getchildren()):
		if (TableElement.tag != SheetName):
			SheetTagList.append('_GRPSTART_'+TableElement.tag)
			SheetTagValueList.append('Required')

		for child in TableElement:
			GenTagForSheet(child,SheetTagList,SheetTagValueList,SheetName)

		if (TableElement.tag != SheetName):
			SheetTagList.append('_GRPEND_'+TableElement.tag)
			SheetTagValueList.append('Required')
	else:
		SheetTagList.append(TableElement.tag)
		# Make sure to use str on TableElement.text only for Non-Empty Strings
		if(str(TableElement.text) != "None"):
			TableElement.text = NormalizeChar(str(TableElement.text))
		SheetTagValueList.append(TableElement.text)



###############################################################
# Function Name: NormalizeChar(NormStr)
# Description:   define one function to make sure the special 
#                characters have the right Xml format in the 
#                Xls file
# Input Value:   the input charater get from xml
# Return Value:  the output charater which need to be added in 
#                xls file.
###############################################################
def NormalizeChar(NormStr):
        ''' Module to handle special characters in the xml.'''
        NormStr = NormStr.replace('&','&amp;')
        NormStr = NormStr.replace('<','&lt;')
        NormStr = NormStr.replace('>','&gt;')
        NormStr = NormStr.replace('"','$quot;')
        NormStr = NormStr.replace('\'', '&apos;')

        return NormStr



###################################################################################
# Function Name: GenXmlForSIPiaPort(rdsname,sheetname)
# Description:   define one function to generate xml script for SIPia port.
# Input Value:   rdsname --- the path name which used to store xml script for sheet
#                sheetname --- the sheet name.
# Return Value:  NULL
###################################################################################
def GenXmlForSIPiaPort(rdsname,sheetname):
	''' Module to generate xml script for sheet name is SIPia port. '''

	global SIPiaOk

	DBTable = ''
	OutDBTable = ''

	# get node ids via dbdump
	if (sheetname == "NGSSSIPia"):
		DBTable = 'cfg.ngss_sipia_port'
	else:
		DBTable = 'appcnfg.ngss_diam_port'

	OutDBTable = DBTable + '.xdat'

	# check if dumpdir is exist.
	if (not os.path.exists(DumpDir)):
		os.mkdir(DumpDir)

	# dbdump db tables
	DBDumpCmd = 'cd %s;dbdump -u lssdba -p lssdba -noversion -table %s' % (DumpDir,DBTable)
	os.system(DBDumpCmd)

	# get unique node ids from db tables
	CatCmd = 'cd %s;cat %s | cut -f1 -d"|" | sort -u > nodes.out' % (DumpDir,OutDBTable)
	os.system(CatCmd)

	# open the node file
	NodeOutFile = DumpDir + '/nodes.out'
	NodeFile = open(NodeOutFile, 'r')

	# final combined output file
	OutPutName = sheetname + '.out'
	if (not os.path.exists(ResPath)):
		os.mkdir(ResPath)
	OutPut = join(ResPath,OutPutName)

	if (os.path.exists(OutPut)):
        	os.remove(OutPut)

	# create a blank main output file in case no records found
	TouchCmd = 'touch %s' % (OutPut)
	os.system(TouchCmd)

	FindNode = 0
	FindOk = 0

	# do sipia port reads for all nodes
	for line in NodeFile:
		NodeID = int(line)
		GenXmlScript(rdsname,sheetname,NodeID)

		ModeInName = rdsname + str(NodeID) + '.xml'
		ModeOutName = sheetname + str(NodeID) + '.out'
		ModOut = join(ResPath,ModeOutName)

		Cmd = 'xml2cfg -h %s -p 9650 -i %s -o %s' % (cnfgip,ModeInName,ModOut)
		os.system(Cmd)

		Okay = 0
		NodeRoot = ElementTree.parse(ModOut).getroot()
		NodeResponse = NodeRoot.getiterator('Response')

		for NumResp in NodeResponse:
			if (NumResp.attrib['Status'] == "OKAY"):
				Okay = 1
				FindOk = 1

		if Okay == 0:
			continue

		# take out what we need from output
		SIPiaNodeOut = open(ModOut, 'r')

		NodeLine = ''
		WroteRecord = 0
		xmltagname = '<%s>' % (sheetname)
		xmltagnameend = '</%s>' % (sheetname)
		xmltagnamecr = '<%s>\n' % (sheetname)
		xmltagnameendcr = '</%s>\n' % (sheetname)

		if FindNode == 0:
			FindNode = 1
			SIPiaOut = open(OutPut, 'a+')
			SIPiaOut.write('<ResponseBatch>\n')
			SIPiaOut.write('<Response Status="OKAY" Action="READ">\n')
			SIPiaOut.write(xmltagnamecr)

		for line2 in SIPiaNodeOut:
			if line2.find('SESSION BEGIN') == -1 and line2.find('<ResponseBatch>') == -1 and line2.find('<Response Status') == -1 and line2.find(xmltagname) == -1 and line2.find(xmltagnameend) == -1 and line2.find('</Response>') == -1 and line2.find('</ResponseBatch>') == -1 and line2.find('SESSION END') == -1:

				# data line
				if line2.find('<NODE>') != -1:
					# this is the NODE_ID line, save it
					NodeLine = '\t%s' % (line2)

				elif line2.find('<Record>') != -1:
					# this is the Record line, write it and set the flag that we wrote it
					# in order to put the NODE_ID line after it
					SIPiaOut.write(line2)
					WroteRecord = 1

				elif WroteRecord == 1:
					# write the NODE_D line, then this
					# first data line
					SIPiaOut.write(NodeLine)
					SIPiaOut.write(line2)
					WroteRecord = 0

				else:
					# all other data lines
					SIPiaOut.write(line2)

		SIPiaNodeOut.close()

	if FindNode == 1:
		SIPiaOut.write(xmltagnameendcr)
		SIPiaOut.write('</Response>\n')
		SIPiaOut.write('</ResponseBatch>\n')
		SIPiaOut.close()

	if FindOk == 1:
		SIPiaOk = 1


	NodeFile.close()
	cdcmd = 'cd %s' % (PwdPath)
	os.system(cdcmd)


#######################################################################
# Function Name: GetSheetByName(XlsWorkBook, SheetName)
# Description:   define one function to get the sheet by input name
# Input Value:   XlsWorkBook  --- the workbook of the input Excel file.
#                SheetName    --- the sheet name.
# Return Value:  sheet        --- if the sheet exists, return it.
#                None         --- if not, return none.
#######################################################################
def GetSheetByName(XlsWorkBook, SheetName):
	""" Get a sheet by name from xlwt.Workbook
		Returns None if no sheet with the given name is present.
	"""

	try:
		for idx in itertools.count():
			sheet = XlsWorkBook.get_sheet(idx)

			# Check if sheet name with postfix.
			sheet.name = RenameWorkSheet(sheet.name)

			if sheet.name == SheetName:
				return sheet

	except IndexError:
		PrintAndSaveLog("ERROR: No Sheet FOUND")

		return None


##################################################################
# Function Name: OpenInputExcel(InputXlsFile)
# Description:   define one function to open the input Excel file.
# Input Value:   InputXlsFile --- the input Excel file.
# Return Value:  workbook     --- if open Excel file success, 
#                                 return the workbook of the 
#                                 input Excel file.
#                sys.exit(1)  --- if fail, exit and print error.
#################################################################
def OpenInputExcel(InputXlsFile):
        ''' Module to open the input Exvel file template. '''
        xlsfile = InputXlsFile
        try:
            workbook = open_workbook(xlsfile,formatting_info=True ,on_demand = True)
            return workbook
        except:
            LogMsg = 'An error was encountered when opening the input file: ' + (InputXlsFile)
            PrintAndSaveLog(LogMsg)
            PrintAndSaveLog('\nThe given input file maybe in an invalid format.')
            PrintAndSaveLog('\nPlease make sure the input file is an Excel file in a binary format.\n')
            sys.exit(1)


###############################################################
# Function Name: GetCnfigIP()
# Description:   define one function to get the config IP 
#                address of this server.
# Input Value:   NULL
# Return Value:  NULL
###############################################################
def GetCnfigIP():
        ''' Module to get the config IP address of this server. '''
        global cnfgip
        cnfgIPCmd = 'grep ^cnfg /var/opt/lib/sysconf/service_ip.data | grep "\-g0" | grep floating | cut -d";" -f6'
        cnfgipLine = os.popen(cnfgIPCmd)
        cnfgip = cnfgipLine.readline().strip('\n')

        CNFGIP = "CNFG IP: " + cnfgip
        PrintAndSaveLog(CNFGIP)




##########################################################################
# Function Name: RenameWorkSheet(InputSheetName)
# Description:   define one function to rename the work sheet.
#                For some input xls file template, the worksheet name maybe
#                contains '-1' or '-2' postfix. According to them, it can't
#                get value. So, it needs to remove the postfix.
# Input Value:   InputSheetName  --- the sheet name got from input xls file.
# Return Value:  OutputSheetName --- the updated sheet name for output xls.
##########################################################################
def RenameWorkSheet(InputSheetName):
        ''' Model to rename the input work sheet name.'''

        SheetNameList = []

        SheetWithPostfix = re.compile(r'.+(?=-\d)')
        SheetNameList = SheetWithPostfix.findall(InputSheetName)

        if (len(SheetNameList) > 0):
                OutputSheetName = SheetNameList[0]
                return OutputSheetName
        else:
                return InputSheetName




##########################################################################
# Function Name: ReadSheetAndWrite(readOutputWB, inputWorkBook, tmpsname) 
#
# Description:   define one function to do xml read by sheet 
#                and write the value to output Excel file.
#
# Input Value:   readOutputWB  --- the workbook of output Excel file.
#                inputWorkBook --- the workbook of input Excel file.
#                tmpsname      --- the sheet name list of input Excel file.
#
# Return Value:  NULL
##########################################################################
def ReadSheetAndWrite(readOutputWB, inputWorkBook, tmpsname):
        ''' Module to do xml read by sheet
        	and write the value to output Excel file.
        '''    
		global SIPiaOK
		global Login_dict

		Login_dict = {}

        # Before do xml read, clean duplicate sheet name with postfix.
        NumSheet = len(tmpsname)
        for sn in range(0, NumSheet):
                tmpsname[sn] = RenameWorkSheet(tmpsname[sn])

        LoopStart = 0
        LoopNext = 1

		# Remove the same list name from the name list. 
        while LoopStart < NumSheet:
                while LoopNext < NumSheet:
                        if (tmpsname[LoopStart] == tmpsname[LoopNext]):
                                del tmpsname[LoopNext]
                                NumSheet = NumSheet - 1
                                LoopNext = LoopStart + 1
                        else:
                                LoopNext = LoopNext + 1

                LoopStart = LoopStart + 1
                LoopNext = LoopStart + 1

        # Loop the sheet name list to do xml read and write.
        for i in range(0, len(tmpsname)):
		
            # Check if request path exists.
            if (not os.path.exists(ReqPath)):
                os.mkdir(ReqPath)

            # Check if response path exists.
            if (not os.path.exists(ResPath)):
                os.mkdir(ResPath)

            rdsname = join(ReqPath, tmpsname[i])
            InName = rdsname + '.xml'

            OutName = tmpsname[i] + '.out'
            OutFile = join(ResPath, OutName)

            #Carl            
            FsdbGlsFlag = GetFsdbGlsFlag(tmpsname[i])	

			# The tmpsname[i] is for GLS or FSDB now if FsdbGlsFlag != 0.
            if FsdbGlsFlag != 0:	
                FsdbGlsIP = GetIP(FsdbGlsFlag)
			
				# "1" is for FSDB and "2" is for GLS.
                if FsdbGlsFlag == 1:	
                    FsdbGlsPort = 7856

					# Used to save Login and Logoff username and password for FSDB.
                    LoginFile = "LoginFilefsdb.xml"	
                    LogoffFile = "LogoffFilefsdb.xml"

                elif FsdbGlsFlag == 2:	
                    FsdbGlsPort = 6856

					# Used to save Login and Logoff username and password for GLS.
                    LoginFile = "LoginFilegls.xml"
                    LogoffFile = "LogoffFilegls.xml"

				# If tmpsname[i] = ClientAdmin-fsdb0, then after reCmd.findall(),
				# tempSName = ClientAdmin.
                reCmd = re.compile(r'.+(?=-\w)')
                tempSName = reCmd.findall(tmpsname[i])

                # The ClientName and Password can only need to retrieve once. 
				# For FSDB and GLS, "ClientName" and "Password" are needed to be 
				# retrieved firstly for the reason that they are required while 
				# login/logoff.
				#
				# NOTE 1: for FSDB, "ClientName" and "Password" should not be deleted 
				#		  until the rest of the sheets beneath the FSDB item retrieved. 	
				# NOTE 2: If there is no "ClientAdmin" is found, then, just raise an
				#		  exception and keep the tool continue.
                if tempSName[0] == "ClientAdmin":
                    Login_dict= getLoginLogoff(inputWorkBook, FsdbGlsFlag, tmpsname[i])
                    Logoff_dict = getLoginLogoff(inputWorkBook, FsdbGlsFlag, tmpsname[i],
								  genLoginFile=0, ActValue="LOGOFF")

                GenXml4FsdbGls(rdsname, tmpsname[i], Login_dict)

                cmd = 'python xml2fsdbgls.py %s %s %s %s %s %s' % (FsdbGlsIP,
					   FsdbGlsPort, LoginFile, LogoffFile, InName, OutFile)

                os.system(cmd)

				# To save the generated responses, which include the login response,
				# logoff response and readout response.
                tempOutName = tmpsname[i] + '_bk' + '.out'
				tempOutFile = join(ResPath, tempOutName)
				cmd = 'cp %s %s' % (OutFile, tempOutFile) 
				os.system(cmd)

				# Considering the generated file, it is not a standard xml file, therefore, 
				# need to get the provisioned "Response" data from the file generated by 
				# os.system(cmd). In the ".out" file, there is no need to save the "LOGIN"
				# or "LOGOFF" "Response".
                getProvRes4FSDB(tempOutFile, OutFile)

                continue

            if tmpsname[i] in ['NGSSSIPia', 'H248Port', 'FS5000SIPia', 'DiammPort']:
                SIPiaOk = 0
                GenXmlForSIPiaPort(rdsname, tmpsname[i])
            else:
                SIPiaOk = 1
                GenXmlScript(rdsname, tmpsname[i],0)
                cmd = 'xml2cfg -h %s -p 9650 -i %s -o %s' % (cnfgip, InName, OutFile)
                os.system(cmd)

            if tmpsname[i] in XmlDict:
                sname = XmlDict[tmpsname[i]]
            else:
                sname = tmpsname[i]

            if SIPiaOk == 0:
                LogMsg = 'No Records found in the System for Table:' + sname
                PrintAndSaveLog(LogMsg)
                PrintAndSaveLog('Info from the Input Template will be retained')
                continue

            root = ElementTree.parse(OutFile).getroot()
            ListNode = root.getiterator('Response')

            # Logic is to read the sheets from the Copied excel workbook - readOutputWB and write in them
            # For tables that have records in the Xml output we will clear the sheet and write the xml
            # For tables that don't have data in the Xml output we will just keep the copied sheet as-is

            for node in ListNode:
                if (node.attrib['Status'] == "OKAY"):

                        if tmpsname[i] in XmlDict:
                                WorkSheet = readOutputWB.add_sheet(tmpsname[i], cell_overwrite_ok=True)
                        else:
                                WorkSheet = readOutputWB.add_sheet(sname, cell_overwrite_ok=True)

                        tables = root.findall('Response/'+sname)

                        SheetTagList = []
                        SheetTagValueList = []

                        SheetTagList.append('XML_ACTION')
                        if tmpsname[i] in UpdateList:
                                SheetTagValueList.append('UPDATE')
                        else:
                                SheetTagValueList.append('CREATE')

                        for table in tables:
                                GenTagForSheet(table, SheetTagList, SheetTagValueList, sname)

                        if sname in XlsDict:
                                sname = XlsDict[sname]

                        # Generate the tag header for sheet in Output Excel file.
                        if sname not in DynamicTableList:
                                GenTagAndWrite(readOutputWB, WorkSheet, SheetTagList, 
											   SheetTagValueList, sname)
                        else:
                                GenTagAndWriForDynTab(readOutputWB, WorkSheet, SheetTagList, 
													  SheetTagValueList, sname)

                else:

                        LogMsg = 'No Records found in the System for Table: ' + sname
                        PrintAndSaveLog(LogMsg)
                        LogMsg = 'Info from the Input Template will be retained.'
                        PrintAndSaveLog(LogMsg)


##############################################################################################
# Function Name: GenTagAndWrite(readOutputWB, WorkSheet, SheetTagList, sname)
# Description:   define one function to generate the tag header for sheet in Output Excel file.
#                Firstly, generate the TAG header from the longest Record.
#                Then, check if the TAG header has more than 256 TAGs.
#                if yes, then we need to split the xml output into multiple TABs.
#                Then, write the values from Row 1 through the last record.
# Input Value:   readOutputWB      --- the workbook of output Excel file.
#                WorkSheet         --- the sheet name from input Excel file.
#                SheetTagList      --- the tag list of sheet.
#                SheetTagValueList --- the value list of this sheet.
#                sname         --- the sheet name.
# Return Value:  NULL
##############################################################################################
def GenTagAndWrite(readOutputWB, WorkSheet, SheetTagList, SheetTagValueList, sname):
        ''' Module to generate the tag header for sheet in Output Excel file. '''

        global numTagsinRec
        global startIndex 
        global endIndex
    
        # Generate the TAG header from the longest Record
        GenTagFrmLongRecord(SheetTagList)

        # For Tables with records get the TAGs from the longest Record
        # For Parameters use the original TAG as is
        finalTag = []
        firstRecstartIndex = 0
        if(numTagsinRec == 0):
            finalTag = SheetTagList 
        else:
            finalTag = SheetTagList[startIndex:endIndex+1]
            #Need to add the TAGS from XML_ACTION to the first _GRPSTART_Record
            if sname in ['IMSDeviceServer','FS5000DeviceServer']:
                    firstRecstartIndex = 1
            else:
                    firstRecstartIndex = SheetTagList.index('_GRPSTART_Record')
    
            for i in range(0, firstRecstartIndex):
                    finalTag.insert(i,SheetTagList[i])

        #For 'SipFilter', 'CRFProgrammableRule', remove useless tag.
        RemoveStart = 0
        RemoveEnd = 0
        FinalTagLen = len(finalTag)
        for i in range(0, FinalTagLen):
                if sname in ['SipFilter', 'CRFProgrammableRule']:
                        if finalTag[i] in ['_GRPSTART_HBR_RULE_EXECUTION_ORDER', '_GRPSTART_HBR_PATTERN_EXECUTION_ORDER']:
                                RemoveStart = i
                                for j in range(i, FinalTagLen):
                                        if finalTag[j] in ['_GRPEND_HBR_RULE_EXECUTION_ORDER','_GRPEND_HBR_PATTERN_EXECUTION_ORDER']:
                                                RemoveEnd = j + 1
                                while RemoveStart < RemoveEnd:
                                        finalTag.remove(finalTag[RemoveStart])
                                        RemoveEnd = RemoveEnd - 1
                                break
                elif sname in ['SipFilterHeaderRule', 'CRFHeaderPatternDefinition']:
                        if finalTag[i] in ['PARAMETER_EXECUTION_ORDER', 'PARAMETER_RULE_EXECUTION_ORDER']:
                                finalTag.remove(finalTag[i])
                                break 

        # check if the TAG header has more than 256 TAGs
        # if yes, then we need to split the xml output into multiple TABs and write the value for it.
        # if no, just write the value for this tag.
        CheckAndWriteTag(readOutputWB, WorkSheet, finalTag, sname, SheetTagList, SheetTagValueList)
     


#################################################################################
# Function Name: GenTagFrmLongRecord(SheetTagList)
# Description:   define one function to generate the tag from the longest record.
#                Loop each record in table's xml response, fing the longest one,
#                use the tag of the longest record as the tag name in output xls.
# Input Value:   SheetTagList --- the tag list of the sheet
# Return Value:  NULL
#################################################################################
def GenTagFrmLongRecord(SheetTagList):
        ''' Module to generate the tag header for sheet from the longest record. '''

        global numTagsinRec
        global startIndex
        global endIndex

        numTagsinRec = 0
        numTagsinRecMax = 0
        startIndex = 0
        endIndex = 0

        #Generate the TAG header from the longest Record
        #Among all the Records find the Record with the most subelements
        #startIndex and endIndex will provide the length of the longest Record
        for i in range(0,len(SheetTagList)):
                if SheetTagList[i] in ["_GRPSTART_Record","_GRPSTART_DNSConfiguration"]:
                        for j in range(i+1,len(SheetTagList)):
                                if SheetTagList[j] in ["_GRPEND_Record","_GRPEND_ComponentParameters"]:
                                        numTagsinRec = j - i
                                        if(numTagsinRecMax ==0):
                                                numTagsinRecMax = numTagsinRec
                                                startIndex = i
                                                endIndex = j
                                        elif(numTagsinRec > numTagsinRecMax):
                                                numTagsinRecMax = numTagsinRec
                                                startIndex = i
                                                endIndex = j
                                        break 



####################################################################################
# Function Name: GenTagAndWriForDynTab(readOutputWB,WorkSheet,
#                          SheetTagList,SheetTagValueList,sname)
# Description:   define one function to write value for table in dynamic table list.
#                For tables in dynamic table list, will create worksheet for each
#                record in them. Because the xml for each record is different.
# Input Value:   readOutputWB      --- the workbook of output Excel file.
#                WorkSheet         --- the sheet name from input Excel file.
#                SheetTagList      --- the tag list of sheet.
#                SheetTagValueList --- the value list of this sheet.
#                sname             --- the sheet name.
# Return Value:  NULL
####################################################################################
def GenTagAndWriForDynTab(readOutputWB,WorkSheet,SheetTagList,SheetTagValueList,sname):
        ''' Model to write value for table in dynamic table list.'''
       
        RecordNum = 0
        #loop the xml response to find the num of record.
        for i in range(0,len(SheetTagList)):
                if(SheetTagList[i] == '_GRPSTART_Record'):
                        RecordNum = RecordNum + 1
        
        #add sheet for each record for those special table
        NameList = []
        if RecordNum > 1:
                WorkSheet.name = sname + "-1"
                for j in range(1,RecordNum):
                        NewSheet = sname + "-" + str(j+1)
                        NameList.append(NewSheet)
       
        #write value for first record.
        NumCol = 0
        FirstStart = 0
        NextStart = 0
        NumRow = 1
        for slen in range(0,len(SheetTagList)):
                if(SheetTagList[slen] != '_GRPEND_Record'):
                        if(NumCol < 256):
                                if(SheetTagList[slen] == '_GRPSTART_Record'):
                                        FirstStart = slen
                                        #For ScscfAsAffiliationTable, when create table, there is default value,
                                        #Need to special handle it.
                                        if (sname == "ScscfAsAffiliationTable") and (NumRow == 1):
                                                NumRow = NumRow + 1
                                                for sl in range(0, FirstStart):
                                                        WorkSheet.write(NumRow,sl,SheetTagValueList[sl],myDATAstyle)
                                                WorkSheet.write(NumRow,0,'UPDATE',myDATAstyle)

                                WorkSheet.write(0,NumCol,SheetTagList[slen],myTAGstyle)
                                WorkSheet.write(NumRow,NumCol,SheetTagValueList[slen],myDATAstyle)
                                NumCol = NumCol + 1
                else:
                        WorkSheet.write(0,NumCol,SheetTagList[slen],myTAGstyle)
                        WorkSheet.write(NumRow,NumCol,SheetTagValueList[slen],myDATAstyle)
                        NextStart = slen + 1
                        NumCol = 0
                        break

        #write value for other records.
        if(len(NameList) > 0):
                for nl in range(0,len(NameList)):
                        NewWorkBook = readOutputWB.add_sheet(NameList[nl],cell_overwrite_ok=True)
                        #write the common tag values for other records.
                        for fl in range(0, FirstStart):
                                NewWorkBook.write(0,fl,SheetTagList[fl],myTAGstyle)
                                NewWorkBook.write(1,fl,SheetTagValueList[fl],myDATAstyle)
                                NumCol = fl + 1

                        for nlen in range(NextStart,len(SheetTagList)):
                                if(SheetTagList[nlen] != '_GRPEND_Record'):
                                        if(NumCol < 256):
                                                NewWorkBook.write(0,NumCol,SheetTagList[nlen],myTAGstyle)
                                                NewWorkBook.write(1,NumCol,SheetTagValueList[nlen],myDATAstyle)
                                                NumCol = NumCol + 1
                                else:
                                        NewWorkBook.write(0,NumCol,SheetTagList[nlen],myTAGstyle)
                                        NewWorkBook.write(1,NumCol,SheetTagValueList[nlen],myDATAstyle)
                                        NextStart = nlen + 1
                                        NumCol = 0
                                        break
                        
        

##################################################################################
# Function Name: CheckAndWriteTag(readOutputWB, WorkSheet, finalTag, 
#                             sname, SheetTagList, SheetTagValueList)
# Description:   define one function to check if the tag header has more than 256.
#                If yes, then we need to split the xml output into multiple TABs 
#                and write the value for it.
#                If no, just write the value for this tag.
# Input Value:   readOutputWB      --- the tag list of the sheet
#                WorkSheet         --- the sheet name from input Excel file.
#                finalTag          --- the final tag of the sheet.
#                sname         --- the sheet name.
#                SheetTagList      --- the tag in the sheet.
#                SheetTagValueList --- the value of tag.
# Return Value:  NULL
#################################################################################
def CheckAndWriteTag(readOutputWB, WorkSheet, finalTag, sname, SheetTagList, SheetTagValueList):   
        ''' Module to check and write value for tag. '''

        #For SipFilter, CRFProgrammableRule,remove useless field and value
        FirstTag = 0
        RemoveStart = 0
        RemoveEnd = 0

        TagLen = len(SheetTagList)
        if sname in ['SipFilter', 'CRFProgrammableRule']:
                while FirstTag < TagLen:
                        if SheetTagList[FirstTag] in ['_GRPSTART_HBR_RULE_EXECUTION_ORDER', '_GRPSTART_HBR_PATTERN_EXECUTION_ORDER']:
                                RemoveStart = FirstTag
                                for i in range(FirstTag, TagLen):
                                        if (SheetTagList[i] != '_GRPEND_Record'):
                                                if SheetTagList[i] in ['_GRPEND_HBR_RULE_EXECUTION_ORDER','_GRPEND_HBR_PATTERN_EXECUTION_ORDER']:
                                                        RemoveEnd = i + 1
                                        else:
                                                break
                                TagLen = TagLen -(RemoveEnd - RemoveStart)
                                while RemoveStart < RemoveEnd:
                                        del SheetTagList[RemoveStart]
                                        del SheetTagValueList[RemoveStart]
                                        RemoveEnd = RemoveEnd - 1
                                FirstTag = 0
                        else:
                                FirstTag = FirstTag + 1

        #check if the TAG header has more than 256 TAGs
        if(len(finalTag) > 256):
            # Add the suffix -1 to the original TAB name
            WorkSheet.name = sname + "-1-1"
   
            # Second TAB has a suffix of -1-2
            newSheet = sname + "-1-2"
            multipleWS = readOutputWB.add_sheet(newSheet,cell_overwrite_ok=True)
            multipleWSTag = finalTag[255:len(finalTag)]
   
            # Write the final TAG header in Row 0
            newTagCol = 0
            StartForSecSheet = 0
            for i in range(0,len(finalTag)):
                    # To avoid error for tables that have more TAGs than 256
                    if (newTagCol < 256):
                            if (sname == 'DiameterProfileTable') and (finalTag[i] == 'PROFILE_NAME'):
                                    StartForSecSheet = i
                            elif (finalTag[i] == 'PROFILEID'):
                                    StartForSecSheet = i

                            if (newTagCol == 255):
                                    WorkSheet.write(0,newTagCol,'_GRPEND_Record',myTAGstyle)
                            else:
                                    WorkSheet.write(0,newTagCol,finalTag[i],myTAGstyle)
                            newTagCol = newTagCol + 1

            # write the extra data in TAB 2
            # Need to add XML_ACTION and also the tableID/NAME in the second TAB
            for i in range(0, StartForSecSheet+1):
                    multipleWSTag.insert(i, finalTag[i])

            deltaTagCol = 0
            for i in range(0,len(multipleWSTag)):
                    multipleWS.write(0,deltaTagCol,multipleWSTag[i],myTAGstyle)
                    deltaTagCol = deltaTagCol + 1
        else:
            # Write the final TAG header in Row 0
            newTagCol = 0
            for i in range(0,len(finalTag)):
                    # To avoid error for tables that have more TAGs than 256
                    if(newTagCol < 256):
                            WorkSheet.write(0,newTagCol,finalTag[i],myTAGstyle)
                            newTagCol = newTagCol + 1

        # Write the values from Row 1 through the last record
        col = 0
        row = 1
        first_GRPSTART = 0
        deltaValCol = 0
        multidevic = 1

        for i in range(0, len(SheetTagList)):
            # To avoid error for tables that have more TAGs than 256
            if (col < 255):
                    # Save the column index from where the Record starts
                    if SheetTagList[i] in ["_GRPSTART_Record", "_GRPSTART_DNSConfiguration"]:
                            if (row==1):
                                    first_GRPSTART = i
                                    #for table in CreatWithDefauRecordList, write from the second row.
                                    if sname in CreatWithDefauRecordList:
                                            row = row + 1
                                            col = 0
                                            continue

                    #For IMSDeviceServer,FS5000DeviceServer, write from _GRPSTART_DNSConfiguration.
                    if (sname in ['IMSDeviceServer','FS5000DeviceServer']) and (row>1) and (multidevic==1):
                            if (SheetTagList[i] != "_GRPSTART_DNSConfiguration"):
                                    continue
                            else:
                                    multidevic=0

                    #For SipFilterHeaderRule, CRFHeaderPatternDefinition,
                    #not write PARAMETER_RULE_EXECUTION_ORDER or PARAMETER_EXECUTION_ORDER
                    if sname in ['SipFilterHeaderRule', 'CRFHeaderPatternDefinition']:
                            if SheetTagList[i] in ["PARAMETER_RULE_EXECUTION_ORDER", "PARAMETER_EXECUTION_ORDER"]:
                                    WorkSheet.write(row,col,'',myDATAstyle)
                                    continue

                    #check if reach the end of record in current row,
                    #if yes, set the col=0 and row+1 to next row;
                    #if no, keep write value for this row.
                    if SheetTagList[i] in ["_GRPEND_Record", "_GRPEND_ComponentParameters"]:
                            #For IMSDeviceServer,FS5000DeviceServer, end from ComponentParameters.
                            if sname in ['IMSDeviceServer','FS5000DeviceServer']:
                                    multidevic=1
                            elif sname in CreatWithDefauRecordList:
                                    if (row == 2):
                                            WorkSheet.write(row,col,SheetTagValueList[i-1],myDATAstyle)
                                            col=col+1
                            if col<(len(finalTag) - 1):
                                    for t in range(col,len(finalTag) - 1):
                                            WorkSheet.write(row,t,'',myDATAstyle)

                            WorkSheet.write(row,len(finalTag) - 1,SheetTagValueList[i],myDATAstyle)
                            row = row +1
                            col = 0
                            
                    else:
                            if (col==0) and (row>1):
                                    # For multiple Rows after the first Row
                                    # Write TAGs until the _GRPSTART_Record
                                    for col in range(0, first_GRPSTART+1):
                                            WorkSheet.write(row,col,SheetTagValueList[col],myDATAstyle)
                                            col = col +1
                            else:
                                    if sname in ['BGCFDARouteTable','IMSDeviceServer','FS5000DeviceServer',
                                                 'NGSSSIPia', 'H248Port', 'FS5000SIPia', 'DiamPort']:
                                            if SheetTagList[i] in ['CRITERIA', 'VALUE']:
                                                    WorkSheet.write(row, col, '_XCHAR_BLANK_', myDATAstyle)
                                            elif SheetTagList[i] in ['MEMBEREXTERNALGROUP','SERVICELABEL',
                                                                     'GLOBAL_PORT_IDENTIFIER']:
                                                    WorkSheet.write(row, col, '',myDATAstyle)
                                            else:
                                                    WorkSheet.write(row, col, SheetTagValueList[i],myDATAstyle)
                                    elif sname in CreatWithDefauRecordList:
                                            if (row == 2):
                                                    WorkSheet.write(row, col, SheetTagValueList[i-1],myDATAstyle)
                                            else:
                                                    WorkSheet.write(row, col, SheetTagValueList[i],myDATAstyle)
                                    else:
                                            WorkSheet.write(row, col, SheetTagValueList[i],myDATAstyle)
                                    col = col +1

                                    # end this row with '_GRPEND_Record' when col is 255.
                                    if (col == 255):
                                            WorkSheet.write(row, col, 'Required', myDATAstyle)
            else:
                    if (SheetTagList[i] == "_GRPEND_Record"):
                            multipleWS.write(row,deltaValCol,SheetTagValueList[i],myDATAstyle)
                            row = row +1
                            col = 0
                            deltaValCol = 0
                    else:
                            if (deltaValCol==0):
                                    # For multiple Rows after the first Row
                                    # Write TAGs until the PROFILEID
                                    for deltaValCol in range(0, first_GRPSTART+1):
                                            multipleWS.write(row,deltaValCol,SheetTagValueList[deltaValCol],myDATAstyle)
                                            deltaValCol = deltaValCol +1

                                    #write id for each record in second sheet.
                                    NumStart = 0
                                    for ns in range(0, len(SheetTagList)):
                                            if (SheetTagList[ns] == "_GRPSTART_Record"):
                                                    NumStart = NumStart + 1
                                                    if (NumStart == row):
                                                            multipleWS.write(row,deltaValCol,SheetTagValueList[ns+1],myDATAstyle)
                                                            break
                                                    else:
                                                            continue

                                    #for record value in second sheet, the XML action should be updaterecords.
                                    multipleWS.write(row,0,'UPDATERECORDS',myDATAstyle)
                                    deltaValCol = deltaValCol +1

                            multipleWS.write(row,deltaValCol,SheetTagValueList[i],myDATAstyle)
                            deltaValCol = deltaValCol + 1
           
        #handle tables with defult values.
        HandleDefaultTableValue(WorkSheet, sname, SheetTagList, SheetTagValueList)


##########################################################################################
# Function Name: HandleDefaultTableValue(WorkSheet, sname, SheetTagList, SheetTagValueList)
# Description:   define one function to handle tables with default value.
#                For tables in CreatWithDefauRecordList, they are tables which will have 
#                default records when they are created. So, need creat the table and keep
#                default records without operation. For tables which in DefulRecordList, 
#                the table is exist with one or more default records. For them, keep default
#                record with update and new record with create.
# Input Value:   WorkSheet         --- the sheet name from input Excel file.
#                sname             --- the sheet name.
#                SheetTagList      --- the tag in the sheet.
#                SheetTagValueList --- the value of tag.
# Return Value:  NULL
##########################################################################################
def HandleDefaultTableValue(WorkSheet, sname, SheetTagList, SheetTagValueList):
        ''' Moudle to handle tables with default value.'''
        numrow = 0
        row = 1
        j = 0
        #find the row number of this sheet.
        for i in range(0,len(SheetTagList)):
                if (SheetTagList[i] == "_GRPSTART_Record"):
                        numrow = numrow +1
        
        #for tables in CreatWithDefauRecordList, write default value from second row.
        if sname in CreatWithDefauRecordList:
                row = row + 1
                numrow = numrow +1

        #handle default value 
        while row < (numrow+1):
            if sname in CreatWithDefauRecordList:
                if (sname == "OnlineChargingTriggerData"):
                        WorkSheet.write(row, 0, 'UPDATE', myDATAstyle)
                else:
                        while j < len(SheetTagList):
                                if (SheetTagList[j] == "_GRPSTART_Record"):
                                        if (sname == "BGCFDirectAssistTable"):
                                                if SheetTagValueList[j+1] in ['411', 'x411', 'xxx5551212']:
                                                        WorkSheet.write(row, 0, '', myDATAstyle)
                                        else:
                                                id = int(SheetTagValueList[j+1])
                                                if (sname == "BGCFPaniAccessTypeTable"):
                                                        if id in range(1,39):
                                                                WorkSheet.write(row, 0, '',  myDATAstyle)
                                                elif (sname == "BGCFPaniAccessInfoTable"):
                                                        if id in range(1,11):
                                                                WorkSheet.write(row, 0, '', myDATAstyle)
                                                elif (sname == "FephFlowPolicyTable"):
                                                        if id in range(1,14):
                                                                WorkSheet.write(row, 0, '', myDATAstyle)
                                                elif (sname == "FephAggPacketPolicyTable"):
                                                        if id in range(1,3):
                                                                WorkSheet.write(row, 0, 'UPDATE', myDATAstyle)
                                                elif (sname == "OnlineChargingProfileTable"):
                                                        if (id == 0):
                                                                WorkSheet.write(row, 0, '', myDATAstyle)
                                                elif (sname == "FephRemoteIpPolicyTable"):
                                                        if (id == 1):
                                                                WorkSheet.write(row, 0, 'UPDATE', myDATAstyle)
                                                elif sname in ["SipTrustedRateLimitSet", "SipUntrustedRateLimitSet",
                                                               "SipAggrTrustedRateLimitSet", "SipAggrUntrustedRateLimitSet"]:
                                                        if id in range(1,7):
                                                                WorkSheet.write(row, 0, 'UPDATE', myDATAstyle)
                                        j = j + 1
                                        break
                                else:
                                        j= j + 1
            elif sname in DefulRecordList:
                while j < len(SheetTagList):
                        if (SheetTagList[j] == "_GRPSTART_Record"):
                                idex = int(SheetTagValueList[j+1])
                                if (sname == "HomeNetworkIdentifierTable"):
                                        if (idex == 1):
                                                WorkSheet.write(row, 0, 'UPDATE', myDATAstyle)
                                elif (sname == "AudioCodec"):
                                        if idex in range(1,50):
                                                WorkSheet.write(row, 0, 'UPDATE', myDATAstyle)
                                elif (sname == "VideoCodec"):
                                        if idex in range(1,25):
                                                WorkSheet.write(row, 0, 'UPDATE', myDATAstyle)
                                else:
                                        if (idex == 0):
                                                WorkSheet.write(row, 0, 'UPDATE', myDATAstyle)
                                j = j + 1
                                break
                        else:
                                j = j + 1 
            elif sname in ['TcpConnections', 'IMSService']:
                WorkSheet.write(row, 0, 'UPDATERECORDS', myDATAstyle)
            elif sname in ['FephEnabledServicesTable', 'FephIpsecParametersTable']:
                WorkSheet.write(row, 0, '', myDATAstyle)
            elif sname in ['NGSSSIPia', 'H248Port', 'FS5000SIPia', 'DiamPort']:
                #Exchange the station 'NODE' and '_GRPSTART_Record' for xlsprov format.
                if (row == 1):
                        WorkSheet.write(row-1, 1, "NODE", myTAGstyle)
                        WorkSheet.write(row-1, 2, "_GRPSTART_Record", myTAGstyle)
                while j < len(SheetTagList):
                        if (SheetTagList[j] == "_GRPSTART_Record"):
                                WorkSheet.write(row, 1, SheetTagValueList[j+1],myDATAstyle)
                                WorkSheet.write(row, 2, SheetTagValueList[j],myDATAstyle)
                                j = j + 1
                                break
                        else:
                                j = j + 1
                
            row = row + 1                                    



#########################################################################################################
# Function Name: HandleSiteSpecificData(readOutputWB)
#
# Description:   Define one function to add Site Specific Data into output xls file. At the same time,
#                add the current release into output xls file(row=0, cel=1)
#
# Input Value:   readOutputWB --- the work book of output xls file.
#
# Return Value:  readOutputWB --- the work book which have added Site Specific Data into output xls file.
#########################################################################################################
def HandleSiteSpecificData(readOutputWB):
        ''' Model to add Site Specific Data into output xls file.'''
        # Add the sheet 'Site Specific Data' into workbook of output xls file.
        SiteSheet = readOutputWB.add_sheet('Site Specific Data',cell_overwrite_ok=True)
        
        # Get the current release of service and write it into Site Specific Data for xlsprov
        GetVersionCmd = 'grep \'SU:\' /etc/pkgconf/version | tail -n 1 | cut -f2 -d\' \''
        CurrentRelease = os.popen(GetVersionCmd).read()

        SiteSheet.write(0, 1, CurrentRelease, myDATAstyle)

        return readOutputWB




###################################################################
# Function Name: InputCMDAnalysis()
# Description:   Define one function to analysis the input command.
#                If valid cmd, check the option and execute.
#                If invalid cmd, print error and the help menu.
# Input Value:   NULL
# Return Value:  NULL
###################################################################
def InputCMDAnalysis():
        '''Module to analysis the input command when use this tool.'''
        global XlsFile

        try:
                opts,args = getopt.getopt(sys.argv[1:], "hf:", ["help","file="])
                
                #analysis the parameters with '-' or '--' in opts from the input command.
                for option, value in opts:
                        if option in ("-h", "--help"):
                                #print the help menu for this tool
                                PrintMenuHelp()
                                sys.exit()
                        elif option in ("-f", "--file"):
                                #save the input excel file to XlsFile
                                XlsFile = value
                #for no parameter with '-' or '--', just save the value to XlsFile from args.
                for param in args:
                        XlsFile = param
                        
        except getopt.GetoptError:
                PrintAndSaveLog('Invalid input option.')
                PrintAndSaveLog('Please follow help menu to use this tool.')
                PrintMenuHelp()
                sys.exit(1)



##################################################################
# Function Name: PrintMenuHelp()
# Description:   define one function to show how to use this tool.
# Input Value:   NULL
# Return Value:  NULL
##################################################################
def PrintMenuHelp():
        '''Module to show how to use this tool.'''

        print '\n~~~~~~~~~~~~~~~~~~~~~~~~~~~HELP MENU~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
        print '| ',sys.argv[0]
        print '|      Do the Save_to_file ability to generate xlsprov template file from '
        print '|      data provisioned in the config DB of an already installed CTS, ISC, '
        print '|      or SCG system.'
        print '|'
        print '|  OPTIONS:'
        print '|  Common Usage: '
        print '|      ./xlsread.py [-f] ExcelFile ' 
        print '|      ./xlsread.py -h '
        print '|'
        print '|      -f | --file    the Excel template file name. '
        print '|                     Note: This Excel template file must cotain all table '
        print '|                           names in the CTS/ISC/SCG system. xlsread.py will'
        print '|                           get the data from installed CTS/ISC/SCG system '
        print '|                           according to those tables.' 
        print '|      -h | --help    Display the help information.'
        print '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~\n'



###############################################################
# Function Name: main()
# Description:   define the main function of this tool
# Input Value:   NULL
# Return Value:  NULL
###############################################################
def main():

        global cnfgip	# IP for config
        global XlsFile	# Input file used as command parameter.

        InputCMDAnalysis()
        inputWorkBook = OpenInputExcel(XlsFile)

        tmpsname = [] 
        
        # Get CNFG IP
        cnfgip = GetIP() 

		if cnfgip == "":
			PrintAndSaveLog("Cannot get the config's IP\n")
			return("")
        
        # tmpsname stores the sheet names from the Input Excel File
        SaveSheetNameFrmXls(tmpsname)
        
        # New Workbook for the Output Excel File
        readOutputWB = xlwt.Workbook()
        
        # Add Site Specific TABs into Output Excel File for xlsprov
        readOutputWB = HandleSiteSpecificData(readOutputWB)
        
        # Do xml read on the Tables listed in the tmpsname - sheetnames list
        # And write the value to output Excel file
        ReadSheetAndWrite(readOutputWB, inputWorkBook, tmpsname)

        readOutputWB._Workbook__active_sheet = 0
        
        # Save the output Excel file to current path.
        readOutputWB.save('CfgXlsDataDump.xls')
        
        LogMsg = '\nLogs are stored under /export/home/lss/logs/xlsRead-dbdump'
        PrintAndSaveLog(LogMsg)
        PrintAndSaveLog("\nRead output is stored in CfgXlsDataDump.xls")

if __name__ == "__main__":
        main()


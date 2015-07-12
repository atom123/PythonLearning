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
UpdateList = ["GlobalParameters", "NGSSParameters", "DiamAppIDParameters",
              "SS7M3uaParameters", "SipErrorTreatmentTable"]

#dynamic table list. For those tables, there are dynamic child tables which can be configured by users,
#the format of xml response of them are not fixed. So, handle them specially.
DynamicTableList = ["SDPMediaSubsPolicyTable","IcsiTable","SipLinkSetTable","SDPProfileTable",
                    "MrfAnnouncementInterProf"]

#for table in this list, need to mark default record to UPDATE and mark the new record to CREATE.
DefulRecordList = ["SCTPProfileTable", "SCTPConnectionManagementProfile","SIPStackTimerProfile",
                   "IMSACRChargingProfileTable","HomeNetworkIdentifierTable", "AudioCodec", "VideoCodec"]

#tables which have muilt records when it was created.
CreatWithDefauRecordList = ["BGCFPaniAccessTypeTable", "BGCFPaniAccessInfoTable", 
                            "BGCFDirectAssistTable", "ScscfAsAffiliationTable", 
                            "OnlineChargingProfileTable", "OnlineChargingTriggerData"]

#set the font style for the Columns and the Rows for the output Excel file 
myTAGstyle = xlwt.easyxf('pattern: pattern solid, fore_colour pale_blue; font: name Palatino Linotype, bold on;')
myDATAstyle = xlwt.easyxf('pattern: pattern solid, fore_colour light_yellow; font: name Palatino Linotype;')

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
# Description:   define one function to save sheet names from input Excel file.
# Input Value:   sname   --- the sheet name.
# Return Value:  NULL
##############################################################################
def SaveSheetNameFrmXls(sname):
        ''' Module to save all sheets name from the input Excel template. '''

        XlsWorkBook = open_workbook(XlsFile, 'rb')
        
        #save all sheets to the list(sname) from input Excel file except 'Index' and 'Site Specific Data'.
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

    	#set the root of xml script tree for sheetname
    	root = ElementTree.Element('Request',{'Action': action})
    	AttriSheet = ElementTree.Element(sheetname)

        if NodeID != 0:
            SubAttriSheet = ElementTree.SubElement(AttriSheet,'NODE')
            SubAttriSheet.text = str(NodeID)

        root.append(AttriSheet)

       	#create an element tree object from the root element.
        tree = ElementTree.ElementTree(root)

        if NodeID == 0:
            tree.write(tabname+'.xml','utf8')
        else:
            strnodeid = str(NodeID)
            modtabname = tabname+strnodeid;
            tree.write(modtabname+'.xml','utf8')


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
                #make sure to use str on TableElement.text only for Non-Empty Strings
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
        DBTable = ''
        OutDBTable = ''
        global SIPiaOk

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
        """Get a sheet by name from xlwt.Workbook
        Returns None if no sheet with the given name is present.
        """

        try:
            for idx in itertools.count():
                sheet = XlsWorkBook.get_sheet(idx)
                #check if sheet name with postfix
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
            LogMsg = 'An error was encountered when open the input file: ' + (InputXlsFile)
            PrintAndSaveLog(LogMsg)
            PrintAndSaveLog('\n!!!The given input file maybe in an invalid format.!!!')
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

#Carl
###############################################################
# Function Name: GetFsdbGlsIp(FsdbGlsFlag)
# Description:   define one function to get the config IP 
# Description:   define one function to get the FSDB or GLS IP #Jeffrey 
#                address of this server.
# Input Value:   FsdbGlsFlag	---	Flag to confirm whether it
#                                   is FSDB or GLS.
#                                   1 - FSDB
#                                   2 - GLS
# Return Value:  NULL
###############################################################
def GetFsdbGlsIP(FsdbGlsFlag):
        ''' Module to get the config IP address of this server. '''
        global fsdbglsip

        if FsdbGlsFlag == 1:	#Jeffrey 
            fsdbglsgrep = "fsdb"
            fsdbglsprint = "CNFG IP: "

        elif FsdbGlsFlag ==2:	#Jeffrey 
            fsdbglsgrep = 'gls'
            fsdbglsprint = "GLS IP: "

        fsdbglsIPCmd = 'grep ^' + fsdbglsgrep + ' /var/opt/lib/sysconf/service_ip.data | grep "\-g0" | grep floating | cut -d";" -f6'
        fsdbglsipLine = os.popen(fsdbglsIPCmd)
        fsdbglsip = fsdbglsipLine.readline().strip('\n')

        FSDBGLSIP = fsdbglsprint + fsdbglsip
        PrintAndSaveLog(FSDBGLSIP)

#Jeffrey 
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
    	aimed_IPCmd = 'grep ^ ' + aimed_Grep            \
			+ ' /var/opt/lib/sysconf/service_ip.data'   \
			+ ' | grep "\-g0" | grep floating'          \
			+ ' | cut -d ";" -f 6'

        aimed_Line= os.popen(aimed_IPCmd)
    	aimed_IP = aimed_Line.readline().strip('\n')

    	AIMEDIP = aimed_Print + aimed_IP 
    	PrintAndSaveLog(AIMEDIP)

        return aimed_IP



###############################################################################
# Function Name: RmIndexSSTabFrmXls(readOutputWB)
# Description:   define one function to remove 'Index' and the 
#                'Site Specific TABs' from the output Excel file.
# Input Value:   readOutputWB --- the workbook of output Excel file.
# Return Value:  readOutputWB --- if 'Index' and the 'Site Specific TABs' exist,
#                                 remove it, return the update workbook.
###############################################################################
def RmIndexSSTabFrmXls(readOutputWB):
        ''' Remove Index and the Site Specific TABs from the copied excel for now. '''

        #get the num of sheets in excel file
        NumWorkBook = open_workbook(XlsFile, 'rb')
        NumSheet = len(NumWorkBook.sheets())

        i = 0
        #loop all sheets in output excel, remove 'Index' and 'Site Specific Data'.
        while i < NumSheet:
            if readOutputWB._Workbook__worksheets[i].name in ['Index', 'Site Specific Data']:
                    readOutputWB._Workbook__worksheets.remove(readOutputWB._Workbook__worksheets[i])
                    NumSheet=NumSheet-1
                    i=0
            else:
                    i=i+1

        readOutputWB._Workbook__active_sheet = 1

        return readOutputWB



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
#Carl
##########################################################################
# Function Name: GetFsdbGlsFlag(FsdbGlsFlag,OutputSheetName,InputSheetName)	
# Function Name: GetFsdbGlsFlag(OutputSheetName,InputSheetName)	#Jeffrey 
# Description:   define one function to get the flag of FSDB and GLS
#                For some input xls file template, the worksheet name maybe
#                contains '-fsdbx' or '-gls' postfix. 
# Input Value:   FsdbGlsFlag --- the flag of FSDB, GLS: 0 for Config DB, 1 for FSDB, 2 for GLS DB
# Input Value:   FsdbGlsFlag --- the flag of FSDB, GLS: 1 for FSDB, 2 for GLS DB	#Jeffrey
#                OutputSheetName  --- the real sheet name for FSDB/GLS table.
#                InputSheetName  --- the sheet name got from input xls file.
# Return Value:  NULL
##########################################################################
#def GetFsdbGlsFlag(FsdbGlsFlag,OutputSheetName,InputSheetName):
def GetFsdbGlsFlag(OutputSheetName,InputSheetName):	#Jeffrey 
        ''' Model to get the flag of FSDB and GLS.'''

        global FsdbGlsFlag	#Jeffrey 
        #FsdbGlsFlag=0	#Jeffrey 

        #fsdbre = re.compile(r'.+(?=-fsdb\d+)')
        fsdbre = re.compile(r'.+(?=-fsdb\d)')	#Jeffrey 
        fsdbList = fsdbre.findall(InputSheetName)

        if len(fsdbList) != 0:	#Jeffrey 
            OutputSheetName = fsdbList[0]
            FsdbGlsFlag = 1
        else:	#Jeffrey 
            glsre = re.compile(r'.+(?=-gls)')
            glslist = glsre.findall(InputSheetName)
            if len(glslist) != 0:	#Jeffrey 
                OutputSheetName = glslist[0]
                FsdbGlsFlag = 2

##########################################################################
# Function Name: ReadSheetAndWrite(readOutputWB, inputWorkBook, tmpsname) 
# Description:   define one function to do xml read by sheet 
#                and write the value to output Excel file.
# Input Value:   readOutputWB  --- the workbook of output Excel file.
#                inputWorkBook --- the workbook of input Excel file.
#                tmpsname      --- the sheet name list of input Excel file.
# Return Value:  NULL
##########################################################################
def ReadSheetAndWrite(readOutputWB, inputWorkBook, tmpsname):
        ''' Module to do xml read by sheet
        and write the value to output Excel file.
        '''    
        global SIPiaOk	#Jeffrey

        #before do xml read, clean duplicate sheet name with postfix.
        NumSheet = len(tmpsname)
        for sn in range(0, NumSheet):
                tmpsname[sn] = RenameWorkSheet(tmpsname[sn])

        LoopStart = 0
        LoopNext = 1		#Jeffrey

        while LoopStart < NumSheet:
                while LoopNext < NumSheet:
                        if (tmpsname[LoopStart] == tmpsname[LoopNext]):
                                del tmpsname[LoopNext]
                                readOutputWB._Workbook__worksheets.remove(readOutputWB._Workbook__worksheets[LoopNext])
                                NumSheet = NumSheet - 1		#Jeffrey
                        else:
                                LoopNext = LoopNext + 1

                LoopStart = LoopStart + 1
                LoopNext = LoopStart + 1

        #loop the sheet name list to do xml read and write.
        for i in range(0, len(tmpsname)):

            # check if request path exists.
            if (not os.path.exists(ReqPath)):
                os.mkdir(ReqPath)
#Jeffrey 
            # check if response path exists.
            if (not os.path.exists(ResPath)):
                os.mkdir(ResPath)

            rdsname = join(ReqPath, tmpsname[i])
            InName = rdsname + '.xml'

            OutName = tmpsname[i] + '.out'	#Jeffrey 
            OutFile = join(ResPath, OutName)

            #Carl            
            GetFsdbGlsFlag(FsdbGlsRealName, tmpsname[i])	#Jeffrey 
            if FsdbGlsFlag > 0:	#Jeffrey 
                GetFsdbGlsIp(tmpsname[i], FsdbGlsFlag)
                if FsdbGlsFlag == 1:	#Jeffrey 
                    FsdbGlsPort = 7856
                else:	#Jeffrey 
                    FsdbGlsPort = 6856
                GenerateLoginLogoff(Login, Logoff, FsdbGlsFlag, tmpsname[i], inputWorkBook)
				#Loin/Logoff is file name
                GenXmlScript(rdsname, tmpsname[i], 0)
                cmd = 'python xml2fsdbgls.py %s %s %s %s %s %s' % (FsdbGlsIp,FsdbGlsPort,Login,Logoff,InName,OutFile)
                os.system(cmd)
                continue


            if tmpsname[i] in ['NGSSSIPia', 'H248Port', 'FS5000SIPia', 'DiammPort']:
                SIPiaOk = 0
                GenXmlForSIPiaPort(rdsname,tmpsname[i])
            else:
                SIPiaOk = 1
                GenXmlScript(rdsname, tmpsname[i], 0)
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

            #Logic is to read the sheets from the Copied excel workbook - readOutputWB and write in them
            #For tables that have records in the Xml output we will clear the sheet and write the xml
            #For tables that don't have data in the Xml output we will just keep the copied sheet as-is
            if tmpsname[i] in XmlDict:
            	WorkSheet = GetSheetByName(readOutputWB,tmpsname[i])
            else:
               	WorkSheet = GetSheetByName(readOutputWB,sname)

            #Need to clear the Worksheet contents before we start writing the XML Output
            numRows = 0
            numCols = 0
            for s in inputWorkBook.sheets():
                #check if sheet name with postfix
                s.name = RenameWorkSheet(s.name)
                if (s.name == sname):
                	numRows = s.nrows
                   	numCols = s.ncols
                  	break

            for node in ListNode:
                if (node.attrib['Status'] == "OKAY"):
                        # Clear all rows for sheet have right response.
                        for rowIndex in range(0,numRows):
                                for colIndex in range (0, numCols):
                                        WorkSheet.write(rowIndex,colIndex,"")

                        tables = root.findall('Response/'+sname)

                        SheetTagList = []
                        SheetTagValueList = []

                        SheetTagList.append('XML_ACTION')
                        if tmpsname[i] in UpdateList:
                                SheetTagValueList.append('UPDATE')
                        else:
                                SheetTagValueList.append('CREATE')

                        for table in tables:
                                GenTagForSheet(table,SheetTagList,SheetTagValueList,sname)
                        if sname in XlsDict:
                                sname = XlsDict[sname]

                        #generate the tag header for sheet in Output Excel file.
                        if sname not in DynamicTableList:
                                GenTagAndWrite(readOutputWB, WorkSheet, SheetTagList, SheetTagValueList, sname)
                        else:
                                GenTagAndWriForDynTab(readOutputWB,WorkSheet,SheetTagList,SheetTagValueList,sname)

                else:

                        # Clear all rows except the first row for sheet have no response.
                        for rowIndex in range(1,numRows):
                                for colIndex in range (0, numCols):
                                        WorkSheet.write(rowIndex,colIndex,"")

                        LogMsg = 'No Records found in the System for Table: ' + sname
                        PrintAndSaveLog(LogMsg)
                        LogMsg =  'Info from the Input Template will be retained.'
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
        for slen in range(0,len(SheetTagList)):
                if(SheetTagList[slen] != '_GRPEND_Record'):
                        if(NumCol < 256):
                                if(SheetTagList[slen] == '_GRPSTART_Record'):
                                        FirstStart = slen
                                WorkSheet.write(0,NumCol,SheetTagList[slen],myTAGstyle)
                                WorkSheet.write(1,NumCol,SheetTagValueList[slen],myDATAstyle)
                                NumCol = NumCol + 1
                else:
                        WorkSheet.write(0,NumCol,SheetTagList[slen],myTAGstyle)
                        WorkSheet.write(1,NumCol,SheetTagValueList[slen],myDATAstyle)
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
                        #for other records, the xml action should be 'UPDATE'.
                        NewWorkBook.write(1,0,'UPDATE',myDATAstyle)

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
        mutlidevic = 1

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
                    if (sname in ['IMSDeviceServer','FS5000DeviceServer']) and (row>1) and (multdevic==1):
                            if (SheetTagList[i] != "_GRPSTART_DNSConfiguration"):
                                    continue
                            else:
                                    multdevic=0

                    #check if reach the end of record in current row,
                    #if yes, set the col=0 and row+1 to next row;
                    #if no, keep write value for this row.
                    if SheetTagList[i] in ["_GRPEND_Record", "_GRPEND_ComponentParameters"]:
                            #For IMSDeviceServer,FS5000DeviceServer, end from ComponentParameters.
                            if sname in ['IMSDeviceServer','FS5000DeviceServer']:
                                    multdevic=1
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
                        WorkSheet.write(row, 0, '', myDATAstyle)
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
                                                elif (sname == "OnlineChargingProfileTable"):
                                                        if (id == 0):
                                                                WorkSheet.write(row, 0, '', myDATAstyle)
                                                elif (sname == "ScscfAsAffiliationTable"):
                                                        if (id == 1):
                                                                WorkSheet.write(row, 0, '', myDATAstyle)
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
                PrintAndSaveLog('Invalid input option!!!')
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

        global XlsFile	# Input file used as command parameter.	
        global cnfgip	# IP for config
        global fsdbip	# IP for fsdb
        global glsip	# IP for gls

        InputCMDAnalysis()
        inputWorkBook = OpenInputExcel(XlsFile)

        tmpsname = [] 
        
        # Get CNFG IP
        cnfgip = GetIP() 	#Jeffrey

        if cnfgip == "":
            PrintAndSaveLog("Cannot get the config's IP\n")
            return ("")
        
        # tmpsname stores the sheet names from the Input Excel File
        SaveSheetNameFrmXls(tmpsname)
        
        # New Workbook for the Output Excel File
        readOutputWB = xlwt.Workbook()
        
        # Work Book that will be written the final output Excel file.
		# The "copy" is the method from workbook object. 
        readOutputWB = copy(inputWorkBook)
        
        # Remove Index and the Site Specific TABs from the copied excel for now
        readOutputWB = RmIndexSSTabFrmXls(readOutputWB)
        
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


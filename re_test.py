#-*- coding:utf-8 -*-
#最上一句如果不加上的话不能输入汉字，不然编译不过
#/usr/bin/python

import os
import sys
import re

####################################################
# 	一行一行的匹配
#	比如Test_2.xml中的内容如下：
#	<Response Staus="OKAY" CongLvl="jeffrey">
#		<a>test1</a>
#		<b>test2</b>
#	</Response>
#	<Response Staus="OKAY" CongLvl="guan">
#		<a>test1</a>
#		<b>test2</b>
#	</Response>
#	<Response Staus="OKAY" CongLvl="zenghui">
#		<a>test1</a>
#		<b>test2</b>
#	</Response>
#	<Response Staus="OKAY" CongLvl="jeguan">
#		<a>test1</a>
#		<b>test2</b>
#	</Response>
#	那么匹配后的得到结果为
#	<Response Staus="OKAY" CongLvl="jeffrey">
#	<Response Staus="OKAY" CongLvl="guan">
#	<Response Staus="OKAY" CongLvl="zenghui">
#	<Response Staus="OKAY" CongLvl="jeguan">
#
####################################################
def re_test():

	# 原文件
	filename = r'C:\Users\jeguan\Desktop\Test_2.xml'
	# 匹配得到的内容存储在Test_2_bk.xml中
	new_file = r'C:\Users\jeguan\Desktop\Test_2_bk.xml'

	# 打开原文件
	open_file = open(filename, 'r')
	read_file = open_file.readlines()
	# 打开目标文件，即：存放匹配结果的文件
	newfile = open(new_file, 'wb')

	# 匹配以<Response开头并且有CongLvl字符串的行,注意，
	# 这里是非lazzy匹配,并且是一行一行匹配，即，遇到
	# '\n'就会结束
	patt =  re.compile(r'^<Response.*CongLvl.*')

	# 遍历原文件的所有行,如果找到就会存盘
	for line in read_file:
		match = patt.search(line)
		if match:
			m = match.group(0)
			newfile.write(m)

	open_file.close()
	newfile.close()


#################################################################################
#	多行匹配，即：可以匹配一个文本中的特定段落。这里主要是要用到re模块中的re.S
#	它表示当用'.'来进行匹配的时候，可以忽略掉'\n'，这一点与'.'正常的规则是不一样的
#	
#	另外一个要注意的地方是这里使用了lazzy匹配的方式。当有多个Response>出现的时候，
#	它只会匹配第一次出现的地方。比如：
#	<Response Staus="OKAY" CongLvl="jeguan">
#		<a>test1</a>
#		<b>test2</b>
#	</Response>
#	<Response Staus="OKAY" CongLvl="zenghuiguan">
#		<a>test3</a>
#		<b>test4</b>
#	</Response>
#	当用lazzy方式的时候，只会匹配到第一次出现Response>的地方
#	本文中匹配得到的结果为：
#	<Response Staus="OKAY" CongLvl="jeguan">
#		<a>test1</a>
#		<b>test2</b>
#	</Response>
#
################################################################################
def re_testsearch():
	# 
	filename = r'C:\Users\jeguan\Desktop\Test_2.xml'
	new_file = r'C:\Users\jeguan\Desktop\Test_2_bk.xml'

	try:
		open_file = open(filename, 'r')
	except:
		print("An error was encountered when opeing the %s" % filename)
		exit(1)

	read_file = open_file.readlines()
	newfile = open(new_file, 'wb')

	# re.S means: Make the '.' special character match any character at all, 
	# including a newline; without this flag, '.' will match anything except a newline.
	# '(.+?)' means: this is a lazzy match. When the fist 'Response>' is found, then
	# it will not try to match the next 'Response>'
	re_patt = re.compile(r'<Response Status="OKAY" CongLvl="LEVEL0"*(.+?)Response>', re.S)

	str1 = ""
	# 把读出的行放在str1中
	for line in read_file:
		str1 = str1 + line

	result = re_patt.search(str1)
	newfile.write(result.group(0))

	print(result.group())

############################################################################
#	The same to re_testsearch(), the difference is ET.fromstring(str1) is 
#	used, that is, there is no need to save the matched "restult" into a file
#	we can analize the content of the "result" directly.
#refer-to:https://docs.python.org/2/library/xml.etree.elementtree.html?highlight=elementtree
############################################################################
def re_testsearch2():

	from xml.etree import ElementTree as ET

	filename = r'C:\Users\jeguan\Desktop\Test_2.xml'

	open_file = open(filename, 'r')
	read_file = open_file.readlines()

	# re.S means: Make the '.' special character match any character at all, 
	# including a newline; without this flag, '.' will match anything except a newline.
	# '(.+?)' means: this is a lazzy match. When the fist 'Response>' is found, then
	# it will not try to match the next 'Response>'
	re_patt = re.compile(r'<Response Status="OKAY" CongLvl="LEVEL0"*(.+?)Response>', re.S)

	str1 = ""
	# 把读出的行放在str1中
	for line in read_file:
		str1 = str1 + line

	# re_patt.search() returns an object for MatchObject; 
	# "result" is a string.
	# 从左到右，去计算是否匹配，如果有匹配就返回，所以，最多只会匹配到一个，而不会是多个，
	# 如果要匹配多个，请问re.findall()
	result_Type = re_patt.search(str1)

	# If cannot find the string then re_patt.search(str1) will return "NoneType" 
	if (result_Type != "NoneType"):
		result = result_Type.group(0)
	
	# This code only used to make it more clear that "result" is used as a tree here.
	tree = result

	root = ET.fromstring(tree)
	#print(root.tag)
	#print(root.attrib)

	# Element has some useful methods that help iterate recursively over all 
	# the sub-tree below it (its children, their children, and so on). 
	# For example, Element.iter()
	#
	global dict_child 
	dict_child = {}

	for child in root.iter():
		dict_child[child.tag] = child.text	
		print(child.tag)
		print(child.attrib)
		print(child.text)

	#print(dict_child)

######################################################################
#	Test re.findall() and re.search()
# 	re.findall() will find ALL the matched string
#		['123', '123', '234']
#	re.search() will only return the FIRST matched string
#		123
######################################################################
def test_findall_search():
	str1 = '123abc123abc234abc'
	
	re_str = re.compile(r'\d+')
	re_findall = re_str.findall(str1)

	print(re_findall)	# ['123', '123', '234']

	re_search = re_str.search(str1)

	print(re_search.group(0))	# 123

####################################################################
#	test for re expression
###################################################################
def test2():
	str1 = 'ClientName-fsdb-1'

	#SheetWithPostfix = re.compile(r'.+(?=-\d)')
	#Result--> ['ClientName-fsdb']

	SheetWithPostfix = re.compile(r'.+(?=-\d)')
	#Result--> ['ClientName-fsdb']

	SheetNameList = SheetWithPostfix.findall(str1)

	print(SheetNameList)

#################################################################
# Test for "cut"
# the input file is as follows:
#
#APPLYING-R30.14.01.7070: R30.14.01.7070 TYPE: OFC DATE: Thu Apr  2 14:34:54 CST 2015
#SU: APPLYING-R30.14.01.7070 TYPE: OFC DATE: Thu Apr  2 14:50:31 CST 2015
#APPLYING-R30.14.01.7070: R30.14.01.7070 TYPE: OFC DATE: Thu Apr  2 14:50:32 CST 2015
#SU: R30.14.01.7070 TYPE: OFC DATE: Thu Apr  2 15:00:12 CST 2015
#APPLYING-R30.14.01.7070: R30.14.01.7070 TYPE: OFC DATE: Thu Apr  2 17:08:28 CST 2015
#SU: APPLYING-R30.14.01.7070 TYPE: OFC DATE: Thu Apr  2 17:24:16 CST 2015
#APPLYING-R30.14.01.7070: R30.14.01.7070 TYPE: OFC DATE: Thu Apr  2 17:24:17 CST 2015
#SU: R30.14.01.7070 TYPE: OFC DATE: Thu Apr  2 17:33:37 CST 2015
#APPLYING-R30.14.01.0700: R30.14.01.0700 TYPE: OFC DATE: Fri May  8 17:37:46 CST 2015
#SU: APPLYING-R30.14.01.0700 TYPE: OFC DATE: Fri May  8 17:55:26 CST 2015
#APPLYING-R30.14.01.0700: R30.14.01.0700 TYPE: OFC DATE: Fri May  8 17:55:27 CST 2015
#SU: R30.14.01.0700 TYPE: OFC DATE: Fri May  8 18:04:36 CST 2015
#APPLYING-R30.22.00: R30.22.00 TYPE: OFC DATE: Mon May 11 12:13:09 CST 2015
#SU: APPLYING-R30.22.00 TYPE: OFC DATE: Mon May 11 12:30:51 CST 2015
#APPLYING-R30.22.00: R30.22.00 TYPE: OFC DATE: Mon May 11 12:30:52 CST 2015
#SU: R30.22.00 TYPE: OFC DATE: Mon May 11 12:40:30 CST 2015
#APPLYING-R30.14.01.0700: R30.14.01.0700 TYPE: OFC DATE: Mon May 11 16:33:39 CST 2015
#SU: APPLYING-R30.14.01.0700 TYPE: OFC DATE: Mon May 11 16:51:10 CST 2015
#APPLYING-R30.14.01.0700: R30.14.01.0700 TYPE: OFC DATE: Mon May 11 16:51:11 CST 2015
#SU: R30.14.01.0700 TYPE: OFC DATE: Mon May 11 17:00:09 CST 2015
#APPLYING-R30.14.01.7070: R30.14.01.7070 TYPE: OFC DATE: Mon Jun 29 22:31:18 CST 2015
#
# The result is:
# 		R30.14.01.0700
################################################################
def cut_test():
	cmd = 'grep \'SU:\' ./test.txt | tail -n 1 | cut -f2 -d\' \''
	os.system(cmd)
	

if __name__ == "__main__":
	#re_test()
	#re_testsearch()
	#test_findall_search()
	#test2()
	

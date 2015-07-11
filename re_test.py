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

	open_file = open(filename, 'r')
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


if __name__ == "__main__":
	#re_test()
	re_testsearch()
	
	

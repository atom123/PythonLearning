#!/usr/bin/python
#-*- coding:utf-8 -*-
import re
import os
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
	 
	filename = r'C:\Users\jeguan\Desktop\Test_2.xml'
	new_file = r'C:\Users\jeguan\Desktop\Test_2_result.xml'

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

	path = r'C:\\Users\\jeguan\\Desktop'
	os.rename(os.path.join(path, 'Test_2_result.xml'), os.path.join(path, 'new_jeguan.xml'))
	print(result.group())


if __name__ == "__main__":
	re_testsearch()

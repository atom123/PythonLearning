#encoding : utf-8

import os
import sys
from xml.etree import ElementTree as ET

def func():

	xlsfile = r'C:\Users\jeguan\Desktop\123.xlsx'

	book = xlrd.open_workbook(xlsfile)

	print("book.nsheets = %d" %  book.nsheets)

	for sheet_index in range(book.nsheets):
		sheet_ix = book.sheet_by_index(sheet_index)
		print sheet_ix.name	
	sheet0 = book.sheet_by_index(0)
	sheet1 = book.sheet_by_index(1)

	print sheet0.row(0)
	print("sheet0.row_silce(4,1)")
	print sheet0.row_slice(5,3)
	print book.sheet_names()
	for sheet_name in book.sheet_names():
		print book.sheet_by_name(sheet_name)

def func2():
	from xml.etree import ElementTree
	xlsfile = r'C:\Users\jeguan\Desktop\Test_1.txt'

	root = ElementTree.parse(xlsfile).getroot()
	ListNode = root.getiterator('rank')
	print ListNode

def main():
	global b
	b = sys.argv[1]
	print(b)

########################################################################
#	Restult for this function
#	<?xml version='1.0' encoding='utf8'?>
#	<Request Action="UPDATE"><jeguan><NODE1>3</NODE1></jeguan></Request>
#########################################################################
def xml_tree():
	# create a root for the tree
	root = ET.Element("Request", {"Action": "UPDATE"})

	Attrib = ET.Element("jeguan")

	SubAttrib = ET.SubElement(Attrib, "NODE1")
	SubAttrib.text = str(34)
	# Adds the element subelement to the end of this elements internal
	# list of subelements
	root.append(Attrib)
	tree = ET.ElementTree(root)

	tree.write("jeugan" + ".xml", "utf8")

	# Writes an element tree or element structure to sys.stdout.
	# This function should be used for debugging only.
	# ElementTree.dump(root)

########################################################################
#	Result for this funciton:
#		<?xml version='1.0' encoding='utf8'?>
#  		<Request Action="UPDATE"><child1 name="HAHA">1<child2>2</child2></child1></Request>
##########################################################################
def xml_tree2():
	root = ET.Element("Request", {"Action": "UPDATE"})

	child1 =  ET.SubElement(root, "child1", {"name": "HAHA"})
	child1.text = str(1)

	child2 = ET.SubElement(child1, "child2")
	child2.text = str(2)

	tree = ET.ElementTree(root)
	tree.write("jeguan" + ".xml", "utf8")

##########################################################################
# Test for ET.fromstring():
# Input string "sting" is as follows:
#  		<Request Action="UPDATE"><child1 name="HAHA">1<child2>2</child2></child1></Request>
##########################################################################
def xml_tree3():
	sting = '<Request Action="UPDATE">'			\
			+ 	'<child1 name="HAHA">1'			\
			+ 		'<child2>2</child2>'		\
			+ 	'</child1>'						\
			+ '</Request>'
	result = ET.fromstring(sting)	

	print(result.tag)
	print(result.attrib)

def xml_parse():
	xfile = 'C:\Users\jeguan\Desktop\Test_2.txt'

	tree = ET.parse(xfile)
	root = tree.getroot()

	print(root.tag)
	print(root.attrib)



if __name__ == "__main__":
	xml_parse()

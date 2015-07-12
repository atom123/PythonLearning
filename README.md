# PythonLearning
Only work as a respository for Python Learning materials.

1. ElementTree.py
    This is used to show how to use the ElementTree in python. We can easily refer to this file to know how to use :
        book = xlrd.open_workbook(xlsfile)
        sheet0 = book.sheet_by_index(0)
        root = ElementTree.parse(xlsfile).getroot()
        SubAttrib = ET.SubElement(Attrib, "NODE1")
    and so on.

2. xlsread.py
    xls read and save.

3. re_test.py
    This is used to match a part of the text, this part can be in the same line or they can be from different line. "Lazzy match" is mentioned in this script.
    (.+?)----> this is Lazzy match.

    For this part, you can refer to my csdn blog: 


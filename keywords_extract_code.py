import xlwt
from xlwt import Workbook
from collections import OrderedDict

keywords=list()

# list of keywords to be extracted
keywords=['abstract', 'assert', 'boolean', 'break', 'byte', 'case', 'catch', 'char',
'class', 'const','continue', 'default',	'do','double','else', 'enum',
'extends', 'final', 'finally','float','for', 'goto', 'if', 'implements',
'import', 'instanceof',	'int',	'interface','long',	'native','new',	'package',
'private', 'protected', 'public', 'return',	'short', 'static',
'strictfp',	'super','switch', 'synchronized', 'this', 'throw',
'throws', 'transient', 'try', 'void', 'volatile',
'while', 'inheritance', 'encapsulation', 'multithreading']

key_count=dict()

# initialising count of each keyword with 0 in dictionary
for keyword in keywords:
    key_count[keyword]=0

# opening the file from where the keywords need to be extracted
fhandle = open("JavaBasics-notes.txt")

# counting the number of each keywords in the file
for line in fhandle:
    words=line.split()
    for word in words:
        if word in keywords:
            key_count[word]=key_count[word]+1

# sorting the items of dictionary by value i.e. there count or weighatge in ascending order
keywords_sorted_by_value = OrderedDict(sorted(key_count.items(), key=lambda x: x[1]))
print(keywords_sorted_by_value)

# Workbook is created
wb = Workbook()
# add_sheet is used to create sheet.
sheet1 = wb.add_sheet('Sheet 1')
# Specifying style
style = xlwt.easyxf('font: bold 1')

# column headers
sheet1.write(0, 0, 'KEYWORD', style)
sheet1.write(0, 1, 'WEIGHTAGE(COUNT)', style)

# writing into excel sheet and saving it
i=1
for k, v in keywords_sorted_by_value.items():
    sheet1.write(i, 0, k)
    sheet1.write(i, 1, v)
    i=i+1
wb.save('keywords_weightage.xls')

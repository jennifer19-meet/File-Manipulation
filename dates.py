import re
import openpyxl

regex_obj_1 = re.compile(r'-(\d){1,2}')
regex_obj_2 = re.compile(r'(\d){1,2}-')

x = openpyxl.load_workbook('book1.xlsx')
sheet = x.get_sheet_by_name('Sheet1')

num = 114

for i in range(1, num):
	place = "B"+str(i)
	if sheet[place].value != None:
		try:
			sheet[place] = regex_obj_1.sub('',str(sheet[place].value))
		except:
			print("could not do line "+ place)

for i in range(1, num):
	place = "C"+str(i)
	if sheet[place].value != None:
		try:
			sheet[place] = regex_obj_2.sub('',str(sheet[place].value))
		except:
			print("could not do line "+ place)

x.save('book1.xlsx')
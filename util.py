import openpyxl

import global_data

def read(file_name):
	book = openpyxl.load_workbook(filename=global_data.excel_path + file_name, read_only=True)
	for sheet in book:
		if sheet.title[0:1] == "_":
			break
		print sheet.title
	#print global_data.excel_path + file_name

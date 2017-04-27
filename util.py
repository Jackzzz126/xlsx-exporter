import openpyxl

import global_data

def read(file_name):
	book = openpyxl.load_workbook(filename=global_data.excel_path + file_name, read_only=True)
	for sheet in book:
		if sheet.title[0:1] == "_":
			break
		print sheet.title
	#print global_data.excel_path + file_name

def pos_index_2_str(row_index, col_index):
	col = col_index + 1
	colStr = ""

	while True:
		if col % 26 > 0:
			colStr = chr(ord("A") + col % 26 - 1) + colStr
			col = col / 26
		else:#col % 26 = 0
			if col == 0:
				break
			else:
				colStr = chr(ord("A") + 26 - 1) + colStr
				col = col / 26 - 1
	return colStr + str(row_index + 1)

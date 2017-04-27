import openpyxl

import global_data

class DataType(object):
	data_type = ""
	value_min = 1
	value_max = 0
	not_null = True
	id_type = ""
	def __init__(self, data_type, value_min, value_max, not_null=True, id_type=""):
		self.data_type = data_type
		self.value_min = value_min
		self.value_max = value_max
		self.not_null = not_null
		self.id_type = id_type

def read_book(file_name):
	book = openpyxl.load_workbook(filename=global_data.excel_path + file_name, read_only=True)
	for sheet in book:
		if sheet.title[0:1] == "_":
			break
		read_sheet(file_name, sheet)

def read_sheet(file_name, sheet):
	print file_name[0:-5] + ":" + sheet.title

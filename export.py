import os
import json
import openpyxl

import global_data
import util

class DataType(object):
	data_type = ""#int, string
	value_min = 1#int
	value_max = 0#int
	min_len = 0#string
	max_len = 0#string
	not_null = True
	is_array = False
	id_type = "id"#id, combineid

def run():
	items = os.listdir(global_data.excel_path)
	for item in items:
		if item[-5:] == ".xlsx" and item[0:2] != "~$":
			read_book_type(item)

	if global_data.gErrors > 0:
		return

def read_book_type(file_name):
	book = openpyxl.load_workbook(filename=global_data.excel_path + file_name, read_only=True)
	for sheet in book:
		if sheet.title[0:1] == "_":
			break
		read_sheet_type(file_name, sheet)

def read_sheet_type(file_name, sheet):
	keyFile = file_name[0:-5]
	keySheet = sheet.title
	if not keyFile in global_data.gTypes.keys():
		global_data.gTypes[keyFile] = {}
	if not keySheet in global_data.gTypes[keyFile].keys():
		global_data.gTypes[keyFile][keySheet] = {}

	col = 0
	while sheet[util.pos_index_2_str(0, col)].value != None:
		keyField = sheet[util.pos_index_2_str(0, col)].value
		util.log("Reading type %s:%s	%s" % (keyFile, keySheet, keyField))

		if keyField in global_data.gTypes[keyFile][keySheet].keys():
			util.add_error(keyFile, keySheet, 0, col, "Duplicate data field")

		global_data.gTypes[keyFile][keySheet][keyField] = read_type(sheet, col)

		col += 1

def read_type(sheet, col):
	typeDesc = json.loads(sheet[util.pos_index_2_str(3, col)].value)
	for k, v in typeDesc.items():
		print k
		print v

	return None


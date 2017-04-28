import os
import json
import openpyxl
import traceback

import global_data
import util

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

		global_data.gTypes[keyFile][keySheet][keyField] = read_type(keyFile, keySheet, sheet, col)

		col += 1

class DataType(object):
	data_type = ""
	min_value = 1
	max_value = 0
	min_len = 0
	max_len = 0
	not_null = True
	id_type = "id"
	is_array = False
	allowd_values = []

def read_type(keyFile, keySheet, sheet, col):
	try:
		type_desc = json.loads(sheet[util.pos_index_2_str(3, col)].value)
		data_type = DataType()
		if type_desc["dataType"] == "int":
			if not isinstance(type_desc["minValue"], int) or\
				not isinstance(type_desc["maxValue"], int) or\
				type_desc["minValue"] > type_desc["maxValue"] or\
				not isinstance(type_desc["notNull"], bool) or\
				not isinstance(type_desc["isArray"], bool) or\
				not isinstance(type_desc["allowedValues"], list):
				raise Exception("")

			for i in type_desc["allowedValues"]:
				if not isinstance(type_desc["allowedValues"][i], int):
					raise Exception("")

			if type_desc["idType"] != None and \
				type_desc["idType"] != "id" and \
				type_desc["idType"] != "combinedid":
				raise Exception("")

			data_type.data_type = type_desc["dataType"]
			data_type.min_value = type_desc["minValue"]
			data_type.max_value = type_desc["maxValue"]
			data_type.not_null = type_desc["notNull"]
			data_type.id_type = type_desc["idType"]
			data_type.is_array = type_desc["isArray"]
			data_type.allowd_values = type_desc["allowedValues"]
		elif type_desc["dataType"] == "float":
			if not isinstance(type_desc["minValue"], float) or\
				not isinstance(type_desc["maxValue"], float) or\
				type_desc["minValue"] > type_desc["maxValue"] or\
				not isinstance(type_desc["notNull"], bool) or\
				not isinstance(type_desc["isArray"], bool) or\
				not isinstance(type_desc["allowedValues"], list):
				raise Exception("")

			for i in type_desc["allowedValues"]:
				if not isinstance(type_desc["allowedValues"][i], float):
					raise Exception("")

			if type_desc["idType"] != None and \
				type_desc["idType"] != "id" and \
				type_desc["idType"] != "combinedid":
				raise Exception("")

			data_type.data_type = type_desc["dataType"]
			data_type.min_value = type_desc["minValue"]
			data_type.max_value = type_desc["maxValue"]
			data_type.not_null = type_desc["notNull"]
			data_type.id_type = type_desc["idType"]
			data_type.is_array = type_desc["isArray"]
			data_type.allowd_values = type_desc["allowedValues"]
		elif type_desc["dataType"] == "string":
			if not isinstance(type_desc["minLen"], int) or\
				not isinstance(type_desc["maxLen"], int) or\
				type_desc["minLen"] < 1 or\
				type_desc["minLen"] > type_desc["maxLen"] or\
				not isinstance(type_desc["notNull"], bool) or\
				not isinstance(type_desc["isArray"], bool) or\
				not isinstance(type_desc["allowedValues"], list):
				raise Exception("")

			for i in type_desc["allowedValues"]:
				if not isinstance(type_desc["allowedValues"][i], str):
					raise Exception("")

			if type_desc["idType"] != None and \
				type_desc["idType"] != "id" and \
				type_desc["idType"] != "combinedid":
				raise Exception("")

			data_type.data_type = type_desc["dataType"]
			data_type.min_len = type_desc["minLen"]
			data_type.max_len = type_desc["maxLen"]
			data_type.not_null = type_desc["notNull"]
			data_type.id_type = type_desc["idType"]
			data_type.is_array = type_desc["isArray"]
			data_type.allowd_values = type_desc["allowedValues"]
		else:
			raise Exception("")

	except Exception:
		print traceback.format_exc()
		util.add_error(keyFile, keySheet, 3, col, "Type describe error")
		return None

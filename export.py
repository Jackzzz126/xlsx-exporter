import os
import json
import traceback
import openpyxl

import global_data
import util

def run():
	items = os.listdir(global_data.excel_path)

	for item in items:
		if item[-5:] == ".xlsx" and item[0:2] != "~$":
			read_book_type(item)
	if global_data.g_errors > 0:
		return

	valid_type()
	if global_data.g_errors > 0:
		return

	for item in items:
		if item[-5:] == ".xlsx" and item[0:2] != "~$":
			pass
			#read_book_data(item)

def read_book_type(file_name):
	book = openpyxl.load_workbook(filename=global_data.excel_path + file_name, read_only=True)
	for sheet in book:
		if sheet.title[0:1] == "_":
			break
		read_sheet_type(file_name, sheet)

def read_book_data(file_name):
	book = openpyxl.load_workbook(filename=global_data.excel_path + file_name, read_only=True)
	for sheet in book:
		if sheet.title[0:1] == "_":
			break
		read_sheet_data(file_name, sheet)

def read_sheet_type(file_name, sheet):
	book_name = file_name[0:-5]
	sheet_name = sheet.title
	if not book_name in global_data.g_types.keys():
		global_data.g_types[book_name] = {}
	if not sheet_name in global_data.g_types[book_name].keys():
		global_data.g_types[book_name][sheet_name] = {}

	col = 0
	while sheet[util.pos_index_2_str(0, col)].value != None:
		field_name = sheet[util.pos_index_2_str(0, col)].value
		util.log("Reading type %s:%s	%s" % (book_name, sheet_name, field_name))

		if field_name in global_data.g_types[book_name][sheet_name].keys():
			util.add_error(book_name, sheet_name, 0, col, "Duplicate data field")

		global_data.g_types[book_name][sheet_name][field_name] = read_type(book_name, sheet_name, sheet, col)

		col += 1

def read_sheet_data(file_name, sheet):
	book_name = file_name[0:-5]
	sheet_name = sheet.title
	if not book_name in global_data.g_types.keys():
		global_data.g_datas[book_name] = {}
	if not sheet_name in global_data.g_types[book_name].keys():
		global_data.g_datas[book_name][sheet_name] = {}

	col = 0
	while sheet[util.pos_index_2_str(0, col)].value != None:
		field_name = sheet[util.pos_index_2_str(0, col)].value
		util.log("Reading type %s:%s	%s" % (book_name, sheet_name, field_name))

		if field_name in global_data.g_types[book_name][sheet_name].keys():
			util.add_error(book_name, sheet_name, 0, col, "Duplicate data field")

		global_data.g_types[book_name][sheet_name][field_name] = read_type(book_name, sheet_name, sheet, col)

		col += 1

class DataType(object):
	data_type = ""
	min_value = 1
	max_value = 0
	min_len = 0
	max_len = 0
	reg_exp = ""
	not_null = True
	id_type = ""
	is_array = False
	allowed_values = []
	ref = ""

def read_type(book_name, sheet_name, sheet, col):
	try:
		type_desc = json.loads(sheet[util.pos_index_2_str(3, col)].value)
		data_type = DataType()
		if type_desc["dataType"] == "ref":
			if not isinstance(type_desc["ref"], unicode):
				raise Exception("")

			data_type.ref = type_desc["ref"]

			return data_type

		elif type_desc["dataType"] == "int":
			if not isinstance(type_desc["minValue"], int) or\
				not isinstance(type_desc["maxValue"], int) or\
				type_desc["minValue"] > type_desc["maxValue"] or\
				("notNull" in type_desc.keys() and not isinstance(type_desc["notNull"], bool)) or\
				("isArray" in type_desc.keys() and not isinstance(type_desc["isArray"], bool)) or\
				("allowedValues" in type_desc.keys() and not isinstance(type_desc["allowedValues"], list)):
				raise Exception("")

			if "allowedValues" in type_desc.keys():
				for v in type_desc["allowedValues"]:
					if not isinstance(v, int):
						raise Exception("")

			if "idType" in type_desc.keys() and \
				type_desc["idType"] != "id" and \
				type_desc["idType"] != "combinedid":
				raise Exception("")

			data_type.data_type = type_desc["dataType"]
			data_type.min_value = type_desc["minValue"]
			data_type.max_value = type_desc["maxValue"]
			if "notNull" in type_desc.keys():
				data_type.not_null = type_desc["notNull"]
			if "idType" in type_desc.keys():
				data_type.id_type = type_desc["idType"]
			if "isArray" in type_desc.keys():
				data_type.is_array = type_desc["isArray"]
			if "allowedValues" in type_desc.keys():
				data_type.allowed_values = type_desc["allowedValues"]

			return data_type
		elif type_desc["dataType"] == "float":
			if not isinstance(type_desc["minValue"], float) or\
				not isinstance(type_desc["maxValue"], float) or\
				type_desc["minValue"] > type_desc["maxValue"] or\
				("notNull" in type_desc.keys() and not isinstance(type_desc["notNull"], bool)) or\
				("isArray" in type_desc.keys() and not isinstance(type_desc["isArray"], bool)) or\
				("allowedValues" in type_desc.keys() and not isinstance(type_desc["allowedValues"], list)):
				raise Exception("")

			if "allowedValues" in type_desc.keys():
				for v in type_desc["allowedValues"]:
					if not isinstance(v, float):
						raise Exception("")

			if "idType" in type_desc.keys() and \
				type_desc["idType"] != "id" and \
				type_desc["idType"] != "combinedid":
				raise Exception("")

			data_type.data_type = type_desc["dataType"]
			data_type.min_value = type_desc["minValue"]
			data_type.max_value = type_desc["maxValue"]
			if "notNull" in type_desc.keys():
				data_type.not_null = type_desc["notNull"]
			if "idType" in type_desc.keys():
				data_type.id_type = type_desc["idType"]
			if "isArray" in type_desc.keys():
				data_type.is_array = type_desc["isArray"]
			if "allowedValues" in type_desc.keys():
				data_type.allowed_values = type_desc["allowedValues"]

			return data_type
		elif type_desc["dataType"] == "string":
			if not isinstance(type_desc["minLen"], int) or\
				not isinstance(type_desc["maxLen"], int) or\
				type_desc["minLen"] < 1 or\
				type_desc["minLen"] > type_desc["maxLen"] or\
				("regExp" in type_desc.keys() and not isinstance(type_desc["regExp"], unicode)) or\
				("notNull" in type_desc.keys() and not isinstance(type_desc["notNull"], bool)) or\
				("isArray" in type_desc.keys() and not isinstance(type_desc["isArray"], bool)) or\
				("allowedValues" in type_desc.keys() and not isinstance(type_desc["allowedValues"], list)):
				raise Exception("")

			if "allowedValues" in type_desc.keys():
				for v in type_desc["allowedValues"]:
					if not isinstance(v, unicode):
						raise Exception("")

			if "idType" in type_desc.keys() and \
				type_desc["idType"] != "id" and \
				type_desc["idType"] != "combinedid":
				raise Exception("")

			data_type.data_type = type_desc["dataType"]
			data_type.min_len = type_desc["minLen"]
			data_type.max_len = type_desc["maxLen"]
			if "regExp" in type_desc.keys():
				data_type.reg_exp = type_desc["regExp"]
			if "notNull" in type_desc.keys():
				data_type.not_null = type_desc["notNull"]
			if "idType" in type_desc.keys():
				data_type.id_type = type_desc["idType"]
			if "isArray" in type_desc.keys():
				data_type.is_array = type_desc["isArray"]
			if "allowedValues" in type_desc.keys():
				data_type.allowed_values = type_desc["allowedValues"]

			return data_type
		else:
			raise Exception("")

	except Exception:
		print traceback.format_exc()
		util.add_error(book_name, sheet_name, 3, col, "Type describe error")
		return None

def valid_type():
	for book_name, sheetTypes in enumerate(global_data.g_types):
		for sheet_name, fieldTypes in sheetTypes:
			col = 0
			for field_name, data_type in fieldTypes:
				col += 1
				if data_type.data_type == "ref":
					ref_names = data_type.ref.split(":")
					if not ref_names[0] in global_data.g_types.keys() or\
						not ref_names[1] in global_data.g_types[ref_names[0]].keys():
						util.add_error(book_name, sheet_name, 3, col, "Ref not exist")

import os
import json
import re
import openpyxl

import util.global_data as global_data
import util.comm as comm

def run():
	items = os.listdir(global_data.excel_path)

	for item in items:
		if item[-5:] == ".xlsx" and item[0:2] != "~$":
			read_book_type(item)
	if global_data.g_errors:
		return

	valid_type()
	if global_data.g_errors:
		return

	for item in items:
		if item[-5:] == ".xlsx" and item[0:2] != "~$":
			read_book_data(item)
	if global_data.g_errors:
		return

	valid_refs()

def read_book_type(file_name):
	book = openpyxl.load_workbook(filename=global_data.excel_path + file_name, read_only=True)
	book_name = file_name[0:-5]
	for sheet in book:
		if sheet.title[0:1] == "_":
			break
		read_sheet_type(book_name, sheet)

def read_book_data(file_name):
	book = openpyxl.load_workbook(filename=global_data.excel_path + file_name, read_only=True)
	book_name = file_name[0:-5]
	for sheet in book:
		if sheet.title[0:1] == "_":
			break
		read_sheet_data(book_name, sheet)

def read_sheet_type(book_name, sheet):
	sheet_name = sheet.title
	if not book_name in global_data.g_types.keys():
		global_data.g_types[book_name] = {}
	if not sheet_name in global_data.g_types[book_name].keys():
		global_data.g_types[book_name][sheet_name] = {}

	col = 0
	while sheet[comm.pos_index_2_str(0, col)].value != None:
		field_name = sheet[comm.pos_index_2_str(0, col)].value
		comm.log("Reading type %s:%s	%s" % (book_name, sheet_name, field_name))

		if field_name in global_data.g_types[book_name][sheet_name].keys():
			comm.add_pos_error(book_name, sheet_name, 0, col, "Duplicate data field")

		global_data.g_types[book_name][sheet_name][field_name] = \
			read_type(sheet, col)

		col += 1

def read_sheet_data(book_name, sheet):
	sheet_name = sheet.title
	if not book_name in global_data.g_datas.keys():
		global_data.g_datas[book_name] = {}
	if not sheet_name in global_data.g_datas[book_name].keys():
		global_data.g_datas[book_name][sheet_name] = {}

	row = 4
	while sheet[comm.pos_index_2_str(row, 0)].value != None:
		key = ""
		line_data = {}

		#get key
		col = 0
		while sheet[comm.pos_index_2_str(0, col)].value != None:
			field_name = sheet[comm.pos_index_2_str(0, col)].value
			data_type = global_data.g_types[book_name][sheet_name][field_name]
			pos_str = comm.pos_index_2_str(row, col)
			if data_type.id_type == "id":
				key = str(sheet[pos_str].value)
				break
			elif data_type.id_type == "combineid":
				if key == "":
					key = str(sheet[pos_str].value)
				else:
					key += "_" + str(sheet[pos_str].value)
			else:
				pass
			col += 1

		col = 0
		while sheet[comm.pos_index_2_str(0, col)].value != None:
			field_name = sheet[comm.pos_index_2_str(0, col)].value
			pos_str = comm.pos_index_2_str(row, col)
			comm.log("Reading data %s:%s	%s" % (book_name, sheet_name, pos_str))
			cell_data = read_data(book_name, sheet_name, field_name, row, col, sheet[pos_str].value)
			line_data[field_name] = cell_data

			col += 1

		if key == "":
			if not global_data.g_datas[book_name][sheet]:
				global_data.g_datas[book_name][sheet] = []
			global_data.g_datas[book_name][sheet].append(line_data)
		else:
			if not global_data.g_datas[book_name][sheet_name]:
				global_data.g_datas[book_name][sheet] = {}
			global_data.g_datas[book_name][sheet][key] = line_data

		row += 1

class DataType(object):
	data_type = ""
	min_value = 1
	max_value = 0
	min_len = 0
	max_len = 0
	reg_exp = ""
	id_type = ""
	is_array = False
	allowed_values = []
	ref = ""

def read_type(sheet, col):
	type_desc = json.loads(sheet[comm.pos_index_2_str(3, col)].value)
	data_type = DataType()
	if type_desc["dataType"] == "ref":
		if not isinstance(type_desc["ref"], unicode):
			raise Exception("")

		data_type.data_type = type_desc["dataType"]
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
			type_desc["idType"] != "combineid":
			raise Exception("")

		data_type.data_type = type_desc["dataType"]
		data_type.min_value = type_desc["minValue"]
		data_type.max_value = type_desc["maxValue"]
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
		if "idType" in type_desc.keys():
			data_type.id_type = type_desc["idType"]
		if "isArray" in type_desc.keys():
			data_type.is_array = type_desc["isArray"]
		if "allowedValues" in type_desc.keys():
			data_type.allowed_values = type_desc["allowedValues"]

		return data_type
	else:
		raise Exception("")

def valid_type():
	for book_name in global_data.g_types:
		for sheet_name in global_data.g_types[book_name]:
			id_count = 0
			dup_field_name = ""
			for field_name in global_data.g_types[book_name][sheet_name]:
				data_type = global_data.g_types[book_name][sheet_name][field_name]

				# check ref
				if data_type.data_type == "ref":
					ref_names = data_type.ref.split(":")

					if not ref_names[0] in global_data.g_types.keys() or\
						not ref_names[1] in global_data.g_types[ref_names[0]].keys():
						comm.add_field_error(book_name, sheet_name, 3, field_name, "Ref not exist")

					ref_has_id = False
					for ref_field_name in global_data.g_types[ref_names[0]][ref_names[1]]:
						ref_data_type = global_data.g_types[ref_names[0]][ref_names[1]][ref_field_name]
						if ref_data_type.id_type != "":
							ref_has_id = True
							break
					if not ref_has_id:
						comm.add_field_error(book_name, sheet_name, 3, field_name, "Ref has no id")

				# id field count
				if data_type.id_type == "id":
					id_count += 1
					dup_field_name = field_name

			if id_count > 1:
				comm.add_field_error(book_name, sheet_name, 3, dup_field_name, "Multi id field")

def read_data(book_name, sheet_name, field_name, row, col, raw_value):
	data_type = global_data.g_types[book_name][sheet_name][field_name]

	if data_type.data_type == "int":
		def valid_int(value_str):
			value = int(value_str)
			if value < data_type.min_value or value > data_type.max_value:
				comm.add_pos_error(book_name, sheet_name, row, col, "Int data out of range")

			if data_type.allowed_values:
				in_allowed = False
				for i in range(0, len(data_type.allowed_values)):
					if data_type.allowed_values[i] == value:
						in_allowed = True
						break
				if not in_allowed:
					comm.add_pos_error(book_name, sheet_name, row, col, "Int data not in allowed values")

		if raw_value is None:
			return 0
		if data_type.is_array:
			values = []
			value_strs = str(raw_value).split(",")
			if not value_strs:
				comm.add_pos_error(book_name, sheet_name, row, col, "Array length is zero")
				return values
			for i in range(0, len(value_strs)):
				if not value_strs[i]:
					comm.add_pos_error(book_name, sheet_name, row, col, "Array member is null")
					continue
				valid_int(value_strs[i])
				values.append(int(value_strs[i]))
			return values
		else:
			valid_int(str(raw_value))
			return int(raw_value)
	elif data_type.data_type == "float":
		def valid_float(value_str):
			value = float(value_str)
			if value < data_type.min_value or value > data_type.max_value:
				comm.add_pos_error(book_name, sheet_name, row, col, "Float data out of range")

			if data_type.allowed_values:
				in_allowed = False
				for i in range(0, len(data_type.allowed_values)):
					if (data_type.allowed_values[i] - value) ** 2 < 0.00001:
						in_allowed = True
						break
				if not in_allowed:
					comm.add_pos_error(book_name, sheet_name, row, col, "Float data not in allowed values")

		if raw_value is None:
			return 0.0
		if data_type.is_array:
			value_strs = str(raw_value).split(",")
			values = []
			if not value_strs:
				comm.add_pos_error(book_name, sheet_name, row, col, "Array length is zero")
				return values
			for i in range(0, len(value_strs)):
				if not value_strs[i]:
					comm.add_pos_error(book_name, sheet_name, row, col, "Array member is null")
					continue
				valid_float(value_strs[i])
				values.append(float(value_strs[i]))
			return values
		else:
			valid_float(str(raw_value))
			return float(raw_value)
	elif data_type.data_type == "string":
		def valid_str(value_str):
			if len(value_str) < data_type.min_len or len(value_str) > data_type.max_len:
				comm.add_pos_error(book_name, sheet_name, row, col, "String length out of range")

			if data_type.allowed_values:
				in_allowed = False
				for i in range(0, len(data_type.allowed_values)):
					if data_type.allowed_values[i] == value_str:
						in_allowed = True
						break
				if not in_allowed:
					comm.add_pos_error(book_name, sheet_name, row, col, "String data not in allowed values")

			if data_type.reg_exp:
				pattern = re.compile(data_type.reg_exp)
				match = pattern.match(value_str)
				if not match or match.group() != value_str:
					comm.add_pos_error(book_name, sheet_name, row, col, "String regular expression valid faile")

		if raw_value is None:
			return ""
		if data_type.is_array:
			value_strs = str(raw_value).split(",")
			values = []
			if not value_strs:
				comm.add_pos_error(book_name, sheet_name, row, col, "Array length is zero")
				return values
			for i in range(0, len(value_strs)):
				if not value_strs[i]:
					comm.add_pos_error(book_name, sheet_name, row, col, "Array member is null")
					continue
				valid_str(value_strs[i])
				values.append(value_strs[i])
			return values
		else:
			valid_str(str(raw_value))
			return str(raw_value)
	elif data_type.data_type == "ref":
		def valid_ref(value_str):
			if data_type.allowed_values:
				in_allowed = False
				for i in range(0, len(data_type.allowed_values)):
					if data_type.allowed_values[i] == value_str:
						in_allowed = True
						break
				if not in_allowed:
					comm.add_pos_error(book_name, sheet_name, row, col, "Refence data not in allowed values")
		if raw_value is None:
			return ""
		if data_type.is_array:
			value_strs = str(raw_value).split(",")
			values = []
			if not value_strs:
				comm.add_pos_error(book_name, sheet_name, row, col, "Array length is zero")
				return values
			for i in range(0, len(value_strs)):
				if not value_strs[i]:
					comm.add_pos_error(book_name, sheet_name, row, col, "Array member is null")
					continue
				valid_ref(value_strs[i])
				values.append(value_strs[i])
			return values
		else:
			valid_ref(str(raw_value))
			return str(raw_value)
	else:
		comm.add_pos_error(book_name, sheet_name, row, col, "Unknown data type")
		return None

def valid_refs():
	for book_name in global_data.g_types:
		for sheet_name in global_data.g_types[book_name]:
			for field_name in global_data.g_types[book_name][sheet_name]:
				data_type = global_data.g_types[book_name][sheet_name][field_name]

				if data_type.data_type != "ref":
					return

				if isinstance(global_data.g_datas[book_name][sheet_name], dict):
					for key in global_data.g_datas[book_name][sheet_name]:
						value = global_data.g_datas[book_name][sheet_name][key][field_name]
						if isinstance(value, list):
							for i in range(0, len(value)):
								valid_ref_value_key(key, value[i], book_name, sheet_name, field_name)
						else:
							valid_ref_value_key(key, value, book_name, sheet_name, field_name)
				elif isinstance(global_data.g_datas[book_name][sheet_name], list):
					for i in range(0, len(global_data.g_datas[book_name][sheet_name])):
						value = global_data.g_datas[book_name][sheet_name][i][field_name]
						if isinstance(value, list):
							for j in range(0, len(value)):
								valid_ref_value_index(i, value[j], book_name, sheet_name, field_name)
						else:
							valid_ref_value_index(i, value, book_name, sheet_name, field_name)

	def valid_ref_value_key(key, value, book_name, sheet_name, field_name):
		data_type = global_data.g_types[book_name][sheet_name][field_name]
		ref_names = data_type.ref.split(":")
		if value not in global_data.g_datas[ref_names[0]][ref_names[1]].keys():
			comm.add_key_error(book_name, sheet_name, key, field_name, "Ref not exist")

	def valid_ref_value_index(index, value, book_name, sheet_name, field_name):
		data_type = global_data.g_types[book_name][sheet_name][field_name]
		ref_names = data_type.ref.split(":")
		if value not in global_data.g_datas[ref_names[0]][ref_names[1]].keys():
			comm.add_field_error(book_name, sheet_name, index + 4, field_name, "Ref not exist")

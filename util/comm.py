import time

from util import global_data

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

def log(log_str):
	time_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
	print time_str + " [LOG] " + log_str

def error(log_str):
	time_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
	print time_str + " [ERR] " + log_str

class Error(object):
	file_name = ""
	sheet_name = ""
	pos = ""
	err_msg = ""
	def __init__(self, file_name, sheet_name, pos, err_msg=""):
		self.file_name = file_name
		self.sheet_name = sheet_name
		self.pos = pos
		self.err_msg = err_msg

def add_pos_error(file_name, sheet_name, row, col, err_msg=""):
	err = Error(file_name, sheet_name, pos_index_2_str(row, col), err_msg)
	global_data.g_errors.append(err)

def add_field_error(file_name, sheet_name, row, field_name, err_msg=""):
	err = Error(file_name, sheet_name, field_name + ":" + str(row + 1), err_msg)
	global_data.g_errors.append(err)

def add_key_error(file_name, sheet_name, key, field_name, err_msg=""):
	err = Error(file_name, sheet_name, field_name + ":" + key, err_msg)
	global_data.g_errors.append(err)

def write_errors():
	log_file = open(global_data.excel_path + 'errors.txt', 'w')
	time_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
	log_file.write(time_str + "\n")

	for _, v in enumerate(global_data.g_errors):
		error("%s:%s %s	%s" % (v.file_name, v.sheet_name, v.pos, v.err_msg))
		log_file.write("%s:%s %s	%s\n" % (v.file_name, v.sheet_name, v.pos, v.err_msg))

	log_file.close()

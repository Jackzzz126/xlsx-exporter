#! /usr/bin/env python
# -*- coding: utf-8 -*-

import os
import xlrd
import codecs
import pdb
import json
import time

def write_log(log):
	log_file = open('error_log.txt', 'a')
	log_file.write(log + "\n")
	log_file.close()

def parseType(typeStr):
	typeStr=typeStr.lower()

	left = typeStr.find("(")
	right = typeStr.find(")")
	para = [int(eval(n)) for n in typeStr[left+1:right].split(",")]

	typeName = typeStr[0:left]
	if typeName != "int"\
		and typeName != "char"\
		and typeName != "bool"\
		and typeName != "array"\
		:
		raise Exception("error type %s" % typeName)

	try:
		if typeName == "int":
			if para[0] >= para[1]:
				raise Exceptionty("error range %s" % typeStr[left:right+1])
		elif typeName == "char":
			if para[0] < 1:
				raise Exceptionty("error range %s" % typeStr[left:right+1])
		elif typeName == "array":
			if para[0] < 1:
				raise Exceptionty("error range %s" % typeStr[left:right+1])
		else:#bool or other
			pass
	except Exception, e:
		raise Exception("error range %s" % typeStr[left:right+1])

	return typeName, para

def validType(valueObj, typePair):
	typeName = typePair[0]
	if typeName == "int":
		try:
			value = int(valueObj)
			if value < typePair[1][0] or value > typePair[1][1]:
				return False, None
			else:
				return True, value
		except Exception, e:
			return False, None
	elif typeName == "char":
		value = unicode(valueObj)
		if len(value) > typePair[1][0]:
			return False, None
		else:
			return True, value.encode('utf-8')
	elif typeName == "array":
		value = str(valueObj).split(",")
		if value[0] == "":
			return True, [0]
		valueList = [int(eval(num)) for num in value]
		listLen = len(valueList)
		listDimension = typePair[1][0]
		for i in range(0, listLen, listDimension):
			for j in range(0, listDimension):
				if valueList[i + j] < typePair[1][j*2 + 1] or valueList[i + j] > typePair[1][j*2 + 2]:
					return False, None
		if listLen % listDimension != 0:
			return False, None
		else:
			return True, valueList
	elif typeName == "bool":
		try:
			value = int(valueObj)
			if value == 1:
				return True, True
			elif value == 0:
				return True, False
			else:
				return False, None
		except Exception, e:
			return False, None

def logError(bookName, sheetName, row, col, errorStr = None):
	colStr = chr(ord("A") + col % 26)
	col = col / 26
	while(True):
		if col > 0:
			colStr = chr(ord("A") + colStr % 26) + colStr
			col = col / 26
		else:
			break
	if errorStr == None:
		errorStr = "invalid value"
	write_log("%s:	%s:	[%d][%s]:	%s." % (bookName, sheetName, row+1, colStr, errorStr))

def write_js1(dict1, name):
	js_file = open('config.js', 'a')
	js_file.write(name + " = {\n")
	for key in dict1:
		data = json.dumps(dict1[key],sort_keys=True)
		js_file.write("\t\""+key + "\":" + data + ",\n")

	js_file.write("};\n")
	js_file.write("exports."+name+" = "+name+";\n")
	js_file.close()

def write_js2(dict1,name):
	js_file = open(name + '.js', 'w')
	js_file.write(name + " = {\n")
	for key1 in dict1:
		dict2=dict1[key1]
		js_file.write("\t\""+ key1 + "\":{\n")
		for key2 in dict2:
			data = json.dumps(dict2[key2],sort_keys=True)
			js_file.write("\t\t\""+key2 + "\":" + data + ",\n")
		js_file.write("\t},\n")
	js_file.write("};\n")
	js_file.write("exports."+name+" = "+name+";\n")
	js_file.close()
def write_js(dict1,name):
	js_file = open(name + '.js', 'w')
	js_file.write(name + " =")
	js_file.write(json.dumps(dict1, sort_keys=True, indent=4, encoding="utf-8", ensure_ascii=False))
	js_file.write(";\n")
	js_file.write("exports."+name+" = "+name+";\n")
	js_file.close()

def convert(file_name):
	book = xlrd.open_workbook(file_name)
	for sheet in book.sheets():
		keyList = []
		for col in xrange(0, sheet.ncols):
			keyList.append(str(sheet.cell_value(1, col)))

		typeList = []
		for col in xrange(0, sheet.ncols):
			typeList.append(parseType(str(sheet.cell_value(2, col))))

		jsDict = {}
		idDict = {}
		for row in xrange(3, sheet.nrows):
			idStr = str(int(sheet.cell_value(row, 0)))
			if idStr == "0":
				logError(file_name, sheet.name, row, 0)
				continue
			else:
				jsDict[idStr] = {}
			if idDict.get(idStr) != None:
				logError(file_name, sheet.name, row, 0, "duplicate id")
			else:
				idDict[idStr] = idStr
				

			for col in xrange(0, sheet.ncols):
				valuePair = validType(sheet.cell_value(row, col), typeList[col])
				if not valuePair[0]:
					logError(file_name, sheet.name, row, col)
				else:
					jsDict[idStr][keyList[col]] = valuePair[1]
		write_js(jsDict, sheet.name)

#! /usr/bin/env python
# -*- coding: utf-8 -*-

import traceback
import os
import argparse

import export
import global_data
import util

def main():
	"""main function"""
	try:
		#parse args
		parser = argparse.ArgumentParser(description="Excel exporter")
		parser.add_argument('-p', dest="path", help='path of the excel files, default "./excel"')
		args = parser.parse_args()
		if args.path:
			if args.path[-1:] != "/":
				args.path += "/"
			global_data.excel_path = args.path

		if os.path.isfile(global_data.excel_path + "errors.txt"):
			os.remove(global_data.excel_path + "errors.txt")
		export.run()
		if global_data.g_errors:
			util.write_errors()
			print "Export finished with error. Please check 'error.txt' for details."
		else:
			print "Export finished success."
		raw_input("Press the <ENTER> key to continue...")
	except Exception, ex:
		print traceback.format_exc()
		print ex

if __name__ == "__main__":
	main()

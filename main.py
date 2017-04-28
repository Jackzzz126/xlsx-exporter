#! /usr/bin/env python
# -*- coding: utf-8 -*-

import traceback
import os

import export
import global_data
import util

def main():
	"""main function"""
	try:
		if os.path.isfile(global_data.excel_path + "errors.txt"):
			os.remove(global_data.excel_path + "errors.txt")
		export.run()
		if global_data.gErrors:
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

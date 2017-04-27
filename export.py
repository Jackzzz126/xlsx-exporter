#! /usr/bin/env python
# -*- coding: utf-8 -*-

"""
main log
"""

import os
import time
import traceback

import global_data
import util

def main():
	"""main function"""
	try:
		log_file = open('errors.txt', 'w')
		time_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
		log_file.write(time_str + "\n")
		log_file.close()

		items = os.listdir(global_data.excel_path)
		for item in items:
			if item[-5:] == ".xlsx":
				util.read(item)

	except Exception, ex:
		print traceback.format_exc()
		print ex

#	print "Covert to json finished. Please check 'error.txt' for errors."
if __name__ == "__main__":
	main()

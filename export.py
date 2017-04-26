#! /usr/bin/env python
# -*- coding: utf-8 -*-

"""
main log
"""

import os
import time
import traceback

import data
import util

if __name__ == "__main__":
	try:
		log_file = open( 'errors.txt', 'w' )
		time_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
		log_file.write(time_str + "\n")
		log_file.close()

		items = os.listdir(data.EXCEL_PATH)
		for item in items:
			if item[-5:] == ".xlsx":
				util.read(item)

	except Exception, e:
		print traceback.format_exc()
		print e;

#	print "Covert to json finished. Please check 'error.txt' for errors."

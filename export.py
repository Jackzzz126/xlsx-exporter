#! /usr/bin/env python
# -*- coding: utf-8 -*-

import os
import time
import traceback

if __name__ == "__main__":
	excel_path = "./excel/"
	try:
		log_file = open( 'errors.txt', 'w' )
		time_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
		log_file.write(time_str + "\n")
		log_file.close()

		items = os.listdir(excel_path)
		for item in items:
			if item[-5:] == ".xlsx":
				print item
				#util.export( item )

	except Exception, e:
		print traceback.format_exc()
		print e;

	print "Covert to json finished. Please check 'error.txt' for errors."

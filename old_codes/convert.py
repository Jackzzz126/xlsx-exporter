#! /usr/bin/env python
# -*- coding: utf-8 -*-

import os
import xlrd
#from xml.dom import minidom
import codecs
import pdb
import json
import time

import convert_utility as utility
#from mission_convert import convert_mission
#from chapter_convert import convert_chapter
#from hero_convert import convert_hero
#from skill_convert import convert_skill
#from bullet_convert import convert_bullet

if __name__ == "__main__":
	try:
		log_file = open( 'error_log.txt', 'w' )
		time_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
		log_file.write(time_str + "\n")
		log_file.close( )

		items = os.listdir( "./" )
		for item in items:
			if item[-4:] == ".xls":
				utility.convert( item )

	except Exception, e:
		import traceback
		print traceback.format_exc()
		print e;

	print "covert to json finished. Please see 'error_log.txt' for errores."
	os.system("pause")

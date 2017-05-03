#! /usr/bin/env python
# -*- coding: utf-8 -*-

import unittest
from util import comm

class UtilTest(unittest.TestCase):
	def setUp(self):
		pass

	def tearDown(self):
		pass

	def test_pos_index_2_str(self):
		self.assertEqual(comm.pos_index_2_str(0, 1 - 1), 'A1', 'pos_index_2_str fail')
		self.assertEqual(comm.pos_index_2_str(9, 26 - 1), 'Z10', 'pos_index_2_str fail')
		self.assertEqual(comm.pos_index_2_str(999, 26 ** 1 * 1 + 26 - 1),\
				'AZ1000', 'pos_index_2_str fail')
		self.assertEqual(comm.pos_index_2_str(0, 26 ** 1 * 2 + 1- 1), 'BA1', 'pos_index_2_str fail')
		self.assertEqual(comm.pos_index_2_str(0, 26 ** 1 * 26 + 2 - 1), 'ZB1', 'pos_index_2_str fail')
		self.assertEqual(comm.pos_index_2_str(0, 26 ** 2 * 1 + 26 ** 1 * 26 + 2 - 1),\
				'AZB1', 'pos_index_2_str fail')
		self.assertEqual(comm.pos_index_2_str(0, 26 ** 2 * 3 + 26 ** 1 * 26 + 2 - 1),\
				'CZB1', 'pos_index_2_str fail')

if __name__ == '__main__':
	unittest.main()

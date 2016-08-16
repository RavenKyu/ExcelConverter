import unittest
from xlsx_handler import xlrd_handler
from xlsx_handler import openpyxl_handler
__author__ = 'raven'


class handler_test(unittest.TestCase):
	wb_xlrd = xlrd_handler("test.xls")
	wb_openpyxl = openpyxl_handler("text.xlsx")

	def test_get_sheet_names(self):
		self.assertIsInstance(self.wb_xlrd.get_sheet_names()['retval'], list)
		self.assertIsInstance(self.wb_openpyxl.get_sheet_names()['retval'], list)
		self.assertEqual(self.wb_xlrd.get_sheet_names()['retval'], 
		                 self.wb_openpyxl.get_sheet_names()['retval'])

	def test_get_sheet_by_index(self):
		self.assertIsInstance(self.wb_xlrd.get_sheet_by_index(0)['retval'].title, "Sheet1")

	def test_get_sheet_by_name(self):
		self.assertIsInstance(self.wb_xlrd.get_sheet_by_name("Sheet1")['retval'].title, "Sheet1")

	def test_get_row_number(self):
		self.assertEqual(self.wb_xlrd.get_row_number()['retval'], 11)

	def test_get_col_number(self):
		self.assertEqual(self.wb_xlrd.get_col_number()['retval'], 14)

	def test_get_row_values(self):
		self.assertIsInstance(self.wb_xlrd.get_row_values(4)['retval'], list)

	def test_get_cell_value(self):
		self.assertEqual(self.wb_xlrd.get_cell_value(4, 8)['retval'], "SSLVPN")


# class openpyxl_handler_test(unittest.TestCase):
# 	wb = openpyxl_handler("test.xlsx")
#
# 	def test_get_sheet_by_name(self):
# 		self.assertEqual(self.wb.get_sheet_by_name("Sheet1")['retval'].title, "Sheet1")
#
# 	def test_get_sheet_by_index(self):
# 		self.assertEqual(self.wb.get_sheet_by_index(0)['retval'].title, "Sheet1")
#
# 	def test_get_sheet_name(self):
# 		self.assertIsInstance(self.wb.get_sheet_names()['retval'], list)
# 		self.assertEqual(self.wb.get_sheet_names()['retval'], [u'Sheet1', u'Sheet2', u'Sheet3'])
#
# 	def test_get_row_number(self):
# 		self.assertEqual(self.wb.get_row_number()['retval'], 11)
#
# 	def test_get_col_number(self):
# 		self.assertEqual(self.wb.get_col_number()['retval'], 14)
#
# 	def test_get_row_values(self):
# 		self.assertIsInstance(self.wb.get_row_values(4)['retval'], list)
#
# 	def test_get_cell_value(self):
# 		self.assertEqual(self.wb.get_cell_value(4, 8)['retval'], "SSLVPN")
#

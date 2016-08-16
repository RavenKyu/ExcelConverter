import unittest
from xlsx_handler import xlrd_handler
from xlsx_handler import openpyxl_handler
__author__ = 'raven'


class HandlerValueMatchingTest(unittest.TestCase):
	wb_xlrd = xlrd_handler("test.xls")
	wb_xlrd.load_workbook()
	wb_openpyxl = openpyxl_handler("test.xlsx")
	wb_openpyxl.load_workbook()

	def test_get_sheet_names(self):
		self.assertIsInstance(self.wb_xlrd.get_sheet_names()['retval'], list)
		self.assertIsInstance(self.wb_openpyxl.get_sheet_names()['retval'], list)
		self.assertEqual(self.wb_xlrd.get_sheet_names()['retval'], 
		                 self.wb_openpyxl.get_sheet_names()['retval'])

	def test_get_sheet_by_index(self):
		self.assertEqual(self.wb_xlrd.get_sheet_by_index(0)['retval'].name,
		                 self.wb_openpyxl.get_sheet_by_index(0)['retval'].title)

	def test_get_sheet_by_name(self):
		self.assertEqual(self.wb_xlrd.get_sheet_by_name("Sheet1")['retval'].name,
		                 self.wb_openpyxl.get_sheet_by_name("Sheet1")['retval'].title)

	def test_get_row_number(self):
		self.assertEqual(self.wb_xlrd.get_row_number()['retval'],
		                 self.wb_openpyxl.get_row_number()['retval'])

	def test_get_col_number(self):
		self.assertEqual(self.wb_xlrd.get_col_number()['retval'],
		                 self.wb_openpyxl.get_col_number()['retval'])

	def test_get_row_values(self):
		self.assertEqual(self.wb_xlrd.get_row_values(4)['retval'],
		                 self.wb_openpyxl.get_row_values(4)['retval'])

	def test_get_cell_value(self):
		self.assertEqual(self.wb_xlrd.get_cell_value(4, 8)['retval'],
		                 self.wb_openpyxl.get_cell_value(4, 8)['retval'])



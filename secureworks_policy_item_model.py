# coding: utf-8

from DataStructure.item_future_zen_policy import FutureZenPolicyItemModel
from DataStructure.item_future_zen_policy import FuturePolicies

from xlsx_handler import ExcelHandler

# class ItemModel(FutureZenPolicyItemModel):
# 	def __init__(self):
# 		FutureZenPolicyItemModel.__init__(self)


class SecrureWorks:
	def __init__(self, filename):
		self.hd = ExcelHandler(model=None, filename=filename)
		self.hd.load()

		self.policies = FuturePolicies()

	def get_items(self):
		for row_idx in xrange(4, self.hd.end_row):
			print self.hd.get_row_values(row_idx)['retval']





SecrureWorks("test.xls").get_items()



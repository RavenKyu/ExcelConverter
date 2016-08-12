# coding: utf-8
__author__ = "임덕규"

class ReturnValue(object):
	def __init__(self, state, retval):
		if not isinstance(state, bool):
			raise TypeError
		self._state = state
		self._retval = retval
		self._rettype = type(self._retval)

	def __repr__(self):
		return {"state": self._state, "retval": self._retval, "rettype": self._rettype}

	def __getitem__(self, item):
		return getattr(self, "_"+item)

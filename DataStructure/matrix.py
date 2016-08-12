# coding: utf-8
__author__ = "임덕규"

class Matrix(list):
	"""
	이중 배열을 이용하여 데이터를 쉽게 원하는 좌표에 붙여 넣을 수 있게 도와주는
	리스트 형태의 데이터 타입.
	"""
	def __init__(self):
		super(Matrix, self).__init__(self)

	def set(self, row, col, value):
		"""
		:param row 정수형, 세로 좌표
		:param col 정수형, 가로 좌표
		:param value 데이터

		원하는 좌표에 value를 삽입.
		존재하지 않는 위치라면 확장을 하여 해당 좌표까지 이동.
		사용자는 해당 배열의 존재 유무를 알 필요 없이 사용이 가능.
		>>> m = Matrix()
		>>> m.set(3, 4, "Hello World")
		# [[], [], [], [None, None, None, None, 'Hello World']]
		"""
		while True:
			try:
				self[row]
			except IndexError:
				self.append([])
				continue

			try:
				self[row][col] = value
				break
			except IndexError:
				self[row].append(None)
				continue
		return

	def set_matrix(self, row, col, matrix, ignore=None):
		"""

		"""

		if not isinstance(matrix, (Matrix, list)):
			raise TypeError
		for r, row_data in enumerate(matrix):
			for c, col_data in enumerate(row_data):
				if col_data == ignore: continue
				self.set(row + r, col + c, col_data)
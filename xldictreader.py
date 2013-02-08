import xlrd, xlwt, xlutils
from itertools import izip_longest as zip_longest

class ExcelDictReader(file):
  
	excel_file 	= None
	dict_reader = None

	def __init__(self, file):
		return_list = []
		self.excel_file = xlrd.open_workbook(file)
		for sheet in self._split_sheets():
			return_list.append(self._sheet_to_dict_list(sheet))
		self.dict_reader = return_list

	def _split_sheets(self):
		return self.excel_file.sheets()

	def _sheet_dict_keys(self, sheet):
		keys = sheet.row_values(0)
		return keys

	def _sheet_row_to_dict(self,row,keys):
		return {k:r for k,r in zip_longest(keys,row)}

	def _sheet_to_dict_list(self,sheet):
		dict_list = []
		keys = self._sheet_dict_keys(sheet)
		for i in range(1,sheet.nrows):
			row = sheet.row_values(i)
			dict_list.append(self._sheet_row_to_dict(row, keys))
		return dict_list

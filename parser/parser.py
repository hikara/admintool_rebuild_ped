import xlrd as pd


class ParseExcelFile:
	def __init__(self):
		self.lookup = {}
		self.indexes = []
		self.arm_data_fields = ["armofficialname", "armpathwayid", "armgene", "armid", "armpathwayname", "armdescription", "armdrug", "armdrugid", "version", "stratum_id"]
		self.top_level_fields = ["histologicdiseaseexclusioncodes", "priortherapy(drugexclusion)", "ihcresults", "non-hotspotrules", "exclusionvariants", "inclusionvariants"]

	def set_index(self, name, index):
		self.lookup[name] = index
		self.lookup[index] = name
		self.indexes.append(index)

	def set_indexes(self, sheet):
		items = sheet.col(0)
		for index in range(0, len(items)):
			if isinstance(items[index].value, str):
				name = items[index].value.lower().replace(" ", "")
				if name in self.top_level_fields:
					self.set_index(name, index)

		self.indexes.append(len(sheet.col(0)) - 1)
		self.indexes.sort()

	@staticmethod
	def check_for_empty_rows(sheet, start, end):
		rows = list(sheet.get_rows())
		for index in range(start, end):
			empty_items = 0
			for item in rows[index]:
				if item.value is None or item.value == '':
					empty_items += 1
			if empty_items == sheet.ncols:
				end = index
				break
		return end

	def set_subsection(self, sheet, parsing_dict, name, start, end):
		parsing_dict[name] = {}
		for index in range(0, sheet.ncols):
			inner_name = sheet.col_values(index)[start + 1].lower().replace(" ", "")
			if inner_name != '':
				parsing_dict[name][inner_name] = sheet.col_values(index)[start + 2:self.check_for_empty_rows(sheet, start, end)]

	def set_arm_data(self, sheet, parsing_dict, end):
		for i in range(0, end):
			for j in range(sheet.ncols):
				if isinstance(sheet.cell_value(i, j), str) and sheet.cell_value(i, j).lower().replace(" ", "") in self.arm_data_fields:
					parsing_dict[sheet.cell_value(i, j).lower().replace(" ", "")] = sheet.cell_value(i, j + 1)


	def parse_workbook(self, local_excel_file_path, excel_sheet_name):
		parsing_dict = {}
		excel_file = pd.open_workbook(local_excel_file_path)
		sheet = excel_file.sheet_by_name(excel_sheet_name)
		self.set_indexes(sheet)
		self.set_arm_data(sheet, parsing_dict, self.indexes[0])
		for index in range(0, len(self.indexes) - 1):
			self.set_subsection(sheet, parsing_dict, self.lookup[self.indexes[index]], self.indexes[index], self.indexes[index + 1])
		print(parsing_dict)

if __name__ == '__main__':
	ParseExcelFile().parse_workbook("sample.xlsx", "APEC1621A")

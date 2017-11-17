# from xlrd import open_workbook
# # from common.dynamoDB_formatter import DynamoDBFormatter
# # from adult.adult_dynamoDB_formatter import AdultDynamoDBFormatter
# # from treatment_arm_json_validator.treatment_arm_json_validator import TreatmentArmJSONValidator
# # import json
#
#
# class ParseExcelFile:
#     def __init__(self):
#         pass
#
#     def loop_through_end_parsing_index(self, sheet, parsing_dict, i, j, end_parsing_index):
#         if end_parsing_index is not None:
#             for r in range(i + 2, end_parsing_index):
#                 if not self.check_for_empty_row(sheet, r):
#                     self.try_looping_through_end_parsing_index(sheet, parsing_dict, i, j, end_parsing_index, r)
#
#     @staticmethod
#     def try_looping_through_end_parsing_index(sheet, parsing_dict, i, j, end_parsing_index, r):
#         for c in range(0, sheet.ncols):
#             try:
#                 if sheet.cell_value(i + 1, c) != '':
#                     value = sheet.cell_value(r, c)
#                     if isinstance(sheet.cell_value(r, c), unicode) or isinstance(sheet.cell_value(r, c), str):
#                         value = value.strip()
#                     parsing_dict[str(sheet.cell_value(i, j)).lower().replace(" ", "")][
#                         str(sheet.cell_value(i + 1, c)).lower().replace(" ", "")].append(value)
#             except IndexError:
#                 break
#
#     @staticmethod
#     def set_end_parsing_index(sheet, top_level, i):
#         for n in range(i + 2, sheet.nrows):
#             if sheet.cell_value(n, 0) in top_level:
#                 return n
#             if n == sheet.nrows - 1:
#                 return sheet.nrows
#
#     @staticmethod
#     def set_parsing_dict(parsing_dict, inner_dict, parsing_dict_key, inner_key, k):
#         if inner_key != '':
#             inner_dict[inner_key] = []
#             if k == 0:
#                 parsing_dict[parsing_dict_key] = inner_dict
#             else:
#                 parsing_dict[parsing_dict_key].update(inner_dict)
#
#     def set_parsing_dict_key_and_inner_key(self, sheet, parsing_dict, i, j):
#         for k in range(0, sheet.ncols):
#             inner_dict = {}
#             parsing_dict_key = str(sheet.cell_value(i, j)).lower().replace(" ", "")
#             try:
#                 inner_key = str(sheet.cell_value(i + 1, k)).lower().replace(" ", "")
#             except IndexError:
#                 break
#
#             self.set_parsing_dict(parsing_dict, inner_dict, parsing_dict_key, inner_key, k)
#
#     @staticmethod
#     def set_parsing_arm_data_dict(sheet, parsing_arm_data_dict, i, j):
#         if isinstance(sheet.cell_value(i, j), unicode):
#             cell_value = sheet.cell_value(i, j).encode('utf-8')
#         else:
#             cell_value = sheet.cell_value(i, j)
#         if str(cell_value).lower().replace(" ", "") in parsing_arm_data_dict.keys():
#             if isinstance(sheet.cell_value(i, j+1), unicode) or isinstance(sheet.cell_value(i, j+1), str):
#                 parsing_arm_data_dict[str(sheet.cell_value(i, j)).lower().replace(" ", "")] = sheet.cell_value(i, (j + 1)).strip()
#             else:
#                 parsing_arm_data_dict[str(sheet.cell_value(i, j)).lower().replace(" ", "")] = sheet.cell_value(i, (j + 1))
#
#     def loop_through_columns(self, sheet, parsing_arm_data_dict, parsing_dict, top_level, i):
#         for j in range(0, sheet.ncols):
#             try:
#                 self.set_parsing_arm_data_dict(sheet, parsing_arm_data_dict, i, j)
#             except IndexError:
#                 break
#
#             if sheet.cell_value(i, j) in top_level:
#                 self.set_parsing_dict_key_and_inner_key(sheet, parsing_dict, i, j)
#                 end_parsing_index = self.set_end_parsing_index(sheet, top_level, i)
#                 self.loop_through_end_parsing_index(sheet, parsing_dict, i, j, end_parsing_index)
#                 break
#
#     def parse_workbook(self, local_excel_file_path, excel_sheet_name):
#         excel_book = open_workbook(local_excel_file_path)
#         parsing_dict = dict()
#         sheet = excel_book.sheet_by_name(excel_sheet_name)
#         parsing_arm_data_dict = {}
#         end_parsing_index = 0
#         top_level = ["Histologic Disease Exclusion Codes", "Prior Therapy (Drug Exclusion)",
#                      "IHC Results", "Non-Hotspot Rules", "Exclusion Variants", "Inclusion Variants"]
#
#         for item in ["ARM Official Name", "Arm Pathway Id", "ARM Gene", "ARM Id", "Arm Pathway Name", "Arm Description",
#                      "ARM Drug", "ARM Drug ID", "ARM Parser", "version", "stratum_id"]:
#             parsing_arm_data_dict[item.lower().replace(" ", "")] = ""
#         for i in range(0, sheet.nrows):
#             self.loop_through_columns(sheet, parsing_arm_data_dict, parsing_dict, top_level, i)
#
#
#         parsing_dict["armdata"] = parsing_arm_data_dict
#         # error_messages = []
#         # print json.dumps(TreatmentArmJSONValidator().validate_treatment_arm(DynamoDBFormatter().format_excel_results(parsing_dict), error_messages))
#         return parsing_dict
#
#     def adult_parse_workbook(self, local_excel_file_path, excel_sheet_name):
#         excel_book = open_workbook(local_excel_file_path)
#         parsing_dict = dict()
#         sheet = excel_book.sheet_by_name(excel_sheet_name)
#         parsing_arm_data_dict = {}
#         end_parsing_index = 0
#         top_level = ["Histologic Disease Exclusion Codes", "Prior Therapy (Drug Exclusion)", "Exclusion Criteria",
#                      "PTEN Results", "IHC Results", "Non-Hotspot Rules", "Exclusion Variants", "Inclusion Variants"]
#
#         for item in ["ARM Offical Name", "Arm Pathway Id", "ARM Gene", "ARM Id", "Arm Pathway Name", "Arm Description",
#                      "ARM Drug", "ARM Drug ID", "version"]:
#             parsing_arm_data_dict[item.lower().replace(" ", "")] = ""
#         for i in range(0, sheet.nrows):
#             self.loop_through_columns(sheet, parsing_arm_data_dict, parsing_dict, top_level, i)
#
#         parsing_dict["armdata"] = parsing_arm_data_dict
#         # error_messages = []
#         # print json.dumps(TreatmentArmJSONValidator().validate_treatment_arm(AdultDynamoDBFormatter().format_excel_results(parsing_dict), error_messages))
#         # print parsing_dict
#         return parsing_dict
#
#     @staticmethod
#     def check_for_empty_row(excel_sheet, index):
#         count = 0
#         returnValue = False
#         for i in range(0, excel_sheet.ncols):
#             try:
#                 if excel_sheet.cell_value(index, i) == '':
#                     count += 1
#             except IndexError:
#                 count += 1
#                 break
#         if int(count) == int(excel_sheet.ncols):
#             returnValue = True
#         return returnValue
#
#     @staticmethod
#     def include_all_sheet_names(workbook_path):
#         try:
#             book = open_workbook(workbook_path)
#         except IOError:
#             raise AttributeError("An invalid path was received for opening the workbook.")
#
#         sheets = book.sheet_names()
#         return sheets
#
# if __name__ == '__main__':
#     ParseExcelFile().adult_parse_workbook("adult_sample.xlsx", "EAY131_H")
#     # ParseExcelFile().parse_workbook("sample.xlsx", "APEC1621B")

import xlwings as xw
from functions import *


class BaseExcel:
    def __init__(self, excel_name):

        self.wb = xw.books[excel_name]

    def get_sheet_head(self, sheet_name):
        sheet = self.wb.sheets[sheet_name]
        cols = sheet.range('1:1')
        cols_list = {}
        i = 1
        for col in cols:
            if not col.value:
                break
            cols_list[column_to_name(i)] = col.value
            i += 1
        return cols_list

    """


        
    """

    def write_result_sheet(self, compare_sheet_name, write_sheet_name, book_name_col, result_col_name, mangodb_result,
                           start_row=2):

        ws = self.wb.sheets[write_sheet_name]
        ws.range(book_name_col + str(start_row)).expand("down").value = None
        ws.range(result_col_name + str(start_row)).expand("down").value = None

        current_row = start_row
        base_cols = self.get_sheet_head(compare_sheet_name)
        name_list = []

        check_list = []

        for r in mangodb_result:
            name_list.append([r['name']])
            check_list.append([check_dict_same(base_cols, r['value'])])
            current_row += 1
        ws.range(book_name_col + str(start_row)).expand("down").value = name_list
        ws.range(result_col_name + str(start_row)).expand("down").value = check_list


if __name__ == "__main__":
    a = BaseExcel("A表1删除列1.6v.xlsm")
    a.write_result_sheet("格式", "测试", "A", "F", mango_col_value(), )

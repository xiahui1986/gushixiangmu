import xlwings as xw
from functions import *
from datetime import datetime as dt


class BaseExcel:
    def __init__(self, excel_name):

        self.wb = xw.books[excel_name]
        self.op_sheet=self.wb.sheets["录入模块"]

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

    def get_book_path(self):
        return self.op_sheet.range("C1").value

    def get_process(self):
        return self.op_sheet.range("I1").value

    def get_files_name(self):
        re=[]
        p=self.get_book_path()+"/"
        for f in self.op_sheet.range("N2").expand("down").value:
            if f:
                re.append(p+f+".XLSX")
        return re

    def get_sheet_name(self, write_sheet_name):
        try:
            ws = self.wb.sheets[write_sheet_name]
        except:
            print("数据表未找到")
            return
        return (ws.range("F3").value).split("(")[0]



    """


        
    """

    def get_file_col(self, write_sheet_name, file_col_name, start_row, ):
        try:
            ws = self.wb.sheets[write_sheet_name]
        except:
            print("数据表未找到")
            return
        files = ws.range(file_col_name + str(start_row)).expand("down").value
        return files

    def write_result_sheet(self, compare_sheet_name, write_sheet_name, book_col_name, check_status_col, check_value_col,
                           mangodb_result,
                           sheet_name="", start_row=2):
        if sheet_name == "":
            if self.get_sheet_name(write_sheet_name):
                sheet_name = self.get_sheet_name(write_sheet_name)
            else:
                sheet_name = "Sheet1"

        try:
            ws = self.wb.sheets[write_sheet_name]
        except:
            print("数据表未找到")
            return
        # ws.range(book_col_name + str(start_row)).expand("down").value = None
        ws.range(check_status_col + str(start_row)).expand("down").value = None
        ws.range(check_value_col + str(start_row)).expand("down").value = None
        print("删除数据完成")
        current_row = start_row
        base_cols = self.get_sheet_head(compare_sheet_name)
        # name_list = []

        # 采用原宏文件的表名顺序
        mangodb_result = list(mangodb_result)
        file_name_list = self.get_file_col(write_sheet_name, book_col_name, start_row)

        # 使用numba加速后从速度25s提升到5s
        print(dt.now())


        file_sort_result = sort_result_by_filelist(base_cols, file_name_list, mangodb_result, sheet_name)
        check_status_list = file_sort_result['check_status_list']
        check_value_list = file_sort_result['check_value_list']
        print(dt.now())

        """
        for r in mangodb_result:
            #name_list.append([r['name']])
            re=check_dict_same(base_cols, r['value'])
            check_status_list.append([re['status']])
            check_value_list.append([re['result']])
            current_row += 1
        """
        # ws.range(book_col_name + str(start_row)).expand("down").value = name_list
        ws.range(check_status_col + str(start_row)).expand("down").value = check_status_list
        ws.range(check_value_col + str(start_row)).expand("down").value = check_value_list


if __name__ == "__main__":
    a = BaseExcel("提取录入增100字段完整宏.xlsm")
    print(a.get_files_name())
    #print(type(a.get_file_col("录入模块", "N", 2)))
    pass
    # a.write_result_sheet("格式", "录入模块", "C",
    #                                      "D", mango_col_value(), )

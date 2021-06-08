import xlwings as xw
from functions import *
from datetime import datetime as dt
from multiprocessing import Process
from mongodbop import MongoDBOP
import math
import random as rd

class BaseExcel:
    def __init__(self, excel_name):

        self.wb = xw.books[excel_name]
        self.op_sheet = self.wb.sheets["录入模块"]

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
        return int(self.op_sheet.range("I1").value)

    def get_books_name(self):
        res=[]
        if isinstance(self.op_sheet.range("N2").expand("down").value,str):
            res.append(self.op_sheet.range("N2").expand("down").value)
        else:
            res=list(self.op_sheet.range("N2").expand("down").value)
        return  res

    def del_fields_data(self):
        max_row=self.op_sheet.range("Q65535").end(-4162).row
        max_column1=self.op_sheet.range("AZZ2").end(-4159).column
        max_column2 = self.op_sheet.range("AZZ1").end(-4159).column
        max_column=max(max_column1,max_column2)
        max_column_name=column_to_name(max_column)
        self.op_sheet.range(f"P2:{max_column_name}{max_row}").value=None

    def get_files_name(self):
        re = []
        p = self.get_book_path() + "/"
        if isinstance(self.op_sheet.range("N2").expand("down").value,str):
            re.append(p + self.op_sheet.range("N2").expand("down").value, + ".XLSX")
        else:
            for f in self.op_sheet.range("N2").expand("down").value:
                if f:
                    re.append(p + f + ".XLSX")
        return re


    def write_cell_value(self, cell_name, val):
        self.wb.sheets["Sheet1"].range(cell_name).value = val

    def get_sheet_name(self, write_sheet_name):
        try:
            ws = self.wb.sheets[write_sheet_name]
        except:
            print("数据表未找到")
            return
        return (ws.range("F3").value).split("(")[0]

    def get_field_value_list(self, sheet_name, field_col_name, start_row, ):
        try:
            ws = self.wb.sheets[sheet_name]
        except:
            print("数据表未找到")
            return []
        field_values = ws.range(field_col_name + str(start_row)).expand("down").value
        return field_values

    """


        
    """

    def write_unit_check(self, write_sheet_name, result_list, start_cell):
        self.wb.sheets[write_sheet_name].range(start_cell).expand("down").value = result_list

    def get_file_col(self, write_sheet_name, file_col_name, start_row, ):
        try:
            ws = self.wb.sheets[write_sheet_name]
        except:
            print("数据表未找到")
            return
        files = ws.range(file_col_name + str(start_row)).expand("down").value
        return files

    # 写入表检查结果到表中
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

    def get_write_processes(self):
        return int(self.op_sheet.range("K1").value)


    def get_data_array(self):
        max_row=self.op_sheet.range("Q65535").end(-4162).row
        target_cells=self.get_field_value_list("录入模块","F",3)
        #max_column1=self.op_sheet.range("AZZ2").end(-4159).column
        #max_column2 = self.op_sheet.range("AZZ1").end(-4159).column
        #max_column=max(max_column1,max_column2)
        max_column_name=column_to_name(len(target_cells)*7+15)
        check_list=self.op_sheet.range(f"P2:{max_column_name}{max_row}").value
        return check_list

    def check_col_same(self):
        check_list=self.get_data_array()
        check_result=[]
        for j in range(math.ceil(len(check_list[0])/7)):#向上取整
            temp_list=[]
            for i in range(len(check_list)):
                temp_list.append(check_list[i][j*7+1])
            if len(set(temp_list))==1:
                check_result.append(["字段一致完整",check_list[0][j*7+2],check_list[0][j*7+2]])
            else:
                check_result.append(["字段不一致完整", check_list[0][j * 7 + 2], check_list[0][j * 7 + 2]])
        return check_result

    def write_p_random(self):
        start_row=2
        start_col=16
        cycle_=7
        cells=self.op_sheet.range("F3").expand("down").value
        files=self.get_files_name()
        for j in range(len(cells)):
            for m in range(len(files)):
                col_name=column_to_name(start_col+cycle_*j)
                self.op_sheet.range(f"{col_name}{m+start_row}").value=rd.random()





def write_excel_table_check(excel_name,cell_name, sheet_name,write_start_cell):
    m=MongoDBOP()

    be = BaseExcel(excel_name)
    books=be.get_books_name()
    write_list= m.get_value_by_cell_sort_by_list(books,cell_name)
    be.write_unit_check(sheet_name,write_list, write_start_cell)

def write_field_data(excel_name):
    be = BaseExcel(excel_name)
    books = be.get_books_name()
    m = MongoDBOP()
    cell_names=list(be.op_sheet.range("F3").expand("down").value)
    max_process=be.get_write_processes()
    i=0
    current_porcess=1
    start_col_num=16
    col_cycle=7
    start_row=2
    result_list=[]
    books_count=len(books)
    for i in range(books_count):
        result_list.append([])
    cell_count=0
    for cell_name in cell_names:
        write_start_cell = column_to_name(i * col_cycle + start_col_num) + str(start_row)
        field_datas=m.get_value_by_cell_sort_by_list(books, cell_name)
        col_name=column_to_name(start_col_num+col_cycle*cell_count)
        excel_cost_data=list(be.op_sheet.range(f"{col_name}{start_row}:{col_name}{start_row+books_count-1}").value)
        for i in range(books_count):
            result_list[i].extend([excel_cost_data[i]])
            result_list[i].extend(field_datas[i][1:])
        cell_count +=1

    be.op_sheet.range("P2").expand("down").value = result_list


def process_write_excel_table_check(excel_name,sheet_name):
    be = BaseExcel(excel_name)
    books = be.get_books_name()
    m = MongoDBOP()
    cell_names=list(be.op_sheet.range("F3").expand("down").value)
    max_process=be.get_write_processes()
    i=0
    current_porcess=1
    start_col_num=16
    col_cycle=7
    start_row=2
    #result_list=[]
    """
    books_count=len(books)
    for i in range(books_count):
        result_list.append([])
    for cell_name in cell_names:
        write_start_cell = column_to_name(i * col_cycle + start_col_num) + str(start_row)
        field_datas=m.get_value_by_cell_sort_by_list(books, cell_name)
        for i in range(books_count):
            result_list[i].extend(field_datas[i])

    #return    result_list
    be.op_sheet.range("P2").expand("down").value = result_list
    """
    p = []
    for cell_name in cell_names:
        if current_porcess>max_process:
            for p_ in p:
                p_.start()
            for p_ in p:
                p_.join()
            current_porcess=1
            p = []

        if current_porcess<=max_process:
            write_start_cell=column_to_name(i*col_cycle+start_col_num)+str(start_row)
            p.append(Process(target=write_excel_table_check,args=(excel_name,cell_name,sheet_name,write_start_cell)))
            current_porcess=current_porcess+1
            i=i+1

    else:
        for p_ in p:
            p_.start()
        for p_ in p:
            p_.join()
        p=[]




if __name__ == "__main__":
    be=BaseExcel("提取录入增100字副本.xlsm")
    print(dt.now())
    be.check_col_same()
    #be.wb.api.screen_updating=False
    #be.write_p_random()
    #be.wb.api.screen_updating = True
    #write_field_data("提取录入增100字段完整宏(3).xlsm")
    print(dt.now())
    #be.del_fields_data()
    #process_write_excel_table_check("提取录入增100字段完整宏.xlsm","录入模块")
    #a = BaseExcel("提取录入增100字段完整宏.xlsm")
    #print(a.get_files_name())
    # print(type(a.get_file_col("录入模块", "N", 2)))
    pass
    # a.write_result_sheet("格式", "录入模块", "C",
    #                                      "D", mango_col_value(), )

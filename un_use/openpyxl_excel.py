import openpyxl as op
from functions import *
from multiprocessing import Process
class OpenExcel:
    def __init__(self,book_name,new_file_name):
        self.book_name=book_name
        self.wb=op.load_workbook(book_name,)
        self.new_file_name=new_file_name

    def write_cell(self,cell_addr,target_value):
        cell_t=get_target_cell_pos(cell_addr)
        sheet_name=cell_t['sheet_name']
        col_name=cell_t['col_name']
        row_name=cell_t['row_name']
        sheet=self.wb.get_sheet_by_name(sheet_name)
        sheet[f"{col_name}{row_name}"].value=target_value

    def save(self):
        self.wb.save(self.new_file_name)

    def __del__(self):
        del self.book_name
        del self.wb

def write_cell(full_name):
    oe=OpenExcel(full_name)
    oe.write_cell("Sheet1(C3)",200)
    oe.save()


def mypool(method,max_p,args_list):
    p=[]
    c_p=1
    for args in args_list:
        if c_p>max_p:
            for p_ in p:
                p_.start()
            for p_ in p:
                p_.join()
            c_p=1
            p=[]
        if c_p<=max_p:
            if type(args) is not list:
                p.append(Process(target=method,args=(args,)))
            else:
                p.append(Process(target=method, args=(*args,)))
            c_p+=1
    else:
        for p_ in p:
            p_.start()
        for p_ in p:
            p_.join()
        p = []

if __name__=="__main__":
    import multiprocessing as mul
    import datetime as dt
    print(dt.datetime.now())


    p=[]

    full_name_list=[]
    for i in range(1,100):
        book_name=f'{i}-aaaaa.XLSX'
        full_name=r"E:\股市数据处理需求编程\shyg\test3/"+book_name
        #pool.apply_async(write_cell,args=(full_name,))
        full_name_list.append(full_name)
    mypool(write_cell,12,full_name_list)
           #p.append(mul.Process(target=write_cell,args=(full_name,)))
    #pool.map(write_cell,full_name_list)


    print(dt.datetime.now())
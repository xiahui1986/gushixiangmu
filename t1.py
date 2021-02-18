import pyodbc
import time
import multiprocessing as mul
from multiprocessing import Manager
import pymongo
from configparser import ConfigParser as cp

import os

myclient = pymongo.MongoClient("mongodb://127.0.0.1:27017/")
mydb = myclient["db"]
mycol = mydb["runoob"]


def read_excel(excel_path, id):
    conn_str = (
            'DRIVER={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};' + excel_path
        # r'DBQ=C:\Users\Administrator\Desktop\辅助.xlsx;'
    )
    cnxn = pyodbc.connect(conn_str, autocommit=True)
    crsr = cnxn.cursor()
    row = crsr.execute("select * from [Sheet1$]").fetchall()
    this_dict = {}
    this_dict["name"] = excel_path
    this_dict["id"] = id

    v = {}
    i = 1
    for field in row[0].cursor_description:
        if "F" in field[0]:
            break
        v["\'" + str(i) + "\'"] = field[0]
        i = i + 1
        # print(field[0])
    # my_dict[id]=this_dict
    this_dict["value"] = v
    mycol.insert(this_dict)
    # my_dict[excel_path] = row[0].cursor_description[0]



def process(max_process,max_file):

    #s = r'DBQ=C:\Users\Administrator\Desktop\辅助.xlsx;'

    for i in range(max_file+1):
        if i==max_file:
            file_name = r"DBQ=E:/股市数据处理需求编程/shyg/shyg/" + str(i) + "-aaaaa.xlsx"
            p.append(mul.Process(target=read_excel, args=(file_name, i)))
            for p_ in p:
                p_.start()
            for p_ in p:
                p_.join()
            p = []
            break
        if i==0:
            j = 1
            p = []
            continue
        if j<=max_process and i<max_file-1:
            file_name = r"DBQ=E:/股市数据处理需求编程/shyg/shyg/" + str(i) + "-aaaaa.xlsx"
            p.append(mul.Process(target=read_excel, args=(file_name, i)))
            j=j+1

            continue
        else:
            for p_ in p:
                p_.start()
            for p_ in p:
                p_.join()
            p = []
            print(i,j)
            j = 1
            file_name = r"DBQ=E:/股市数据处理需求编程/shyg/shyg/" + str(i) + "-aaaaa.xlsx"
            p.append(mul.Process(target=read_excel, args=(file_name, i)))
            j=j+1
            continue


    #    for i in range(500):
    #        p.append(Process(target=read_excel,args=(s,)))



# sql="select name from tables where object_id('[Sheet1$]')"
# f=crsr.execute(sql).fetchall()

if __name__ == "__main__":

    manager = Manager()
    # my_list = manager.list()
    my_dict = manager.dict()
    mycol.delete_many({})
    t1 = time.time()
    process(2,5)
    t2 = time.time()
    print(my_dict)
    print(t1)
    print(t2)

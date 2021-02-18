"""
存储一些方法

"""
from configparser import ConfigParser as CP
import pymongo
import pyodbc
import os
import multiprocessing as mul


# excel 列名变数字
def colname_to_num(colname):
    if type(colname) is not str:
        return colname
    col = 0
    power = 1
    for i in range(len(colname) - 1, -1, -1):
        ch = colname[i]
        col += (ord(ch) - ord('A') + 1) * power
        power *= 26
    return col


# excel 数字变列名
def column_to_name(colnum):
    if type(colnum) is not int:
        return colnum
    str_ = ''
    while (not (colnum // 26 == 0 and colnum % 26 == 0)):
        temp = 25
        if (colnum % 26 == 0):
            str_ += chr(temp + 65)
            colnum = colnum - 1
        else:

            str_ += chr(colnum % 26 - 1 + 65)
        colnum //= 26
        # print(str)
    # 倒序输出拼写的字符串
    return str_[::-1]


# 获取路径下的文件
def get_path_files(_path=os.getcwd()):
    return list(map(lambda x: _path + "/" + x, os.listdir(_path)))


def check_is_excel(file_name):
    ext_name = os.path.splitext(file_name)[1]
    if ext_name in ['.xls', '.xlsx', '.xlsb', '.xlsm', '.XLSX', '.XLSB', '.XLSM', 'XLS']:
        return True
    else:
        return False
    pass

def get_cp():
    cp = CP()
    cp.read("configs.cfg",encoding="utf-8-sig")
    return cp

def write_cp(section,key,value):
    cp = CP()
    cp.read("configs.cfg", encoding="utf-8-sig")
    cp.set(section,key,value)

    cp.write(open("configs.cfg",'w',encoding="utf-8-sig"))
    return "写入成功"

def get_mango_client_name():
    return get_cp().get("mongo", "client")


def get_mango_db_name():
    return get_cp().get("mongo", "db")


def get_mango_col_name():
    return get_cp().get("mongo", "col")


def get_excel_driver():
    return get_cp().get("excel", "driver")


def get_mango_col():
    myclient = pymongo.MongoClient(get_mango_client_name())
    mydb = myclient[get_mango_db_name()]
    mycol = mydb[get_mango_col_name()]
    return mycol


def insert_mango_col(this_dict):
    get_mango_col().insert(this_dict)


def del_mango_col():
    get_mango_col().delete_many({})


def mango_col_value():
    return get_mango_col().find()


def read_excel_write_db(excel_path, sheet_name="Sheet1", id=0):
    conn_str = (
            get_excel_driver() + "DBQ=" + excel_path + ";"
        # r'DBQ=C:\Users\Administrator\Desktop\辅助.xlsx;'
    )
    cnxn = pyodbc.connect(conn_str, autocommit=True)
    crsr = cnxn.cursor()
    row = crsr.execute("select * from [" + sheet_name + "$]").fetchall()
    this_dict = {}
    this_dict["name"] = excel_path
    this_dict["id"] = id
    v = {}
    i = 1
    for field in row[0].cursor_description:
        if "F" in field[0]:
            break
        v[column_to_name(i)] = field[0]
        # v["\'" + str(i) + "\'"] = field[0]
        i = i + 1

    this_dict["value"] = v
    # return this_dict
    insert_mango_col(this_dict)


def process(file_path, max_process, max_file=-1):
    j = 1
    p = []
    current_files = 1
    files = get_path_files(file_path)
    for excel_file in files:
        if max_file != -1:
            if current_files > max_file:
                continue
        if not check_is_excel(excel_file):
            continue
        if j <= max_process:
            # file_name = r"DBQ=E:/股市数据处理需求编程/shyg/shyg/" + str(i) + "-aaaaa.xlsx"
            p.append(mul.Process(target=read_excel_write_db, args=(excel_file,)))
            j = j + 1
            current_files = current_files + 1
            continue
        else:
            for p_ in p:
                p_.start()
            for p_ in p:
                p_.join()
            p = []
            print(current_files, j)
            j = 1
            # file_name = r"DBQ=E:/股市数据处理需求编程/shyg/shyg/" + str(i) + "-aaaaa.xlsx"
            p.append(mul.Process(target=read_excel_write_db, args=(excel_file,)))
            current_files = current_files + 1
            j = j + 1
            continue
    else:
        # file_name = r"DBQ=E:/股市数据处理需求编程/shyg/shyg/" + str(i) + "-aaaaa.xlsx"
        # p.append(mul.Process(target=read_excel_write_db, args=(excel_file, 0)))
        for p_ in p:
            p_.start()
        for p_ in p:
            p_.join()
        p = []


def check_dict_same(dict1,dict2):
    if dict1==dict2:
        return "格式相同"
    s=""
    for key in dict1:
        if dict1[key]!=dict2[key]:
            s+=f"列：{key} 的值不相同，分别为{dict1[key]},{dict2[key]};"
    return s


if __name__ == "__main__":
    del_mango_col()
    process(r"E:/股市数据处理需求编程/shyg/shyg/", 3, 20)

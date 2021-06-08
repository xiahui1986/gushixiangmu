"""
存储一些方法

"""
from configparser import ConfigParser as CP
import pymongo
import pyodbc
import os
import multiprocessing as mul
from multiprocessing import Pool,Manager
#from numba import jit
#from numba import cuda
import warnings
from datetime import datetime as dt
import re
import time

warnings.filterwarnings('ignore')

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

#分解如sheet2(j2),分解为Sheet2, j ,2
def get_target_cell_pos(cell_value):
    sheet_name=cell_value.split("(")[0]
    temp=cell_value.split('(')[1].split(")")[0]
    cell_name=re.split(r'(\D+)',temp)
    col_name=cell_name[1].upper()
    row_name=cell_name[2]
    return {
        "sheet_name":sheet_name,
        "col_name":col_name,
        "row_name":row_name,
    }



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


def get_mango_client_name():
    return "mongodb://127.0.0.1:27017/"


def get_mango_db_name():
    return "db"


def get_mango_col_name():
    return "runoob"


def get_excel_driver():
    return "DRIVER={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};"



def get_mango_col():
    myclient = pymongo.MongoClient(get_mango_client_name())
    mydb = myclient[get_mango_db_name()]
    mycol = mydb[get_mango_col_name()]
    return mycol


def insert_mango_col(this_dict):
    get_mango_col().insert(this_dict)

def insert_mango_many(data_list):
    get_mango_col().insert_many(data_list)

def del_mango_col():
    get_mango_col().delete_many({})


def mango_col_value():
    return get_mango_col().find()

def mango_col_find_fields(sheet_name):
    return  get_mango_col().find({}, {"name": 1, sheet_name: {"value": 1,"error":1}})




"""
def get_file_name(file_path):
    (filepath, tempfilename) = os.path.split(file_path)
    (filename, extension) = os.path.splitext(tempfilename)
    return filename
"""

#@jit
def sort_result_by_filelist(base_cols,file_list,result_list,sheet_name):

    check_status_list = []
    check_value_list = []
    # 采用原宏文件的表名顺序
    result={"check_status_list":"","check_value_list":""}
    for i in file_list:
        flag = 0
        for r in result_list:

            if i == r['name'].split("/")[-1].split(".")[0]:
                flag = 1
                try:
                    v=r[sheet_name]['value']
                except:
                    try:
                        error=r[sheet_name]['error']
                        check_status_list.append(["检查不成功"])
                        check_value_list.append([error])
                        continue
                    except:
                        check_status_list.append(["检查不成功"])
                        check_value_list.append(["检查不成功且无错误信息"])
                        continue

                re = check_dict_same(base_cols, r[sheet_name]['value'])

                check_status_list.append([re['status']])
                check_value_list.append([re['result']])
                break
        if flag == 0:
            check_status_list.append([""])
            check_value_list.append([""])
    result["check_status_list"]=check_status_list
    result["check_value_list"] = check_value_list
    return result


def check_dict_same(dict1,dict2):
    result = {}

    if dict1==dict2:
        result["status"] = f"格式相同"
        result["result"] = f"检查成功"
        return result
    s="格式不同："

    for key in dict2:
        try:
            if dict1[key]!=dict2[key]:
                s+=f"第{key} 列[{dict2[key]}]不同;"
        except KeyError:

            s += f"源表中第{key}在格式表中不存在;"

    result["status"]=f"格式不同"
    result["result"] = s
    return result

#sheet_name="Sheet1",
def read_excel_write_db(process_op_dict,excel_path, sheet_name, id=0):
    conn_str = (
            get_excel_driver() + "DBQ=" + excel_path + ";"
        # r'DBQ=C:\Users\Administrator\Desktop\辅助.xlsx;'
    )
    #print(conn_str)
    try:
        cnxn = pyodbc.connect(conn_str, autocommit=True)
        crsr = cnxn.cursor()
    except:
        print(f"{excel_path}连接失败 ")
        process_op_dict["error"] = f"{excel_path}连接失败 "
        return
    this_dict = {}
    this_dict["name"] = excel_path

    this_dict[sheet_name]=sheet_to_dict(process_op_dict,crsr,sheet_name,excel_path)
    #this_dict["Sheet2"] = sheet_to_dict(crsr, "Sheet2")
    #this_dict["Sheet3"] = sheet_to_dict(crsr, "Sheet3")

    insert_mango_col(this_dict)



def sheet_to_dict(process_op_dict,crsr,sheet_name,excel_path):
    this_dict = {}
    try:
        row = crsr.execute("select * from [" + sheet_name + "$]").fetchall()
    except:
        print(f"{excel_path}中的表{sheet_name}不存在")
        #process_op_dict["error"]=f"{excel_path}中的表{sheet_name}不存在,或excel连接错误 "
        this_dict["error"]=f"{sheet_name}不存在"
        return this_dict

    this_dict["id"] = ""
    v = {}
    i = 1

    for field in row[0].cursor_description:
        if "F" in field[0]:
            break
        v[column_to_name(i)] = field[0]
        # v["\'" + str(i) + "\'"] = field[0]
        i = i + 1
    this_dict["value"] = v
    return this_dict





def process(process_op_dict,file_list,sheet_name, max_process,max_file=-1):
    j = 1
    p = []
    current_files = 1
    #files = get_path_files(file_path)

    ts=dt.now()
    print("开始时间：",ts)


    for excel_file in file_list:

        if process_op_dict["stop_process"]==1:
            print("数据写入操作手工中止", )
            break
        if max_file != -1:
            if current_files > max_file:
                continue

        if j <= max_process:
            # file_name = r"DBQ=E:/股市数据处理需求编程/shyg/shyg/" + str(i) + "-aaaaa.xlsx"
            p.append(mul.Process(target=read_excel_write_db, args=(process_op_dict,excel_file,sheet_name)))
            j = j + 1
            current_files = current_files + 1
            continue
        else:
            for p_ in p:
                p_.start()
            for p_ in p:
                p_.join()
            p = []
            process_op_dict["current_file"]=current_files-1
            print(current_files-1, j-1)
            j = 1
            # file_name = r"DBQ=E:/股市数据处理需求编程/shyg/shyg/" + str(i) + "-aaaaa.xlsx"
            p.append(mul.Process(target=read_excel_write_db, args=(process_op_dict,excel_file,sheet_name)))
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
    te=dt.now()
    process_op_dict["file_count"] = current_files-1

    print(f"检查文件数量{j}")
    print(process_op_dict["file_count"] )
    print("结束时间：",te)
    process_op_dict["end_time"]=dt.now()
    print("耗时",te-ts)

def read_excel_data(process_op_dict,excel_path,book_name, sheet_name):
    #print("start:",dt.now())
    conn_str = (
            get_excel_driver() + "DBQ=" + excel_path +"/"+ book_name+".XLSX;"
    )
    try:
        cnxn = pyodbc.connect(conn_str, autocommit=True)
        crsr = cnxn.cursor()
    except:
        print(f"{excel_path}连接失败 ")
        process_op_dict["error"] = f"{excel_path}连接失败 "
        return
    #print("excel connect:", dt.now())
    this_dict = {}
    try:
        row = crsr.execute("select * from [" + sheet_name + "$]").fetchall()
    except:
        print(f"{excel_path}中的表{sheet_name}不存在")
        process_op_dict["error"]=f"{excel_path}中的表{sheet_name}不存在,或excel连接错误 "
        this_dict["error"]=f"{sheet_name}不存在"
        return
    #print("check sheet:", dt.now())
    write_mongo_data=[]
    row_num=1
    fields=row[0].cursor_description
    #print("get_fields:", dt.now())
    for r in row:
        data={}
        j=0
        data["excel_path"] = excel_path
        data["book_name"] = book_name
        data["sheet_name"] = sheet_name
        data["uid"] = row_num
        for d in  fields:

            data[column_to_name(j+1)]={"name":d[0],"value":r[j]}
            j+=1
        write_mongo_data.append(data)
        row_num+=1
    #print("for write list:", dt.now())
    insert_mango_many(write_mongo_data)
    #print("write db:", dt.now())


def process_write_data_db(process_op_dict,excel_path,books, sheet_name,max_processes,max_file=-1):
    j = 1
    p = []
    current_files = 1
    #files = get_path_files(file_path)

    ts=dt.now()
    print("开始时间：",ts)
    for excel_file in books:
        if process_op_dict["stop_process"]==1:
            print("数据写入操作手工中止", )
            break
        if max_file != -1:
            if current_files > max_file:
                continue
        if j <= max_processes:
            p.append(mul.Process(target=read_excel_data, args=(process_op_dict,excel_path,excel_file, sheet_name)))
            j = j + 1
            current_files = current_files + 1
            continue
        else:
            for p_ in p:
                p_.start()
            for p_ in p:
                p_.join()
            p = []
            process_op_dict["current_file"]=current_files-1
            print(current_files-1, j-1)
            j = 1
            p.append(mul.Process(target=read_excel_data, args=(process_op_dict,excel_path,excel_file, sheet_name)))
            current_files = current_files + 1
            j = j + 1
            continue
    else:
        for p_ in p:
            p_.start()
        for p_ in p:
            p_.join()
        p = []
    te=dt.now()
    process_op_dict["file_count"] = current_files-1

    print(f"处理文件数量{current_files-1}")
    print(process_op_dict["file_count"] )
    print("结束时间：",te)
    process_op_dict["end_time"]=dt.now()
    print("耗时",te-ts)


def read_excel_data_by_field(process_op_dict,excel_path,book_name, sheet_name,cells):
    conn_str = (
            get_excel_driver() + "DBQ=" + excel_path +"/"+ book_name+".XLSX;"
    )
    try:
        cnxn = pyodbc.connect(conn_str, autocommit=True)
        crsr = cnxn.cursor()
    except:
        print(f"{excel_path}/{book_name}连接失败 ")
        process_op_dict["error"] = f"{excel_path}/{book_name}连接失败 "
        return
    #print("excel connect:", dt.now())
    this_dict = {}
    try:
        row = crsr.execute("select * from [" + sheet_name + "$]").fetchall()
    except:
        print(f"{excel_path}中的表{sheet_name}不存在")
        process_op_dict["error"]=f"{excel_path}中的表{sheet_name}不存在,或excel连接错误 "
        this_dict["error"]=f"{sheet_name}不存在"
        return
    fields=row[0].cursor_description
    data = {}
    data["excel_path"] = excel_path
    data["book_name"] = book_name
    data["sheet_name"] = sheet_name
    for cell_name in cells:
        split_cell_name=get_target_cell_pos(cell_name)
        sheet_name=split_cell_name["sheet_name"]
        col_name=split_cell_name["col_name"]
        row_name=int(split_cell_name["row_name"])
        col_num=colname_to_num(col_name)
        try:
            data[col_name +str(row_name)]={
                "head":row[row_name-2][0]+row[row_name-2][1],
                "name":fields[col_num-1][0],
                "value":row[row_name-2][col_num-1]
            }
        except:
            data[col_name +str(row_name)]={
                "head":"wrong_date",
                "name":fields[col_num-1][0],
                "value":"data get failed",
            }
            process_op_dict["error"]=book_name + "的"+cell_name +"获取失败"
    insert_mango_col(data)

def process_read_excel_data_by_field(process_op_dict,excel_path,books_name, sheet_name,cells,max_process):
    i=0
    current_porcess=1
    p = []
    for book_name in books_name:
        if process_op_dict["stop_process"]==1:
            print("数据写入操作手工中止", )
            break
        if current_porcess>max_process:
            for p_ in p:
                p_.start()
            for p_ in p:
                p_.join()
            process_op_dict["current_file"] = i
            current_porcess=1
            p = []

        if current_porcess<=max_process:
            p.append(mul.Process(target=read_excel_data_by_field,args=(process_op_dict, excel_path, book_name, sheet_name, cells)))
            current_porcess=current_porcess+1
            i=i+1
    else:
        for p_ in p:
            p_.start()
        for p_ in p:
            p_.join()

        p=[]

    process_op_dict["file_count"] = i
    process_op_dict["end_time"]=dt.now()


def mypool(method,max_p,args_list,process_op_dict):
    p=[]
    c_p=1
    file_count=0
    change_point=1000
    for args in args_list:
        try:
            if process_op_dict["stop_process"]==1:
                process_op_dict["file_count"] =file_count
                #print("数据写入操作手工中止,处理文件数：", file_count)
                break
        except:
            pass
        if c_p>max_p:
            for p_ in p:
                p_.start()
            for p_ in p:
                p_.join()
            file_count=file_count+max_p
            process_op_dict["current_file"] = file_count
            c_p=1
            p=[]
        if c_p<=max_p:
            if type(args) is not list:
                p.append(mul.Process(target=method,args=(args,process_op_dict)))
            else:
                p.append(mul.Process(target=method, args=(*args,process_op_dict)))
            c_p+=1
    else:
        for p_ in p:
            p_.start()
        for p_ in p:
            p_.join()
        process_op_dict["file_count"]=len(args_list)
        p = []



if __name__ == "__main__":
    del_mango_col()
    from BaseExcelOP import BaseExcel


    be=BaseExcel("提取录入增100字段完整宏.xlsm")
    max_process = be.get_process()
    #read_excel_data_by_field({}, 'E:/股市数据处理需求编程/shyg/shyg',"1-aaaaa", "Sheet1",list(be.op_sheet.range("F3").expand("down").value))
    cells = list(be.op_sheet.range("F3").expand("down").value)
    books_name=be.get_books_name()

    process_op_dict = Manager().dict()
    process_op_dict["stop_process"] = 0
    process_op_dict["current_file"] = 0
    process_op_dict["error"] = ""
    process_op_dict["file_count"] = 0
    process_read_excel_data_by_field(process_op_dict, 'E:/股市数据处理需求编程/shyg/shyg', books_name, "Sheet1", cells,max_process)
    pass

    #c=get_mango_col()
    #x=list(c.find({},{"name":1,"Sheet1":{"data":1}}))
    #for i in c.find({},{"name":1,"Sheet1":{"data":1}}):
    #    print(i)
    #    break
    #print(list(c.find({},{"name":1,"Sheet1":{"data":1}})))
    """
    del_mango_col()
    from BaseExcelOP import BaseExcel
    be=BaseExcel("提取录入增100字段完整宏.xlsm")
    books=be.get_books_name()
    sheet_name=be.get_sheet_name("录入模块")
    max_process=be.get_process()

    execl_path=be.get_book_path()
    #read_excel_data({},execl_path,books[0], sheet_name)
    process_write_data_db(process_op_dict, execl_path,books,sheet_name,max_process,-1)


    """
    """
    t1=dt.now()
    p=[]
    for i in range(1,201):
        p.append(mul.Process(target=read_excel_data, args=({}, f'E:/股市数据处理需求编程/shyg/shyg/{i}-aaaaa.XLSX',f"{i}-aaaaa", "Sheet1")))
        #read_excel_data({}, f'E:/股市数据处理需求编程/shyg/shyg/{i}-aaaaa.XLSX', "Sheet1")
    for j in p:
        j.start()
    for j in p:
        j.join()
    t2 = dt.now()
    print(t1,t2)
    pass
    #read_excel_write_db('E:/股市数据处理需求编程/shyg/shyg/1-aaaaa.XLSX',"Sheet4")
    #process(r"E:/股市数据处理需求编程/shyg/shyg/", 3, 20)
    """
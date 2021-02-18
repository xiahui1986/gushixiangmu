import pyodbc
import time
#from multiprocessing import Process
import multiprocessing as mul
from multiprocessing import Manager
import pymongo

myclient = pymongo.MongoClient("mongodb://127.0.0.1:27017/")
mydb = myclient["db"]
mycol = mydb["runoob"]
def read_excel(excel_path,id):

    conn_str = (
        r'DRIVER={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};'+excel_path
        #r'DBQ=C:\Users\Administrator\Desktop\辅助.xlsx;'
        )
    cnxn = pyodbc.connect(conn_str, autocommit=True)

    crsr = cnxn .cursor()
    #for worksheet in crsr.tables():
    #    print(worksheet)
    #print(excel_path)


    row=crsr.execute("select * from [Sheet1$]").fetchall()
    """
    this_dict={}
    this_dict["name"]=excel_path
    this_dict["id"] = id

    v={}
    i=1
    for field in row[0].cursor_description:
        v["\'"+str(i)+"\'"] = field[0]
        i=i+1
        #print(field[0])
    #my_dict[id]=this_dict
    this_dict["value"] = v
    mycol.insert_one( this_dict)
    pass
    #my_dict[excel_path] = row[0].cursor_description[0]
    """
def process():
    p=[]
    s=r'DBQ=C:\Users\Administrator\Desktop\辅助.xlsx;'
    for i in range(529):
        if i > 0:
            file_name = r"DBQ=E:/股市数据处理需求编程/shyg/shyg/" + str(i) + "-aaaaa.xlsx"
            read_excel(file_name, i)
    #        p.append(mul.Process(target=read_excel, args=(file_name,i)))

#    for i in range(500):
#        p.append(Process(target=read_excel,args=(s,)))
    #for i in p:
    #    i.start()
    #for i in p:
    #    i.join()


#sql="select name from tables where object_id('[Sheet1$]')"
#f=crsr.execute(sql).fetchall()

if __name__=="__main__":
    manager = Manager()
    #my_list = manager.list()
    my_dict = manager.dict()
    mycol.delete_many({})
    t1=time.time()
    process()
    t2 = time.time()
    print(my_dict)
    print(t1)
    print(t2)

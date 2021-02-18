import openpyxl as op
import multiprocessing as mul
import time
import sqlite3
import pymongo
myclient=pymongo.MongoClient("mongodb://192.168.18.135:27017/")
mydb=myclient["db"]
mycol=mydb["runoob"]
conn = sqlite3.connect('test.db')
c=conn.cursor()
sqls=[]

def read_execl(execl_name):
    ex=op.open(execl_name,read_only=True)
    sh=ex.get_sheet_by_name("Sheet1")
    value_=str(sh["A1"].value)


    sqls.append("insert into sheet1 (value__) values (\'"+ value_ +"\') ")
    #c.execute("insert into sheet1 (value__) values (\'"+ value_ +"\') ")
    mydict={"A1":value_}
    mycol.insert(mydict)
    ex.close()
    print(execl_name)

def processes():
    p=[]
    for j in range(3):
        for i in range(2):
            if i>0:
                file_name="E:/股市数据处理需求编程/shyg/shyg/"+str(i)+"-aaaaa.xlsx"
                p.append(mul.Process(target=read_execl,args=(file_name,)))
    for i in p:
        i.start()
    for i in p:
        i.join()
    #conn.commit()
    conn.close()
    print(12)


if __name__=="__main__":
    t1=time.time()
    processes()
    t2 = time.time()
    print(t1)
    print(t2)



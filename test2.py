import openpyxl as op
import multiprocessing as mul
import time



def read_execl(execl_name):
    ex = op.open(execl_name, read_only=True)



def processes():
    p = []
    for i in range(200):
        if i > 0:
            file_name = "E:/股市数据处理需求编程/shyg/shyg/" + str(i) + "-aaaaa.xlsx"
            p.append(mul.Process(target=read_execl, args=(file_name,)))
    for i in range(199):
        p[i].start()


if __name__ == "__main__":
    t1 = time.time()
    processes()
    t2 = time.time()
    print(t1)
    print(t2)


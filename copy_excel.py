import openpyxl as op
import time
from shutil import copyfile
def copy_excel(base_path,base_file,name_base,start_num,end_num,):
    #xl = op.load_workbook(base_path + base_file)
    for i in range(start_num,end_num):
        copyfile(base_path + base_file,base_path+str(i)+name_base)
        print(i)
        #xl.save(base_path+str(i)+name_base)
        #time.sleep(1)


if __name__=="__main__":
    base_path=r'E:/股市数据处理需求编程/shyg/shyg/'
    base_file=r'1-aaaaa.XLSX'
    name_base=r'-aaaaa.XLSX'

    copy_excel(base_path,base_file,name_base,1561,5000)
    print('finish')
from functions import *
from BaseExcelOP import BaseExcel

def test():
    be=BaseExcel("A表1删除列1.6v.xlsm")
    be_dict=be.get_sheet_head("格式")
    mango_values=mango_col_value()
    for mv in mango_values:
        differ = set(mv["value"])^ set(be_dict)
        print(check_dict_same(be_dict,mv["value"]))


if __name__=="__main__":
    test()
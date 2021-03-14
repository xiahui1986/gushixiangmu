import PySimpleGUI as sg
from functions import *
from datetime import datetime
import multiprocessing
from multiprocessing import Manager
import sys,os
from threading import Thread
import time
from mongodbop import MongoDBOP
import mongodbop
import win32_excel_no_use
#import sql_excel_no_use
import subprocess
import traceback

def gui_get_max_process(excel_name):
    from BaseExcelOP import BaseExcel
    be = BaseExcel(excel_name)
    return int(be.get_process())


def gui_get_path(excel_name):
    from BaseExcelOP import BaseExcel
    be = BaseExcel(excel_name)
    return be.get_book_path() + "/"


def write_sheet_data(excel_name):
    import BaseExcelOP as BEOP
    print("数据写入开始:" + str(dt.now()))
    BEOP.write_field_data(excel_name)
    print("数据写入结束:" + str(dt.now()))
    del BEOP

def sheet_write_result(excel_name):
    from BaseExcelOP import BaseExcel
    be = BaseExcel(excel_name)
    be.op_sheet.range("C3").expand("down").value = []
    be.op_sheet.range("D3").expand("down").value = []
    be.op_sheet.range("E3").expand("down").value = []
    print("数据清除完成")
    print("检查开始:" + str(dt.now()))
    f = be.check_col_same()
    be.write_unit_check("录入模块", f, "C3")
    del be
    del BaseExcel
    print("检查结果写入完成" + str(dt.now()))

def thread_write_stoptime(t1, process_op_dict, controller_end, controller_use):

    cf = process_op_dict["current_file"]
    error = process_op_dict["error"]

    while 1:
        if cf != process_op_dict["current_file"]:
            print("当前处理文件数:", process_op_dict["current_file"])
            cf = process_op_dict["current_file"]
        if process_op_dict["end_time"] != t1:
            t2 = process_op_dict["end_time"]
            controller_end.update(t2)
            controller_use.update(t2 - t1)
            print(f"结束时间为{t2}")
            print("完成读取文件数量为：" + str(process_op_dict["file_count"]) + "个")
            print("数据库读取完成")
            t1=t2

        if process_op_dict["error"] != error:
            print(process_op_dict["error"])
            error = process_op_dict["error"]

        time.sleep(0.1)


if __name__ == "__main__":

    try:
        default_name_flag=0
        excel_name = sys.argv[1]
    except:
        excel_name = "提取录入增100字段完整宏(3).xlsm"
        default_name_flag=1
    multiprocessing.freeze_support()  # 支持多进程
    sg.theme('LightGreen1')  # 设置当前主题

    text_size = (40, 1)
    # 界面布局，将会按照列表顺序从上往下依次排列，二级列表中，从左往右依此排列
    layout = [

        [sg.Text('启动时间', size=text_size), sg.InputText(key="start_time", readonly=True)],
        [sg.Text('结束时间', size=text_size), sg.InputText(key="end_time", readonly=True)],
        [sg.Text('耗时', size=text_size), sg.InputText(key="use_time", readonly=True)],
        [sg.Text('当前工作薄', size=text_size), sg.InputText(key="current_book", )],
        [sg.Text('当前工作表', size=text_size), sg.InputText(key="current_sheet", readonly=True)],
        [sg.Text('输出信息', size=text_size)],
        [sg.Output(size=(300, 8))],
        [  # sg.Button('修改配置', key="change_config", ),
            sg.Button('启动读取数据内容', key="start", visible=False),
            sg.Button('停止读取进程', key="stop_read_process"),
            sg.Button('写入校对结果', key="write_result", visible=False),
            sg.Button('关闭窗口', key="close")
        ],
        [
            sg.Button('表检查数据库写入', key="sheet_write_db", ),
            sg.Button('表检查 数据 写入', key="sheet_write_data", ),
            sg.Button('表检查 结果 写入', key="sheet_write_result", ),
            sg.Button('删除字段数据', key="delete_field_data", ),
        ], [
            #sg.Button('odbc数据写入', key="odbc_write_excel", ),
            sg.Button('excel写入并保存', key="write_excel_save", ),
        ]]

    # 创造窗口
    window = sg.Window('Window Title', layout, size=(800, 430))
    process_op_dict = Manager().dict()
    process_op_dict["stop_process"] = 0
    process_op_dict["current_file"] = 0
    process_op_dict["error"] = ""
    process_op_dict["file_count"] = 0

    # 事件循环并获取输入值
    while True:
        try:
            event, values = window.read()

            if event == "start":
                process_op_dict["stop_process"] = 0
                from BaseExcelOP import BaseExcel

                try:
                    be = BaseExcel(excel_name)
                    sheet_name = be.get_sheet_name("录入模块")
                    file_list = be.get_files_name()
                    window['current_book'].update(excel_name)
                    window['current_sheet'].update(sheet_name)
                    del be
                except:
                    print(f"工作薄{excel_name}不存在或被占用，请检查或稍等")
                    continue
                del BaseExcel
                if not sheet_name:
                    print("表格中F3没有表号")
                    continue
                t1 = datetime.now()
                del_mango_col()
                print("数据删除完成")
                window['start_time'].update(t1)

                p = multiprocessing.Process(target=process, args=(
                    process_op_dict, file_list, sheet_name, gui_get_max_process(excel_name),
                ))
                p.start()

                process_op_dict["end_time"] = t1
                t_ = Thread(target=thread_write_stoptime,
                            args=(t1, process_op_dict, window['end_time'], window['use_time']))
                t_.start()

                # process(process_op_dict["stop_process"],gui_get_path(excel_name),sheet_name, gui_get_max_process(excel_name),100)
                # t2 = datetime.now()
                # window['end_time'].update(t2)
                # window['use_time'].update(t2 - t1)
                # print("数据写入数据库完成")
            if event == "write_result":
                from BaseExcelOP import BaseExcel

                baseExcelName = excel_name
                baseExcelSheet = "格式"

                be = BaseExcel(baseExcelName)
                sheet_name = be.get_sheet_name("录入模块")
                be.write_result_sheet(baseExcelSheet, "录入模块", "N",
                                      "S", "R",
                                      mango_col_find_fields(sheet_name), sheet_name="", start_row=2)
                """
    
                try:
                    be = BaseExcel(excel_name)
                    try:
                        sheet_name = be.get_sheet_name("录入模块")
                        be.write_result_sheet("格式", "录入模块", "N",
                                              "S", "R",
                                              mango_col_find_fields(sheet_name), sheet_name="", start_row=2)
                        print("写入完成")
                    except:
                        print("写入失败")
                        continue
                except:
                    print("表格未打开或不存在")
                    continue
                """
            if event == "stop_read_process":
                process_op_dict["stop_process"] = 1
                print("收到手动停止操作指令")

            """
            if event == "sheet_write_db":
                process_op_dict["stop_process"] = 0
                from BaseExcelOP import BaseExcel

                try:
                    be = BaseExcel(excel_name)
                    books = be.get_books_name()
                    sheet_name = be.get_sheet_name("录入模块")
                    max_process = be.get_process()
                    execl_path = be.get_book_path()
                    cells = list(be.op_sheet.range("F3").expand("down").value)
                    window['current_book'].update(excel_name)
                    window['current_sheet'].update(sheet_name)
                    del be

                except:
                    print(f"工作薄{excel_name}不存在或被占用，请检查或稍等")
                    continue
                del BaseExcel
                if not sheet_name:
                    print("表格中F3没有表号")
                    continue
                t1 = datetime.now()
                del_mango_col()
                print("数据删除完成")
                window['start_time'].update(t1)
                p = multiprocessing.Process(target=process_read_excel_data_by_field, args=(
                    process_op_dict, execl_path, books, sheet_name, cells, max_process,
                ))
                p.start()
                process_op_dict["end_time"] = t1
                t_ = Thread(target=thread_write_stoptime,
                            args=(t1, process_op_dict, window['end_time'], window['use_time']))
                t_.start()
            """
            if event == "sheet_write_result":
                sheet_write_result_thread=Thread(target=sheet_write_result,args=(excel_name,))
                sheet_write_result_thread.start()
                """
                m = MongoDBOP()
                from BaseExcelOP import BaseExcel
                be = BaseExcel(excel_name)
                field_values = be.get_field_value_list("录入模块", "F", 3)
                books = be.get_books_name()
                #
                be.op_sheet.range("C3").expand("down").value = []
                be.op_sheet.range("D3").expand("down").value = []
                be.op_sheet.range("E3").expand("down").value = []
                print("数据清除完成")
                print("检查开始:" + str(dt.now()))
                # f = mongodbop.thread_read_check_result(excel_name)
                f = be.check_col_same()
                be.write_unit_check("录入模块", f, "C3")
                del be
                del BaseExcel
                print("检查结果写入完成" + str(dt.now()))
                """

            if event == "sheet_write_data":
                write_data_t=Thread(target=write_sheet_data,args=(excel_name,))
                write_data_t.start()

            if event == "delete_field_data":
                from BaseExcelOP import BaseExcel
                be = BaseExcel(excel_name)
                print("数据删除开始:" + str(dt.now()))
                be.del_fields_data()
                print("数据删除结束:" + str(dt.now()))
                del BaseExcel
                del be

            if event == "odbc_write_excel":
                process_op_dict["stop_process"] = 0
                if default_name_flag:
                    excel_name = values["current_book"]
                else:
                    window['current_book'].update(excel_name)
                print("开始写入时间:", datetime.now())
                t1 = datetime.now()
                process_op_dict["end_time"] = t1
                t_ = Thread(target=thread_write_stoptime,
                            args=(t1, process_op_dict, window['end_time'], window['use_time']))
                t_.start()
                window['start_time'].update(t1)
                p = multiprocessing.Process(target=sql_excel_no_use.write_file_data, args=(process_op_dict, excel_name))
                p.start()

            if event == "write_excel_save":
                path_=os.getcwd()
                subprocess.Popen([path_ + "/mul_save_excel/venv/scripts/python.exe ",path_+"/mul_save_excel/guis.py ",excel_name],shell=True)
                """
                process_op_dict["stop_process"] = 0
                if default_name_flag:
                    excel_name = values["current_book"]
                else:
                    window['current_book'].update(excel_name)
                print("开始写入时间:", datetime.now())
                t1 = datetime.now()
                process_op_dict["end_time"] = t1
                t_ = Thread(target=thread_write_stoptime,
                            args=(t1, process_op_dict, window['end_time'], window['use_time']))
                t_.start()
                window['start_time'].update(t1)
                t__=Thread(target=win32_excel.write_file_data,args=(process_op_dict, excel_name))
                t__.start()
                """
                #p = multiprocessing.Process(target=win32_excel.write_file_data, args=(process_op_dict, excel_name))
                #p.start()

            if event in (None, 'close'):  # 如果用户关闭窗口或点击`Cancel`
                break
            # print('You entered ', values[0])
        except Exception as e:
            print(traceback.print_exc())
            continue


    window.close()

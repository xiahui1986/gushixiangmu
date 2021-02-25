import PySimpleGUI as sg
from functions import *
from datetime import datetime

"""
import multiprocessing

if __name__ == "__main__":
    multiprocessing.freeze_support() #支持多进程
    sg.theme('LightGreen1')  # 设置当前主题
    cp = get_cp()
    text_size = (40, 1)
    # 界面布局，将会按照列表顺序从上往下依次排列，二级列表中，从左往右依此排列
    layout = [
        [sg.Text('mangodb连接字符串', size=text_size), sg.InputText(key='client', default_text=get_mango_client_name())],
        [sg.Text('mangod数据库名', size=text_size), sg.InputText(key='db', default_text=get_mango_db_name())],
        [sg.Text('mangod数据集合名', size=text_size), sg.InputText(key='col', default_text=get_mango_col_name())],

        [sg.Text('excel驱动', size=text_size), sg.InputText(key='driver', default_text=cp.get("excel", "driver"))],
        [sg.Text('excel数据路径', size=text_size), sg.InputText(key='bookPath', default_text=cp.get("excel", "bookPath"))],
        [sg.Text('基础excel工作薄名', size=text_size),
         sg.InputText(key='baseExcelName', default_text=cp.get("excel", "baseExcelName"))],
        [sg.Text('基础excel工作薄名比对表名', size=text_size),
         sg.InputText(key='baseExcelSheet', default_text=cp.get("excel", "baseExcelSheet"))],
        [sg.Text('excel数据比对表名', size=text_size),
         sg.InputText(key='compareSheet', default_text=cp.get("excel", "compareSheet"))],

        [sg.Text('基础excel工作薄写入表名', size=text_size),
         sg.InputText(key='name', default_text=cp.get("writeSheet", "name"))],
        [sg.Text('基础excel工作薄写入表，写入开始行号', size=text_size),
         sg.InputText(key='writeStartRow', default_text=cp.get("writeSheet", "writestartrow"))],
        [sg.Text('基础excel工作薄写入表，表名列号', size=text_size),
         sg.InputText(key='bookCol', default_text=cp.get("writeSheet", "bookCol"))],
        [sg.Text('基础excel工作薄写入表，检查结果列号', size=text_size),
         sg.InputText(key='resultCol', default_text=cp.get("writeSheet", "resultCol"))],

        [sg.Text('使用线程数', size=text_size), sg.InputText(key='processes', default_text=cp.get("process", "processes"))],
        [sg.Text('允许最大文件数量', size=text_size), sg.InputText(key='maxFile', default_text=cp.get("process", "maxfile"))],

        [sg.Text('启动时间', size=text_size), sg.InputText(key="start_time", readonly=True)],
        [sg.Text('结束时间', size=text_size), sg.InputText(key="end_time", readonly=True)],
        [sg.Text('耗时', size=text_size), sg.InputText(key="use_time", readonly=True)],
        [sg.Text('输出信息', size=text_size)],
        #[sg.Output(size=(300, 8))],
        [sg.Button('修改配置', key="change_config", ),
         sg.Button('启动读取数据内容', key="start", ),
         sg.Button('写入校对结果', key="write_result"),
         sg.Button('关闭窗口', key="close")]]

    # 创造窗口
    window = sg.Window('Window Title', layout, size=(800, 640))
    # 事件循环并获取输入值
    while True:
        event, values = window.read()
        if event == 'change_config':
            write_cp("mongo", "client", values['client'])
            write_cp("mongo", "db", values['db'])
            write_cp("mongo", "col", values['col'])
            write_cp("excel", "driver", values['driver'])
            write_cp("excel", "bookPath", values['bookPath'])

            write_cp("excel", "baseExcelName", values['baseExcelName'])
            write_cp("excel", "baseExcelSheet", values['baseExcelSheet'])
            write_cp("excel", "compareSheet", values['compareSheet'])

            write_cp("writeSheet", "name", values['name'])
            write_cp("writeSheet", "writeStartRow", values['writeStartRow'])
            write_cp("writeSheet", "bookCol", values['bookCol'])
            write_cp("writeSheet", "resultCol", values['resultCol'])
            write_cp("process", "processes", values['processes'])
            sg.popup("写入成功")

        if event == "start":
            t1 = datetime.now()
            del_mango_col()
            print("数据删除完成")
            window['start_time'].update(t1)
            process(values['bookPath'], int(values['processes']), int(values['maxFile']))
            t2 = datetime.now()
            window['end_time'].update(t2)
            window['use_time'].update(t2 - t1)
            print("数据写入数据库完成")
        if event == "write_result":
            from BaseExcelOP import BaseExcel
            """
"""
            be = BaseExcel(values['baseExcelName'])
            be.write_result_sheet(values['baseExcelSheet'], values['name'], values['bookCol'],
                                          values['resultCol'],
                                          mango_col_value(),start_row=int(values['writeStartRow'] ))
            """
"""
            try:
                be = BaseExcel(values['baseExcelName'])
                try:
                    be.write_result_sheet(values['baseExcelSheet'], values['name'], values['bookCol'],
                                          values['resultCol'],
                                          mango_col_value(),start_row=values['writeStartRow'] )
                    print("写入完成")
                except:
                    print("写入失败")
                    continue
            except:
                print("表格未打开或不存在")
                continue



        if event in (None, 'close'):  # 如果用户关闭窗口或点击`Cancel`
            break
        # print('You entered ', values[0])

    window.close()


"""
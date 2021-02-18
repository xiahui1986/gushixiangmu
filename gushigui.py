import PySimpleGUI as sg
from functions import *
sg.theme('LightGreen1')   # 设置当前主题
cp=get_cp()
# 界面布局，将会按照列表顺序从上往下依次排列，二级列表中，从左往右依此排列
layout = [  [sg.Text('mangodb连接字符串'), sg.InputText(key='client',default_text=get_mango_client_name())],
            [sg.Text('mangod数据库名'), sg.InputText(key='db',default_text=get_mango_db_name())],
            [sg.Text('mangod数据集合名'), sg.InputText(key='col',default_text=get_mango_col_name())],

            [sg.Text('excel驱动'), sg.InputText(key='driver',default_text=cp.get("excel","driver"))],
            [sg.Text('excel数据路径'), sg.InputText(key='bookPath',default_text=cp.get("excel","bookPath"))],
            [sg.Text('基础excel工作薄名'), sg.InputText(key='baseExcelName',default_text=cp.get("excel","baseExcelName"))],
            [sg.Text('基础excel工作薄名比对表名'), sg.InputText(key='baseExcelSheet',default_text=cp.get("excel","baseExcelSheet"))],
            [sg.Text('excel数据比对表名'), sg.InputText(key='compareSheet',default_text=cp.get("excel","compareSheet"))],

            [sg.Text('基础excel工作薄写入表名'), sg.InputText(key='name',default_text=cp.get("writeSheet","name"))],
            [sg.Text('基础excel工作薄写入表，表名列号'), sg.InputText(key='bookCol',default_text=cp.get("writeSheet","bookCol"))],
            [sg.Text('基础excel工作薄写入表，检查结果列号'), sg.InputText(key='resultCol',default_text=cp.get("writeSheet","resultCol"))],

            [sg.Text('启动时间'), sg.InputText(key="start_time",readonly=True),sg.Text('结束时间'),
             sg.InputText(key="end_time",readonly=True),sg.Text('耗时'),sg.InputText(key="use_time",readonly=True)],

            [sg.Button('修改配置',key="change_config" ,),sg.Button('启动读取数据内容',key="start" ,),sg.Button('写入校对结果',key="end"), sg.Button('关闭窗口',key="close")] ]

# 创造窗口
window = sg.Window('Window Title', layout)
# 事件循环并获取输入值
while True:
    event, values = window.read()
    if event=='change_config':

        sg.popup(write_cp("mongo","client",values['client']))

    if event in (None, 'close'):   # 如果用户关闭窗口或点击`Cancel`
        break
    #print('You entered ', values[0])

window.close()
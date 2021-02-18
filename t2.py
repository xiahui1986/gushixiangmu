import win32com.client

import adodbapi as ab

file_name="E:/股市数据处理需求编程/shyg/shyg/1-aaaaa.xlsx"
file_name="'E:\zhj\zhj_vba_proj\wip hours -.xlsx'"
DSN = " Provider=Microsoft.ACE.OLEDB.12.0;data source='E:\zhj\zhj_vba_proj\wip hours -.xlsx'  ;Extended Properties='Excel 14.0 Xml;HDR=YES' "
DSN=r"Provider=MSDASQL;Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};ReadOnly=False;Dbq='E:\zhj\zhj_vba_proj\wip hours -.xlsx'"
strSql="SELECT * FROM [Sheet1$]"
conn=ab.connect(connection_string=DSN)
cursor=conn.cursor()
cursor.execute(strSql)
result=cursor.fetchall()
pass



cn=win32com.client.Dispatch("ADODB.Connection")
cn.open(DSN)

Rs = win32com.client.Dispatch("ADODB.Recordset")
Rs.open(strSql,cn,3,3)



Rs.ActiveConnection = DSN
Rs.Source = r"SELECT * FROM [Sheet1$]"
Rs.CursorType = 0
Rs.CursorLocation = 2
Rs.LockType = 1
Rs.Open()
numRows = 0
while not Rs.EOF:
    print(r'id:',Rs.Fields.Item("id").Value.encode('gbk'))
    if Rs.Fields.Item("name").Value != None:
        print(r'  name:',Rs.Fields.Item("name").Value.encode('gbk'))
    numRows+=1
    Rs.MoveNext()
print('Total Rows:',numRows)
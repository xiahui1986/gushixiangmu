import win32com.client
file_name="E:/股市数据处理需求编程/shyg/shyg/1-aaaaa.xlsx"
DSN = " Provider=Microsoft.ACE.OLEDB.12.0;data source= "   +  file_name  +   " ;Extended Properties='Excel 12.0 Xml;HDR=YES' "
Rs = win32com.client.Dispatch("ADODB.Recordset")
Rs.ActiveConnection = DSN
Rs.Source = r"SELECT * FROM dbo.Sheet1"
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
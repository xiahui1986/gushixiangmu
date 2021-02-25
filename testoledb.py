if __name__=="__main__":
    import win32com.client as wc
#
    conn = wc.Dispatch(r'ADODB.Connection')
    #C:\Users\xiah\Desktop\新建 Microsoft Excel 工作表.xlsx
    #E:\股市数据处理需求编程\shyg\shyg\1 - aaaaa.XLSX
    DSN = r"PROVIDER=Microsoft.ACE.OLEDB.12.0;DATA SOURCE=C:\Users\xiah\Desktop\新建 Microsoft Excel 工作表.xlsx;Persist Security Info=False;Extended Properties=Excel 12.0"
    conn.Open(DSN)

    rs=wc.Dispatch(r"ADODB.Recordset")
    rs.Open("select * from [Sheet1$]",conn,1,3)
    for i in range(rs.Fields.Count):
        print(rs(i).Name)

    pass
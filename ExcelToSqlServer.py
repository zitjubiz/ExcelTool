import xlrd  # https://xlrd.readthedocs.io/en/latest/
import pyodbc

cnstr = pyodbc.connect('''DRIVER={SQL Server};SERVER=192.168.0.1;  DATABASE=db;UID=sa;PWD=sa''')
cursor = cnstr.cursor()

data = xlrd.open_workbook('02.01.xlsx')
table = data.sheets()[0]
print(table.nrows)
# 忽略第一行行头
for i in range(table.nrows):
    if (i <= 0): continue


    shop = str(table.cell(i, 1).value).strip()
    recvName = str(table.cell(i, 2).value).strip()
    phone = str(table.cell(i, 3).value)
    address = str(table.cell(i, 4).value).strip()
    province = str(table.cell(i, 5).value).strip()
    city = str(table.cell(i, 6).value).strip()
    region = str(table.cell(i, 7).value).strip()
    jieduan = str(table.cell(i, 8).value).strip()
    cusid =''
    for row in cursor.execute("select ccusid from customer where ccusphone='" + phone + "'"):
        cusid = row.ccusid

    print("""INSERT INTO [dbo].[Materials]([cCusID],[Mproid],[iDot],[sendDate] ,[isUse]
           ,[address] ,[phone],[pients],[icnt],[cstate],[ccity],[ccity2],[shop],[jieduan])VALUES(""")
    print("'", cusid, "',1177,600,getdate(),0,'", address, "','", phone, "','", recvName,
          "',1,'", province, "','", city, "','", region, "','", shop, "','", jieduan, "')", sep='')

print('请把sql语句在数据库运行')

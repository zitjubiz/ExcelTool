import xlrd
import csv
with open('super.csv', 'w') as csvfile:
    writer = csv.writer(csvfile,lineterminator='\n')
    data = xlrd.open_workbook('supershop.xlsx')
    table = data.sheets()[0]

    for i in range(table.nrows-1):
        if(i==0):
            continue
        shopid=str(table.cell(i+1,4).value)[:-2]
        for j in range(12):
            if(str(table.cell(i+1,10+j).value).strip()==''):
                baseSaleCnt = 0
            else:
                baseSaleCnt= int(table.cell(i+1,10+j).value)
                
            if(str(table.cell(i+1,23+j).value).strip()==''):
                baseSaleCntEx = 0
            else:
                baseSaleCntEx= int(table.cell(i+1,23+j).value)
            
            #print(shopid,",",baseSaleCnt,",",baseSaleCntEx,",0.3,0,0,",j+1,",0,0,0,0,0,0")    	
            writer.writerow([shopid, baseSaleCnt, baseSaleCntEx,'0.3','0','0',j+1,'0','0','0','0','0','0'])

    print('导出到super.csv,请用bcp导入数据库')        
    print('bcp huodong_wx20180101_shop  in "super.csv"  -c -t "," -S 192.168.0.1 -d dbName -U "UserName" -P "pwd"') 

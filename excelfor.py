# -*-coding:utf-8-*-

import xlrd
import xlwt
import sys

#新建excel
newexcel = xlwt.Workbook()
tablenew = newexcel.add_sheet("sheet",cell_overwrite_ok=True)
#读取所有原excel sheet
excel = xlrd.open_workbook('/tmp/excel.xls')
table0 = excel.sheets()[0]
table1 = excel.sheets()[1]
table2 = excel.sheets()[2]
table3 = excel.sheets()[3]

#获取原excel所有sheet行数
nrow0 = table0.nrows
nrow1 = table1.nrows
nrow2 = table2.nrows
nrow3 = table3.nrows

#获取原excel第一个sheet列数
ncol0 = table0.ncols

#新表复制盘点表数据
s = 0
for y in range(0,ncol0):
	for x in range(0,nrow0):
		value0 = table0.cell(x,y).value
		tablenew.write(x,y,value0)
		s = s + 1
		#print value0
	newexcel.save('/tmp/newxls.xls')
print("新表创建完成！")
print("开始写入库存数据！")	
#获取原excel第一个sheet所有编码并与第二，三，四个sheet编码对比，如果编码相同则获取该sheet需要的值。没有相同的则为0
for i in range(2,nrow0):
	code0 = str(table0.cell(i,1).value)
	number1 = str(0)
	number2 = str(0)
	number3 = str(0)

	for j in range(1,nrow1):
		code1 = str(table1.cell(j,1).value)
		if code0 == code1:
			number1 = int(table1.cell(j,5).value)
			break
	for k in range(1,nrow2):
		code2 = str(table2.cell(k,0).value)
		if code0 == code2:
			number2 = int(table2.cell(k,4).value)
			break
	for l in range(1,nrow3):
		code3 = str(table3.cell(l,1).value)
		if code0 == code3:
			number3 = int(table3.cell(l,5).value)
			break
	#写入新表所需数据并保存
	tablenew.write(i,7,number1)
	tablenew.write(i,8,number2)
	tablenew.write(i,9,number3)
	newexcel.save('/tmp/newxls.xls')
	#print("已写入："+str(100*(i+10)/nrow0)+"%")
	sys.stdout.write("已写入："+str(100*(i+10)/nrow0)+"%"+"\r")
	#print("正在写入："+"产品编码："+code0+"天合库存："+str(number1)+"曜居库存："+str(number2)+"残次库存："+str(number3))
print("获取数据结束！")
		

		
		

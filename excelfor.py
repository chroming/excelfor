# -*-coding:gbk-*-

#���������ڻ�ȡ��һ��sheet��Ŀ���ڵڶ����ĸ�sheet��Ѱ�Ҷ�Ӧ��ֵ��д���±�


import xlrd
import xlwt
import sys

print("�뽫��Ҫ�����excel���뱾������ͬĿ¼�¡�")
name = raw_input("������Ҫ�����excel�ļ���(������չ��xls)�����������ʼ����")
print("��ʼ�����±�")
#�½�excel
newexcel = xlwt.Workbook()
tablenew = newexcel.add_sheet("%s"%name,cell_overwrite_ok=True)
#��ȡ����ԭexcel sheet
excel = xlrd.open_workbook('*.xls')
table0 = excel.sheets()[0]
table1 = excel.sheets()[1]
table2 = excel.sheets()[2]
table3 = excel.sheets()[3]

#��ȡԭexcel����sheet����
nrow0 = table0.nrows
nrow1 = table1.nrows
nrow2 = table2.nrows
nrow3 = table3.nrows

#��ȡԭexcel��һ��sheet����
ncol0 = table0.ncols

#�±����̵������
s = 0
for y in range(0,ncol0):
	for x in range(0,nrow0):
		value0 = table0.cell(x,y).value
		tablenew.write(x,y,value0)
		s = s + 1
		#print value0
	newexcel.save('newxls.xls')
print("�±�����ɣ�")
print("��ʼд�������ݣ�")	
#��ȡԭexcel��һ��sheet���б��벢��ڶ��������ĸ�sheet����Աȣ����������ͬ���ȡ��sheet��Ҫ��ֵ��û����ͬ����Ϊ0
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
	#д���±��������ݲ�����
	tablenew.write(i,7,number1)
	tablenew.write(i,8,number2)
	tablenew.write(i,9,number3)
	newexcel.save('newxls.xls')
	#print("��д�룺"+str(100*(i+10)/nrow0)+"%")
	sys.stdout.write("��д�룺"+str(100*(i+10)/nrow0)+"%"+"\r")
	#print("����д�룺"+"��Ʒ���룺"+code0+"��Ͽ�棺"+str(number1)+"�׾ӿ�棺"+str(number2)+"�дο�棺"+str(number3))
print("��ȡ���ݽ��������newxls.xls�鿴��ȡ�����")
		

		
		

import xlrd
import xlwt
import xlsxwriter
from xlutils.copy import copy

dataset = xlrd.open_workbook('/Users/zhaoyuxuan/Desktop/normalize.xlsx')
data = dataset.sheets()[0]
workbook = xlsxwriter.Workbook('output.xlsx')
worksheet = workbook.add_worksheet()
row=-1
k=0
for n in data.col_values(3):
	row +=1
	lng=data.cell_value(row,0)
	lat=data.cell_value(row,1)
	n=round(n,3)
	if n!=0:
		worksheet.write(k,0,lng)
		worksheet.write(k+1,0,lat)
		worksheet.write(k+2,0,n)
		k=k+3


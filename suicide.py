import xlrd
import xlwt
import xdrlib,sys
import xlsxwriter
#open excel file 
data = xlrd.open_workbook('../Desktop/data.xlsx')
		
#get sheet
rate = data.sheets()[0]
identifier = data.sheets()[1]
population = data.sheets()[2]
i=-1

workbook = xlsxwriter.Workbook('test.xlsx')
worksheet = workbook.add_worksheet()
for n in rate.col_values(1):
	i=i+1
	for row in range(0,identifier.nrows):
		for col in range(0,identifier.ncols):
			if n==identifier.cell_value(row,col):
				m=population.cell_value(row,col)*rate.cell_value(i,2)
				# //divide 86400 by the output from the above calculation
				m=m/365
				m=86400/m
				worksheet.write_number(row,col,m)

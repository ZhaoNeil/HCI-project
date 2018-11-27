import xlrd
import xlwt
import xdrlib,sys
import xlsxwriter
#open excel file 
data = xlrd.open_workbook('../Desktop/dataset suicide.xlsx')
		
#get sheet
rate = data.sheets()[3]
identifier = data.sheets()[1]
population = data.sheets()[2]

def GetPopulation():
	i=-1
	workbook = xlsxwriter.Workbook('suicide location.xlsx')
	worksheet = workbook.add_worksheet()
	for n in rate.col_values(1):
		i=i+1
		for row in range(0,identifier.nrows):
			for col in range(0,identifier.ncols):
				if n==identifier.cell_value(row,col):
					m=population.cell_value(row,col)*rate.cell_value(i,2)
					m=m/365
					m=m/1000
					m=86400/m
					if m>0:
						worksheet.write_number(row,col,m)
	return

GetPopulation()

suicide_population = xlrd.open_workbook('../Desktop/suicide location.xlsx')
location = suicide_population.sheets()[0]

def GetLocation():
	i=-181
	row=0
	col=0
	workbook = xlsxwriter.Workbook('location_lat&lng.xlsx')
	worksheet = workbook.add_worksheet()
	for m in range(0,location.ncols):
		i=i+1
		j=91
		for n in range(0,location.nrows):
			j=j-1
			k=location.cell_value(n,m)
			if k!="":
				worksheet.write_number(row,col,i)
				worksheet.write_number(row,col+1,j)
				worksheet.write_number(row,col+2,k)
				row=row+1
	return

GetLocation()


import xlrd 
from xlwt import Workbook

loc = ("C:/Users/Umarbek Nasimov/Desktop/Oldways/StoreOrdersExport_Example.xlsx")

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0) #make a new sheet object
 
num_of_rows=sheet.nrows #number of rows
num_of_cols=sheet.ncols #number of columns
print(num_of_rows, num_of_cols)
print(sheet.cell_value(1,1)) #get the specific text at row 1, col 1

for i in range(num_of_cols):
	print(sheet.cell_value(0,i)) #print all the column names 


wb_new = Workbook()
sheet1 = wb_new.add_sheet("Modified Workbook")

sheet1.write(0,0, "Umarbek")
wb_new.save("test.xls")
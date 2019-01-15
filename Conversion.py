import xlrd 
from xlwt import Workbook


loc_order_report = ("StoreOrdersExport_Example.xlsx")
loc_sku_map = ("SKU_class_item.xlsx")

wb = xlrd.open_workbook(loc_order_report)
sheet = wb.sheet_by_index(0) #make a new sheet object


num_of_rows=sheet.nrows #number of rows
num_of_cols=sheet.ncols #number of columns
print(num_of_rows, num_of_cols)
print(sheet.cell_value(1,1)) #get the specific text at row 1, col 1

for i in range(num_of_cols):
	print(sheet.cell_value(0,i)) #print all the column names 


wb_new = Workbook()
sheet1 = wb_new.add_sheet("Modified Workbook")

sheet1.write(0,0, "Customer")
sheet1.write(0,1, "Date")
sheet1.write(0,2, "Ref No.")
sheet1.write(0,3, "Class")
sheet1.write(0,4, "Payment method")
sheet1.write(0,5, "Memo")
sheet1.write(0,6, "Item")
sheet1.write(0,7, "Quantity")
sheet1.write(0,8, "Amount")
sheet1.write(0,9, "Amount of Sales Receipt")
sheet1.write(0,10, "Amount of transaction")
sheet1.write(0,11, "Amount Deposited")
sheet1.write(0,12, "Date deposited to CTC")
sheet1.write(0,13, "Template Name")

wb_new.save("test.xls")
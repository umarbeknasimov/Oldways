# import xlrd # read xlsx or xls files
# from xlwt import Workbook # write xls files
from openpyxl import Workbook, load_workbook # read and modify Excel 2010 files


loc_order_report = ("StoreOrdersExport_Example.xlsx")
loc_sku_map = ("SKU_class_item.xlsx")

wb_order_report = load_workbook(loc_order_report)
# sheet = wb_order_report.sheet_by_index(0) #make a new sheet object

wb_sku_map = load_workbook(loc_sku_map)

'''
num_of_rows=sheet.nrows #number of rows
num_of_cols=sheet.ncols #number of columns
print(num_of_rows, num_of_cols)
print(sheet.cell_value(1,1)) #get the specific text at row 1, col 1

for i in range(num_of_cols):
	print(sheet.cell_value(0,i)) #print all the column names 
'''

wb_new = Workbook()
# sheet1 = wb_new.add_sheet("Modified Workbook") # xlwt 
ws1 = wb_new.active
ws1.title = "Sales Receipts"


# write column names
column_names = ['Customer', 'Date', 'Ref No.', 'Class', 'Payment method', 'Memo', 'Item', 'Quantity', 'Amount', 'Amount of Sales Receipt', 'Amount of transaction', 'Amount Deposited', 'Date deposited to CTC', 'Template Name']

for i in range(len(column_names)):
	ws1.cell(0+1, i+1, column_names[i])

 # for each row in store order report
 	# for each number in total quanitity
 		# customer
 		# date
 		# ref no
 		# search up class from sku
 		# payment
 		# memo
 		#item based on sku
 		# quantity unknown
 		# amount of sales received unknown
 		# amount of transaction = cost of total order
 		# amount deposited blank
 		# column M blank
 		template name = Customer Sales Receipt



wb_new.save("Sales Receipts.xlsx")
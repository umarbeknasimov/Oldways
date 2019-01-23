# import xlrd # read xlsx or xls files
# from xlwt import Workbook # write xls files
from openpyxl import Workbook, load_workbook # read and modify Excel 2010 files


class Product:
    def __init__(self, str_):
        """Initializes a product from a comma-spliced string"""
        # First, make a dictionary of the fields
        field_list = [[item.strip() for item in field_value.split(':')]\
                      for field_value in str_.split(',')]
        field_dict = {field: value for field, value in field_list}

        # Then get the necessary fields from it
        self.sku = field_dict["Product SKU"]

        #indexing sku table:
        sku_row = 0
        for i in range(1, sku_sheet.max_row + 1):
            if sku_sheet['C'+str(i)].value == self.sku:
                sku_row = i

        self.class_ = '' if sku_row == 0 else sku_sheet['A'+str(sku_row)].value 
        self.item = '' if sku_row == 0 else sku_sheet['B'+str(sku_row)].value 
        self.quantity = field_dict["Product Qty"] # string because not used for calculations
        self.unit_amount = field_dict["Product Unit Price"]
        self.total_amount = field_dict["Product Total Price"]

    def write_data(self, ws, row_index):
        """Writes the data to the worksheet at row_index"""
        items = {"Class": self.class_,
                 "Item": self.item,
                 "Quantity": self.quantity,
                 "Amount": self.unit_amount}
        
        for field, value in items.items():
            ws.cell(row = row_index, column = Order.column_indexes[field]).value = value

        
class Order:
    column_names = ['Customer', 'Date', 'Ref No.', 'Class', 'Payment method', 'Memo', \
                    'Item', 'Quantity', 'Amount', 'Amount of Sales Receipt', 'Amount of transaction', \
                    'Amount Deposited', 'Date deposited to CTC', 'Template Name']
    column_indexes = {name: i + 1 for i, name in enumerate(column_names)}
    
    def __init__(self, ws, row_index):
        """Initializes an order from a worksheet and the row of that order"""
        # First, dictionary. The first row contains the names of the fields
        field_dict = {field: value for field, value in \
                      zip((cell.value for cell in \
                               list(ws.iter_rows(min_row = 1,         max_row = 1)        )[0]), \
                          (cell.value for cell in \
                               list(ws.iter_rows(min_row = row_index, max_row = row_index))[0]))}

        # Then, fields
        self.customer = "PRODUCTS"
        self.date = field_dict["Order Date"] # TODO: Figure out format, as there is a conflict
        self.ref_no = field_dict["Customer Name"]
        self.payment = field_dict["Payment Method"]
        self.memo = field_dict["Order ID"]
        self.total_amount = field_dict["Order Total (ex tax)"] # Tax is 0 for all examples given, though
        self.template = "Customer Sales Receipt"
        self.products = [Product(product) for product in field_dict["Product Details"].split('|')]

    def write_data(self, ws, row_index):
        """
        Writes the data to the worksheet at row_index, and
        returns the index of the row to insert the next order
        """
        items = {"Customer": self.customer,
                 "Date": self.date,
                 "Ref No.": self.ref_no,
                 "Payment method": self.payment,
                 "Memo": self.memo,
                 "Template Name": self.template}

        for i, product in enumerate(self.products):
            # Each row of the same order has the same overall data
            for field, value in items.items():
                ws.cell(row = row_index + i, column = Order.column_indexes[field]).value = value
            # but the products are different
            product.write_data(ws, row_index + i)

        # Total amount is on the last line of the order
        ws.cell(row = row_index + len(self.products) - 1, \
                column = Order.column_indexes["Amount of transaction"]).value = self.total_amount

        return row_index + len(self.products)


loc_order_report = ("DefaultOrderExportReport_Jan182019.xlsx")
loc_sku_map = ("SKU_class_item.xlsx")

wb_order_report = load_workbook(loc_order_report)
# sheet = wb_order_report.sheet_by_index(0) #make a new sheet object

wb_sku_map = load_workbook(loc_sku_map)
sku_sheet = wb_sku_map.worksheets[0]

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


for i, column_name in enumerate(Order.column_names):
    ws1.cell(0+1, i+1, column_name)


# for each row in store order report
    # for each number in total quanitity
        # customer
        # date
        # ref no
        # search up class from sku
        # payment
        # memo
        # item based on sku
        # quantity unknown
        # amount of sales received unknown
        # amount of transaction = cost of total order
        # amount deposited blank
        # column M blank
# template name = Customer Sales Receipt

orders = [Order(wb_order_report.active, row) for row in range(2, wb_order_report.active.max_row + 1)]
curr_row = 2
for order in orders:
    curr_row = order.write_data(ws1, curr_row)

wb_new.save("Sales Receipts.xlsx")

import os
import glob
import csv
from xlsxwriter.workbook import Workbook
#desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') 
#shutil.copy(txtName, desktop)
#print (os.path.join('c:', os.sep, 'sourcedir'))
for csvfile in glob.glob("*.csv"):
    workbook = Workbook(csvfile[:-4] + '.xlsx')
    worksheet = workbook.add_worksheet()
    with open(csvfile, 'rt', encoding='utf8') as f:
        reader = csv.reader(f)
        for r, row in enumerate(reader):
            for c, col in enumerate(row):
                worksheet.write(r, c, col)
        #try:
        workbook.close()
        #except:
	        #print("csv file does not exist")

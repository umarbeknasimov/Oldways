# Store Order Report to Sales Receipt Conversion Readme

### How to run Conversion.exe
1. Make sure the store order report .csv file you want to convert is saved in the same folder as Conversion.exe and SKU_class_item.xlsx, and make sure it is titled "OrderExport.csv". 
2. Run Conversion.exe by double clicking the file.
3. The formatted sales receipt .xlsx file will appear in the same folder as Conversion.exe under the title "Sales Receipts [MM-DD-YYYY] [HH-MM].xlsx". Running Conversion.exe will also produce an .xlsx version of the store order report.
4. After receiving the converted sales receipt file, the store order report .csv and .xlsx may be deleted if you like. Otherwise, rename the .csv file so that it does not conflict with any future .csv files you wish to convert.

### How to Update the SKU Chart

 1. Replace the current file SKU_class_item.xlsx with an updated SKU chart with the same name. Make sure column A in the new file represents Class, column B represents Item, and column C represents SKU.

### How to Change Conversion.exe/Conversion.py
1. Download Python 3
    1. Mac: Follow the instructions [here](https://ehmatthes.github.io/pcc/chapter_01/osx_setup.html) until you reach the beginning of the section about Sublime. The most recent version of Python (as of Jan 2019) is 3.7, not 3.5 as it was when this article was written, so you will see 3.7.x instead of 3.5.x in the last step.
    2. Windows: Go [here](to https://www.python.org/downloads/release/python-368/) and install "Windows x86-64 executable installer". During the installation, check the "Add Python 3.6 to PATH" box (important).
2. Install the openpyxl library
    1. Mac: type "pip install openpyxl" into Terminal
    2. Windows: run "python -m pip install openpyxl" in Command Line.
3. Install the xlsxwriter library
    1. Mac: type "pip install xlsxwriter" into Terminal
    2. Windows: run "python -m pip install xlsxwriter" in Command Line
4. Conversion.py contains comments (either preceded by # or surrounded by triple quotes) that indicate the function of nearby code. 
5. In particular, you may wish to change the file name of the store order report .csv or the file name of the resulting sales receipt .xlsx. Instructions on changing the file name of the .csv file to be converted are found in comments on line 10. Instructions on changing the name of the sales receipt .xlsx file are found in comments at the end of Conversion.py on lines 252-253.
6. After saving necessary edits to Conversion.py, you may choose to either run the .py file directly without converting to an .exe file, or you can convert to an .exe file.
    1. If you choose not to convert to .exe, then run Conversion.py by opening the file in an IDE (directions for installing Sublime Text found [here](https://ehmatthes.github.io/pcc/chapter_01/osx_setup.html)) and running from there. In Sublime, Ctrl+B (Windows) or Command+B (Mac) will run the file.
    2. If you choose to convert to .exe, then 
        1. Download PyInstaller by typing "pip install pyinstaller" into Terminal for Mac or "pip install pyinstaller" into Command Line for Windows.
        2. Type "pyinstaller -w -F Conversion.py" into Terminal for Mac or "pyinstaller -w -F Conversion.py" into Command Line for Windows.
        3. The .exe file should appear in your file explorer.

### Contact
Feel free to contact joshuaa@mit.edu, unasimov@mit.edu, and/or szhi@mit.edu if you have any questions.
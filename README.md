# Store Order Report to Sales Receipt Conversion Readme

### How to run Conversion.exe
1. Download Python 3
    1. Mac: Follow the instructions [here](https://ehmatthes.github.io/pcc/chapter_01/osx_setup.html) until you reach the beginning of the section about Sublime. The most recent version of Python (as of Jan 2019) is 3.7, not 3.5 as it was when this article was written, so you will see 3.7.x instead of 3.5.x in the last step.
    2. Windows:
2. Install the openpyxl library
    1. Mac: type "pip install openpyxl" into Terminal
    2. Windows: type "pip install openpyxl" into Command Line?
3. Make sure the store order report .csv file you want to convert is saved in the same folder as Conversion.exe and SKU_class_item.xlsx, and make sure it is titled "__[title]__.csv". 
4. Run Conversion.exe by double clicking the file.
5. The formatted sales receipt .xlsx file will appear in the same folder as Conversion.exe under the title "Sales Receipts [date] [time].xlsx". Running Conversion.exe will also produce an .xlsx version of the store order report.
6. After receiving the converted sales receipt file, the store order report .csv and .xlsx may be deleted if you like.

### How to Update the SKU Chart

 1. Replace the current file SKU_class_item.xlsx with an updated SKU chart with the same name. Make sure column A in the new file represents Class, column B represents Item, and column C represents SKU.

### How to Change Conversion.exe/Conversion.py
1. Conversion.py contains comments (either preceded by # or surrounded by triple quotes) that indicate the function of nearby code. 
2. In particular, you may wish to change the file name of the store order report .csv or the file name of the resulting sales receipt .xlsx. [instructions on how to change .csv name]. Instructions on changing the name of the sales receipt .xlsx file are found in comments at the end of Conversion.py [around line 250].
3. After saving necessary edits to Conversion.py, you may choose to either run the .py file directly without converting to an .exe file, or you can convert to an .exe file.
    1. If you choose not to convert to .exe, then run Conversion.py by opening the file in an IDE (directions for installing Sublime Text found [here](https://ehmatthes.github.io/pcc/chapter_01/osx_setup.html)) and running from there. In Sublime, Ctrl+B (Windows) or Command+B (Mac) will run the file.
    2. If you choose to convert to .exe, then type "_____" into Terminal for Mac or "____" into Command Line for Windows.

### Contact
Feel free to contact joshuaa@mit.edu, unasimov@mit.edu, and/or szhi@mit.edu if you have any questions.
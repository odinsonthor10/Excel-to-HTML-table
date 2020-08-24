##Author: Swapnankit  mailid: bswapnankit@gmail.com

This repository contains ExToHtml.py file
This file converts excel sheet to Html table and gives the .html file.

Use this function  toHtml(filename, sheetname,filesave,custom_row=False,custom_col=False,col_start,col_end,row_start, row_end)   to convert an excel file into html table

ExToHtml.py contains two functions 

1. toDictionary(filename, sheetname,custom_row=False,custom_col=False,col_start,col_end,row_start, row_end) 

filename=This function accepts excel file in .xls or .xlsx format. 

sheetname= sheet in the workbook you want to render in HTML 

custom_row= set True if you want to give start row nad end row number. If custom_row is True, set row_start and row_end

custom_col= set True if you want to give start column nad end column number. If custom_col is True, set col_start and col_end

THIS FUNCTION RETURNS DICTIONRY OF EXCEL SHEET CONTAIN, WHICH WILL BE USED LATER FOR RENDERING HTML TABLE

2. toHtml(filename, sheetname,filesave,custom_row=False,custom_col=False,col_start,col_end,row_start, row_end)

filename=This function accepts excel file in .xls or .xlsx format. 

sheetname= sheet in the workbook you want to render in HTML 

filesave= give a name to save the html file. eg. some_file.html

custom_row= set True if you want to give start row nad end row number. If custom_row is True, set row_start and row_end

custom_col= set True if you want to give start column nad end column number. If custom_col is True, set col_start and col_end

The function will create and html file in the same folder

NOTE: The above function will execute faster if the file is in .xls format
      .xlsx files will give the same result but the processing time is way more than .xls.





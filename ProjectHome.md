This project is used to store several macros I wrote to import multi Excel workbooks as worksheets into an already-opened Excel workbook.

You can select a range in a worksheet(no matter continuous or discontinuous), which stores the names of worksheets to be imported. After executing the macro, worksheets will be created if it doesn’t exist, or it will be updated according to the specified  Excel worksheet in another Excel workbook. The first column is the serial number of worksheet, the second column denotes department name and the worksheet name should be like “01 Department A”.

The imported/created worksheets can be sorted by the selected serial number.

The macros are smart enough to match the Excel files to be imported. The name is like “01 XXXX\_Department  A\_20110707.xls”.

After importing, specified cells in every other worksheets can be automatically copied to cells on the same row next to department.
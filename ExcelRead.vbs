set objExcel= CreateObject("Excel.Application")
Set objWorkbook= objExcel.Workbooks.Open("D:\Sumit\Personal\SeleniumTest.xlsx")
Set objSheet=objWorkbook.Worksheets("SeExcelsheet")
columncount=objSheet.usedrange.columns.count
rowcount=objSheet.usedrange.rows.count
msgbox columncount
msgbox rowcount
for i=1 to columncount
for j = 1 to rowcount
fieldvalue = objSheet.cells(j,i)
'msgbox fieldvalue
objExcel.Cells(4,4)="sumit"
objExcel.ActiveWorkbook.Save
'below line throws error of existing file but saves file
'objWorkbook.SaveAs "c:\testing321.xls"
next
next
objExcel.ActiveWorkBook.Save
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit

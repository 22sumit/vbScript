Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook1= objExcel.Workbooks.Open("C:\test1.xls")
Set objWorkbook2= objExcel.Workbooks.Open("C:\test2.xls")

Set objRange = objWorkbook1.Worksheets("Sheet1").UsedRange.Copy
objWorkbook2.Worksheets("Sheet1").Range("A1").PasteSpecial objRange

objWorkbook1.Save
objWorkbook1.Close

objWorkbook2.Save
objWorkbook2.Close
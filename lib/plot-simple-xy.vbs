Include(".\parameters.vbs")
Include("..\Excel.vbs")
Dim xl
set xl = new ExcelPlotter
putil.TempBasePath = "."
wbFile = "..\workbooks\SimpleXYPlot.xlsm"
EchoX "Opening workbook at path: %x", wbFile
xl.OpenWorkBook(wbFile)
EchoX "Active workbook name is: %x", xl.GetActiveWorkbook.Name
call xl.SimpleXYPlot(data, destDir)
xl.CloseWorkBook
set xl = nothing
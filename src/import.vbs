Include(".\parameters.vbs")
Include(".\Excel.vbs")
Dim xl
set xl = new Excel
EchoX "Opening workbook at path: %x", wbFile
xl.OpenWorkBook(wbFile)
EchoX "Active workbook name is: %x", xl.GetActiveWorkbook.Name
xl.ImportVBAComponents(sourceDir)
xl.CloseWorkBook
set xl = nothing
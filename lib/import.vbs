Include(".\parameters.vbs")
Include("..\Excel.vbs")
Dim xl
set xl = new Excel
xl.OpenWorkBook("..\test\Excel_MVC_Creator.xlsm")
EchoX "Active workbook name is: %x", xl.GetActiveWorkbook.Name
xl.ImportVBAComponents(NULL)
xl.CloseWorkBook
set xl = nothing
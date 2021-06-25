call npx vbspm /file:lib\export.vbs /workbook:test\Excel_MVC_Creator.xlsm /destination:Excel_MVC_Creator 
call npx vbspm /file:lib\import.vbs /workbook:test\Excel_MVC_Creator.xlsm /source:Excel_MVC_Creator
call npx vbspm /file:lib\plot-simple-xy.vbs /workbook:workbooks\SimpleXYPlot.xlsm /destination:ChartPlots /workbook:dummy /debug:true /data:A,B,1,1,2,4,3,9,4,0
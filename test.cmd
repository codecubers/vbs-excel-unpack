call builder.cmd
call del /f .\ChartPlots\SimpleXYPlot.png
call unlink ./ChartPlots/SimpleXYPlot.png
call cscript //nologo build/plot-simple-xy-bundle.vbs /workbook:workbooks\SimpleXYPlot.xlsm /destination:ChartPlots /workbook:dummy /debug:true /data:A,B,1,1,2,4,3,9,4,0
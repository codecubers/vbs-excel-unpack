' Dim dutil, d, col, wbFile
' set dutil = new DictUtil
' set d = argsDict

'Sort by keys (named arguments or index)
' call dutil.SortDictionary(d, 1)
' EchoX "Parameter Keys: %x", join(d.Keys, "|| ")
' EchoX "Parameter Items: %x", join(d.Items, "|| ")

' set col = new Collection
' set col.Obj = d
Dim wbFile, sourceDir, destDir, data
If Wscript.Arguments.Named.Exists("workbook") Then
    wbFile = Wscript.Arguments.Named("workbook")
    EchoX "Excel workbook to be packed/unpacked: %x", wbFile
Else
    Echo "No excel workbook supplied as a parameter. Nothing to unpack."
    WScript.Quit
End If

If Wscript.Arguments.Named.Exists("source") Then
    sourceDir = Wscript.Arguments.Named("source")
    EchoX "Excel workbook will be packed from directory: %x", sourceDir
End If

If Wscript.Arguments.Named.Exists("destination") Then
    destDir = Wscript.Arguments.Named("destination")
    EchoX "Excel workbook will be unpacked to directory: %x", destDir
End If

If Wscript.Arguments.Named.Exists("data") Then
    data = Wscript.Arguments.Named("data")
    EchoX "Data received: %x", data
End If


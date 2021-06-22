' Dim dutil, d, col, wbFile
' set dutil = new DictUtil
' set d = argsDict

'Sort by keys (named arguments or index)
' call dutil.SortDictionary(d, 1)
' EchoX "Parameter Keys: %x", join(d.Keys, "|| ")
' EchoX "Parameter Items: %x", join(d.Items, "|| ")

' set col = new Collection
' set col.Obj = d
Dim wbFile
If Wscript.Arguments.Named.Exists("workbook") Then
    wbFile = Wscript.Arguments.Named("workbook")
    EchoX "Excel workbook to be unpacked: %x", wbFile
Else
    Echo "No excel workbook supplied as a parameter. Nothing to unpack."
    WScript.Quit
End If
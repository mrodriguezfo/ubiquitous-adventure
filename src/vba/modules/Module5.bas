Attribute VB_Name = "Module5"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Sheets("Margen int").Select
    Range("O393").Select
    Selection.End(xlUp).Select
    Range("RetornoPasivo[[#Headers],[Costo/activo]]").Select
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlToRight)).Select
End Sub
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro4 Macro
'

'
    Sheets("Margen int").Select
    Range("RetornoPasivo[[#Headers],[Costo/activo]]").Select
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFill Destination:=Range("O413:S416")
    Range("O413:S416").Select
End Sub
Sub Macro5()
Attribute Macro5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro5 Macro
'

'
    LastRow = Range("A1048576").End(xlUp).Row
    Range("O2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFill Destination:=Range("O2:S" & LastRow)
    
End Sub

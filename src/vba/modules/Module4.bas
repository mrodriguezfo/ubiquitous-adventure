Attribute VB_Name = "Module4"
Sub Actualizar_query_MDF()
Attribute Actualizar_query_MDF.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro5 Macro
'
Application.ScreenUpdating = False


    Sheets("VaR y dur").Range("D2").ListObject.QueryTable.Refresh BackgroundQuery:=False

    Sheets("VaR y dur").Range("L2").ListObject.QueryTable.Refresh BackgroundQuery:=False

ThisWorkbook.Sheets("Informe").Activate

End Sub

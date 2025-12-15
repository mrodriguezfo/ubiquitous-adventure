Attribute VB_Name = "Módulo1"
Sub savedata()

    ' Definir variables
    Dim retornos, datos As Workbook
    Dim hoja As Variant
    Dim rango1, rango2, rango3 As Range
    Dim iniciomes, Reporte As Long
    Dim lastsheet, lastfortable, start, lastoftable, columna, fila As Long
    Dim hojas As Variant
    Dim rango As String
    Dim hojadatos As Worksheet

    'Evitar el salto entre hojas mientras corre la macro
    Application.ScreenUpdating = False

    'Poner como valores datos del día anterior
    Sheets("Informe").Select
    Range("H10:I15").Select
    Selection.Copy
    Range("P10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

    ThisWorkbook.Sheets("Margen int").Select
    Range("A1").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    ThisWorkbook.Sheets("Margen int").Range("X1").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    
    LastRow = Range("A1048576").End(xlUp).Row
    Range("O2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFill Destination:=Range("O2:S" & LastRow)

    
    ThisWorkbook.Sheets("Portafolios activos").Select
    Range("A1").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    ThisWorkbook.Sheets("Portafolios activos").Range("V1").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    
    ThisWorkbook.Sheets("Revision retornos diarios").Select
    Range("A1").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    
    ThisWorkbook.Sheets("utilidades").Select
    Range("A1").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False


'''''' Macro vieja que actualizaba los datos ''''''
'
'    'Hoja donde se pegan los datos
'    Set hojadatos = ThisWorkbook.Sheets("Datos")
'
'    'Hojas de donde se copian los datos
'    hojas = Array("INV", "LIQ", "FISCHER", "GOLDMAN", "INTER-ACT", "INTER-PAS", "AGREGADO", "TOTAL", "FOCA", "CRED-OPERATIVO", "TOTAL+CREDITOS")
'
'    'Procesamiento
'
'    '1 ) Eliminar la tabla.
'        If hojadatos.ListObjects.Count > 0 Then
'           hojadatos.ListObjects("Datos").Unlist
'        End If
'
'        hojadatos.Rows(2 & ":" & hojadatos.Rows.Count).Delete  ' Borrar los datos del día anterior.
'
'    ' 2) Loop para extraer los datos de cada hoja.
'
'    For Each hoja In hojas
'        With ThisWorkbook.Sheets(hoja)
'            .Activate
'            lastsheet = .Cells(Rows.Count, 1).End(xlUp).Row
'            lastfortable = .Cells(lastsheet - 2, 1).End(xlUp).Row
'            start = .Cells(lastfortable, 1).End(xlUp).Row
'
'            Set rango1 = .Range("A7:H7")     ' Debe mantenerse el formato de Abacus para que funcione la macro.
'            Set rango2 = .Range("A" & lastsheet & ":H" & lastsheet)
'            rango = "A" & start & ":H" & lastfortable
'            Set rango3 = .Range(rango)
'
'            lastoftable = hojadatos.Cells(Rows.Count, 1).End(xlUp).Row
'            rango = "A" & lastoftable + 1 & ":H" & lastoftable + 1
'            rango1.Copy hojadatos.Range(rango)
'
'            lastoftable = hojadatos.Cells(Rows.Count, 1).End(xlUp).Row
'            rango = "A" & lastoftable + 1 & ":H" & lastoftable + 1
'            rango3.Copy hojadatos.Range(rango)
'
'            lastoftable = hojadatos.Cells(Rows.Count, 1).End(xlUp).Row
'            rango = "A" & lastoftable + 1 & ":H" & lastoftable + 1
'            rango2.Copy hojadatos.Range(rango)
'        End With
'    Next hoja
'    'End

    Application.ScreenUpdating = True
'    ThisWorkbook.Sheets("Macro").Activate

End Sub

Sub GenerarReporte()

    Dim wb_df As Workbook
    Dim hojainforme, hojacorreos As Worksheet
    Dim rango1, rango2 As Range
    'Dim rng As Range
    Dim rutareporte, rutacopia, rutadf As Variant
    Dim TempFilePath, fecha_reporte As String

    fecha_reporte = Format(Range("al").Value, "mm-dd-yyyy")

    Set hojainforme = ThisWorkbook.Sheets("Informe")
    Set hojacorreos = ThisWorkbook.Sheets("Correo")

    Set rango1 = hojainforme.Range("A1:M93")
    Set rango2 = hojainforme.Range("A94:M199")
    rutareporte = "S:\InfoCore\Aplicaciones\Modelos Información\Retornos\Retorno de Portafolios.pdf"
    rutacopia = "S:\InfoCore\Aplicaciones\Modelos Información\Retornos\Retorno de Portafolios " & _
                  fecha_reporte & ".pdf"
    rutadf = "S:\InfoCore\Aplicaciones\Paginas web\Otros Reportes\Diario - Retorno Portafolios Consulta.xls"


    ' Desactivar saltos
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    ' Crear pdf
    hojainforme.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
               rutareporte, Quality:=xlQualityStandard, IncludeDocProperties:=True, _
               IgnorePrintAreas:=False, OpenAfterPublish:=False

    ' Crear copia pdf
    hojainforme.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
                rutacopia, Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                IgnorePrintAreas:=False, OpenAfterPublish:=False

'    ' Crear copia del archivo de Excel para la DF.
'
'    ThisWorkbook.Sheets(Array("Informe", "Mov Ajuste Saldos Diarios", "Valor Mercado Diario VaR Serie", "Retorno Hist.", "ÍNDICES", "Gráficos", "Indice sub portafolio Inversion")).Copy
'
'    With ActiveWorkbook
'        .SaveAs Filename:= _
'            rutadf
'        .Close
'    End With
'
'    ThisWorkbook.SaveCopyAs rutadf



    'Crear imagen
    hojainforme.Activate
    Application.CutCopyMode = False
    rango1.CopyPicture
    With ActiveSheet.ChartObjects.Add(Left:=rango1.Left, Top:=rango1.Top, Width:=rango1.Width, Height:=rango1.Height)
        .ShapeRange.Line.Visible = msoFalse
        .Name = "Hoja1"
        .Activate
    End With
    ActiveChart.Paste
    ActiveChart.Export "S:\InfoCore\Aplicaciones\Modelos Información\Retornos\Adjuntos\Hoja1.jpg"
    ActiveSheet.ChartObjects("Hoja1").Delete

    'Crear imagen
    hojainforme.Activate
    rango2.CopyPicture
    With ActiveSheet.ChartObjects.Add(Left:=rango2.Left, Top:=rango2.Top, Width:=rango2.Width, Height:=rango2.Height)
        .ShapeRange.Line.Visible = msoFalse
        .Name = "Hoja2"
        .Activate
    End With
    ActiveChart.Paste (7)
    ActiveChart.Export "S:\InfoCore\Aplicaciones\Modelos Información\Retornos\Adjuntos\Hoja2.jpg"
    ActiveSheet.ChartObjects("Hoja2").Delete


    ' Crear correo
    Dim msg, firma, financiera, riesgos, contabilidad, estudios, signature, tiempo As String
    Dim hora As Integer
    Dim OL As Object '*rng As Range*'
    Dim EmailItem As Object

    msg = Range("Mensaje").Value
    firma = Environ("appdata") & "\Microsoft\Signatures\DR.htm"
    rutareporte = "S:\InfoCore\Aplicaciones\Modelos Información\Retornos\Retorno de portafolios.pdf"


    Set OL = CreateObject("Outlook.Application")
    Set EmailItem = OL.CreateItem(0)
    EmailItem.Display

'    With EmailItem
'            .BodyFormat = 2
'            .Subject = "Reporte Diario - Para revisión"
'            .Recipients.Add ("ygomez@flar.net")
'            .Attachments.Add (rutareporte)
'            .HTMLBody = "<p><img src = 'S:\InfoCore\Aplicaciones\Modelos Información\Retornos\Adjuntos\Hoja1.jpg' ></p>" & _
'                        "<p><img src = 'S:\InfoCore\Aplicaciones\Modelos Información\Retornos\Adjuntos\Hoja2.jpg'> </p>"
'    End With

    ' Direccionar la firma

    If Dir(firma) <> "" Then
        signature = GetBoiler(firma)
    Else
        signature = ""
    End If

    On Error Resume Next

    ' Nombres de las listas de distribución de correos electrónicos
    financiera = "Direccion_Financiera"
    riesgos = "Direccion_Riesgos"
    contabilidad = "Contabilidad"
    'estudios = "Direccion_Estudios_Economicos_FLAR"


    'Hora del día

    hora = Hour(Now)

    If hora < 12 Then
        tiempo = "Buenos días,"
        Else
        tiempo = "Buenas tardes,"
    End If


   With EmailItem
            .BodyFormat = 2
            .Subject = "Informe diario de valor de mercado y retorno al " & Range("al").text
            .Recipients.Add (Range("Destinatarios").Value)
            .Recipients.Add (financiera)
            .Recipients.Add (riesgos)
            .Recipients.Add (contabilidad)
            '.Recipients.Add (estudios)
            .Attachments.Add (rutareporte)
            .HTMLBody = "<BODY style=font-size:11pt;font-family:Verdana;color:rgb(38,58,144);line-height:2><p>" & tiempo & _
                        "<BODY style=font-size:11pt;font-family:Verdana;color:rgb(38,58,144);line-height:1><p>" & msg & "</p>" & _
                        "<style=line-height:2><p><img src = 'S:\InfoCore\Aplicaciones\Modelos Información\Retornos\Adjuntos\Hoja1.jpg' <br/> </p>" & _
                        "<style=line-height:2><p><img src = 'S:\InfoCore\Aplicaciones\Modelos Información\Retornos\Adjuntos\Hoja2.jpg'> </p>" & _
                        "<p>" & signature & "</p></BODY>"

    End With

    ' Desactivar saltos
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True


    'Call cargue_bmk_diario
    Call cargue_bmk_diario_nuevo

End Sub

Sub inicio_mes()

    Dim hoja_retornos, hoja_saldos, hoja_reporte As Worksheet
    Dim tabla_saldos As ListObject
    Dim rango1, rango2 As Range
    'Dim rng As Range
    Dim last_row_1, last_row_2 As Double
    Dim rangos_i, rangos_f, rangos_trim_i, rangos_trim_f As Collection


    ' Set collection
    Set rangos_i = New Collection
    Set rangos_f = New Collection
    Set rangos_trim_i = New Collection
    Set rangos_trim_f = New Collection

    'Inversiones
    rangos_i.Add "A"          'fecha
    rangos_f.Add "C"          'retorno bmk
    rangos_trim_i.Add "E"     'retorno trimestral
    rangos_trim_f.Add "F"     'retorno trimestral bmk

    'Liquidez
    rangos_i.Add "Y"          'retorno mtd
    rangos_f.Add "Z"          'retorno mtd bmk
    rangos_trim_i.Add "AB"     'retorno trimestral
    rangos_trim_f.Add "AC"     'retorno trimestral bmk

    'FOCA
    rangos_i.Add "AQ"
    rangos_f.Add "AR"
    rangos_trim_i.Add "AT"
    rangos_trim_f.Add "AU"


    'Agregado
    rangos_i.Add "CP"
    rangos_f.Add "CQ"
    rangos_trim_i.Add "CS"
    rangos_trim_f.Add "CT"

    'BNP Paribas
    rangos_i.Add "GR"
    rangos_f.Add "GS"
    rangos_trim_i.Add "GU"
    rangos_trim_f.Add "GV"


    'Goldman Sachs
    rangos_i.Add "HQ"
    rangos_f.Add "HR"
    rangos_trim_i.Add "HT"
    rangos_trim_f.Add "HU"

    'Total
    rangos_i.Add "IN"
    rangos_f.Add "IO"
    rangos_trim_i.Add "IQ"
    rangos_trim_f.Add "IR"

    'CAP
    rangos_i.Add "JF"
    rangos_f.Add "JG"
    rangos_trim_i.Add "JI"
    rangos_trim_f.Add "JJ"

    'Portafolio de operaciones
    rangos_i.Add "JS"
    rangos_f.Add "JS"
    rangos_trim_i.Add "JT"
    rangos_trim_f.Add "JT"

    'Portafolio de patrimonio
    rangos_i.Add "JZ"
    rangos_f.Add "JZ"
    rangos_trim_i.Add "KA"
    rangos_trim_f.Add "KA"


    ' Definir hojas a modificar
    Set hoja_retornos = ThisWorkbook.Sheets("Retorno Hist.")
    Set hoja_saldos = ThisWorkbook.Sheets("Mov ajuste saldos diarios")
    Set hoja_reporte = ThisWorkbook.Sheets("Informe")

    ' Actualizar campos de retornos.
    With hoja_retornos
        .Activate
        last_row_1 = .Cells(.Rows.Count, "A").End(xlUp).Row
        last_row_2 = last_row_1 - 1
        
        
        last_row_3 = last_row_1 + 1
        ActiveSheet.ListObjects("Inversiones").Resize Range("$A$4:$U$" & last_row_3)
        ActiveSheet.ListObjects("Liquidez").Resize Range("$X$4:$AM$" & last_row_3)
        ActiveSheet.ListObjects("AGREGADO").Resize Range("$CO$4:$DG$" & last_row_3)
        ActiveSheet.ListObjects("BNP").Resize Range("$GQ$4:$HK$" & last_row_3)
        ActiveSheet.ListObjects("GS").Resize Range("$HP$4:$IJ$" & last_row_3)
        ActiveSheet.ListObjects("TOTAL").Resize Range("$IM$4:$JB$" & last_row_3)
        ActiveSheet.ListObjects("CAP").Resize Range("$JE$4:$JP$" & last_row_3)
        ActiveSheet.ListObjects("OPERACIONES").Resize Range("$JR$4:$JW$" & last_row_3)
        ActiveSheet.ListObjects("PATROMONIO").Resize Range("$JY$4:$KC$" & last_row_3)

        MsgBox (last_row_1)
        Rows(last_row_1).Select
        Selection.AutoFill Destination:=Rows(last_row_1 & ":" & (last_row_1 + 1))

        For i = 1 To rangos_i.Count
        ' Valor tabla inversiones
            Set rango1 = .Range(rangos_i(i) & last_row_1 & ":" & rangos_f(i) & last_row_1)
            rango1.Copy
            rango1.PasteSpecial Paste:=xlPasteValuesAndNumberFormats

            Set rango2 = .Range(rangos_trim_i(i) & last_row_2 & ":" & rangos_trim_f(i) & last_row_2)
            rango2.AutoFill Destination:=.Range(rango2, rango2.Offset(1, 0))
        Next i
    End With


    ' Actualizar referencias

   ' With hoja_reporte
    '    .Activate
    '    .Cells.Replace What:=last_row_1, Replacement:=(last_row_1 + 1), LookIn:=xlFormulas, LookAt:=xlPart, _
    '    SearchOrder:=xlByColumns, MatchCase:=False
   ' End With

    'Call load
    Call load_nuevo


End Sub

Function GetBoiler(ByVal sFile As String) As String
'Dick Kusleika
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(sFile).OpenAsTextStream(1, -2)
    GetBoiler = ts.readall
    ts.Close
End Function

Sub load()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sConnString As String
    Dim test As String
    Dim hoja_retornos As Worksheet
 
    ' Create the connection string.
    sConnString = "Provider=SQLOLEDB;Data Source=FLAR-TSQL2017EX\SQL2017EX;" & _
                  "Initial Catalog=RiesgoDB;" & _
                  "Integrated Security=SSPI;"
    
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    conn.Open sConnString
    
    Set hoja_retornos = ThisWorkbook.Sheets("Retorno Hist.")
    
    hoja_retornos.Activate
    
    UltLinea = Range("A1048576").End(xlUp).Row
    
    
        
        'Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Range("A" & i).Value & "','" & Range("B" & i).Value & "','" & Range("C" & i).Value & "','" & Range("D" & i).Value & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea)) & "-" & Month(Range("A" & UltLinea)) & "-" & Day(Range("A" & UltLinea)) & "','INVECOMPOSITE BMK','TWRRM-T-D BMK','" & Range("C" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea)) & "-" & Month(Range("A" & UltLinea)) & "-" & Day(Range("A" & UltLinea)) & "','INVECOMPOSITE BMK','TWRR3month BMK','" & Range("F" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea)) & "-" & Month(Range("A" & UltLinea)) & "-" & Day(Range("A" & UltLinea)) & "','INVECOMPOSITE BMK','TWRRY-T-D BMK','" & Range("I" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea)) & "-" & Month(Range("A" & UltLinea)) & "-" & Day(Range("A" & UltLinea)) & "','INVECOMPOSITE BMK','TWRRIncept.Ann. BMK','" & Range("L" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea)) & "-" & Month(Range("A" & UltLinea)) & "-" & Day(Range("A" & UltLinea)) & "','INVECOMPOSITE BMK','TWRR1yr. BMK','" & Range("Q" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea)) & "-" & Month(Range("A" & UltLinea)) & "-" & Day(Range("A" & UltLinea)) & "','INVECOMPOSITE BMK','TWRR3yr.Ann. BMK','" & Range("T" & UltLinea - 1).Value & "';")
    
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea)) & "-" & Month(Range("A" & UltLinea)) & "-" & Day(Range("A" & UltLinea)) & "','TOTAL BMK','TWRRM-T-D BMK','" & Range("CQ" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea)) & "-" & Month(Range("A" & UltLinea)) & "-" & Day(Range("A" & UltLinea)) & "','TOTAL BMK','TWRR3month BMK','" & Range("CT" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea)) & "-" & Month(Range("A" & UltLinea)) & "-" & Day(Range("A" & UltLinea)) & "','TOTAL BMK','TWRRY-T-D BMK','" & Range("CW" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea)) & "-" & Month(Range("A" & UltLinea)) & "-" & Day(Range("A" & UltLinea)) & "','TOTAL BMK','TWRRIncept.Ann. BMK','" & Range("CZ" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea)) & "-" & Month(Range("A" & UltLinea)) & "-" & Day(Range("A" & UltLinea)) & "','TOTAL BMK','TWRR1yr. BMK','" & Range("DC" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea)) & "-" & Month(Range("A" & UltLinea)) & "-" & Day(Range("A" & UltLinea)) & "','TOTAL BMK','TWRR3yr.Ann. BMK','" & Range("DF" & UltLinea - 1).Value & "';")
    
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea)) & "-" & Month(Range("A" & UltLinea)) & "-" & Day(Range("A" & UltLinea)) & "','ADMINISTRADORES BMK','TWRRM-T-D BMK','" & Range("HR" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea)) & "-" & Month(Range("A" & UltLinea)) & "-" & Day(Range("A" & UltLinea)) & "','ADMINISTRADORES BMK','TWRR3month BMK','" & Range("HU" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea)) & "-" & Month(Range("A" & UltLinea)) & "-" & Day(Range("A" & UltLinea)) & "','ADMINISTRADORES BMK','TWRRY-T-D BMK','" & Range("HX" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea)) & "-" & Month(Range("A" & UltLinea)) & "-" & Day(Range("A" & UltLinea)) & "','ADMINISTRADORES BMK','TWRRIncept.Ann. BMK','" & Range("ID" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea)) & "-" & Month(Range("A" & UltLinea)) & "-" & Day(Range("A" & UltLinea)) & "','ADMINISTRADORES BMK','TWRR1yr. BMK','" & Range("IA" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea)) & "-" & Month(Range("A" & UltLinea)) & "-" & Day(Range("A" & UltLinea)) & "','ADMINISTRADORES BMK','TWRR3yr.Ann. BMK','" & Range("II" & UltLinea - 1).Value & "';")
    
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea)) & "-" & Month(Range("A" & UltLinea)) & "-" & Day(Range("A" & UltLinea)) & "','AGREGADO BMK','TWRRM-T-D BMK','" & Range("JG" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea)) & "-" & Month(Range("A" & UltLinea)) & "-" & Day(Range("A" & UltLinea)) & "','AGREGADO BMK','TWRR3month BMK','" & Range("JJ" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea)) & "-" & Month(Range("A" & UltLinea)) & "-" & Day(Range("A" & UltLinea)) & "','AGREGADO BMK','TWRRY-T-D BMK','" & Range("JM" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea)) & "-" & Month(Range("A" & UltLinea)) & "-" & Day(Range("A" & UltLinea)) & "','AGREGADO BMK','TWRR1yr. BMK','" & Range("JO" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea)) & "-" & Month(Range("A" & UltLinea)) & "-" & Day(Range("A" & UltLinea)) & "','AGREGADO BMK','TWRR3yr.Ann. BMK','" & Range("JP" & UltLinea - 1).Value & "';")
        
    conn.Close
End Sub




Sub load_nuevo()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sConnString As String
    Dim test As String
    Dim hoja_retornos As Worksheet
 
    ' Create the connection string.
    sConnString = "Provider=SQLOLEDB;Data Source=FLAR-PSQL2017\PSQL2017;" & _
                  "Initial Catalog=RiesgoDB;" & _
                  "User ID=sql_ops;Password=e^!n3+@eu?y0Cz5:"
    
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    conn.Open sConnString
    
    Set hoja_retornos = ThisWorkbook.Sheets("Retorno Hist.")
    
    hoja_retornos.Activate
    
    UltLinea = Range("A1048576").End(xlUp).Row
    
    
        
        'Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Range("A" & i).Value & "','" & Range("B" & i).Value & "','" & Range("C" & i).Value & "','" & Range("D" & i).Value & "';")
        'Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - 1)) & "-" & Month(Range("A" & UltLinea - 1)) & "-" & Day(Range("A" & UltLinea - 1)) & "','INVECOMPOSITE BMK','TWRRM-T-D BMK','" & Range("C" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - 1)) & "-" & Month(Range("A" & UltLinea - 1)) & "-" & Day(Range("A" & UltLinea - 1)) & "','INVECOMPOSITE BMK','TWRR3month BMK','" & Range("F" & UltLinea - 1).Value & "';")
        'Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - 1)) & "-" & Month(Range("A" & UltLinea - 1)) & "-" & Day(Range("A" & UltLinea - 1)) & "','INVECOMPOSITE BMK','TWRRY-T-D BMK','" & Range("I" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - 1)) & "-" & Month(Range("A" & UltLinea - 1)) & "-" & Day(Range("A" & UltLinea - 1)) & "','INVECOMPOSITE BMK','TWRRIncept.Ann. BMK','" & Range("L" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - 1)) & "-" & Month(Range("A" & UltLinea - 1)) & "-" & Day(Range("A" & UltLinea - 1)) & "','INVECOMPOSITE BMK','TWRR1yr. BMK','" & Range("Q" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - 1)) & "-" & Month(Range("A" & UltLinea - 1)) & "-" & Day(Range("A" & UltLinea - 1)) & "','INVECOMPOSITE BMK','TWRR3yr.Ann. BMK','" & Range("T" & UltLinea - 1).Value & "';")
    
        'Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - 1)) & "-" & Month(Range("A" & UltLinea - 1)) & "-" & Day(Range("A" & UltLinea - 1)) & "','TOTAL BMK','TWRRM-T-D BMK','" & Range("CQ" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - 1)) & "-" & Month(Range("A" & UltLinea - 1)) & "-" & Day(Range("A" & UltLinea - 1)) & "','TOTAL BMK','TWRR3month BMK','" & Range("CT" & UltLinea - 1).Value & "';")
        'Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - 1)) & "-" & Month(Range("A" & UltLinea - 1)) & "-" & Day(Range("A" & UltLinea - 1)) & "','TOTAL BMK','TWRRY-T-D BMK','" & Range("CW" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - 1)) & "-" & Month(Range("A" & UltLinea - 1)) & "-" & Day(Range("A" & UltLinea - 1)) & "','TOTAL BMK','TWRRIncept.Ann. BMK','" & Range("CZ" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - 1)) & "-" & Month(Range("A" & UltLinea - 1)) & "-" & Day(Range("A" & UltLinea - 1)) & "','TOTAL BMK','TWRR1yr. BMK','" & Range("DC" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - 1)) & "-" & Month(Range("A" & UltLinea - 1)) & "-" & Day(Range("A" & UltLinea - 1)) & "','TOTAL BMK','TWRR3yr.Ann. BMK','" & Range("DF" & UltLinea - 1).Value & "';")
    
        'Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - 1)) & "-" & Month(Range("A" & UltLinea - 1)) & "-" & Day(Range("A" & UltLinea - 1)) & "','ADMINISTRADORES BMK','TWRRM-T-D BMK','" & Range("HR" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - 1)) & "-" & Month(Range("A" & UltLinea - 1)) & "-" & Day(Range("A" & UltLinea - 1)) & "','ADMINISTRADORES BMK','TWRR3month BMK','" & Range("HU" & UltLinea - 1).Value & "';")
        'Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - 1)) & "-" & Month(Range("A" & UltLinea - 1)) & "-" & Day(Range("A" & UltLinea - 1)) & "','ADMINISTRADORES BMK','TWRRY-T-D BMK','" & Range("HX" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - 1)) & "-" & Month(Range("A" & UltLinea - 1)) & "-" & Day(Range("A" & UltLinea - 1)) & "','ADMINISTRADORES BMK','TWRRIncept.Ann. BMK','" & Range("ID" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - 1)) & "-" & Month(Range("A" & UltLinea - 1)) & "-" & Day(Range("A" & UltLinea - 1)) & "','ADMINISTRADORES BMK','TWRR1yr. BMK','" & Range("IA" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - 1)) & "-" & Month(Range("A" & UltLinea - 1)) & "-" & Day(Range("A" & UltLinea - 1)) & "','ADMINISTRADORES BMK','TWRR3yr.Ann. BMK','" & Range("II" & UltLinea - 1).Value & "';")
    
        'Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - 1)) & "-" & Month(Range("A" & UltLinea - 1)) & "-" & Day(Range("A" & UltLinea - 1)) & "','AGREGADO BMK','TWRRM-T-D BMK','" & Range("JG" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - 1)) & "-" & Month(Range("A" & UltLinea - 1)) & "-" & Day(Range("A" & UltLinea - 1)) & "','AGREGADO BMK','TWRR3month BMK','" & Range("JJ" & UltLinea - 1).Value & "';")
        'Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - 1)) & "-" & Month(Range("A" & UltLinea - 1)) & "-" & Day(Range("A" & UltLinea - 1)) & "','AGREGADO BMK','TWRRY-T-D BMK','" & Range("JM" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - 1)) & "-" & Month(Range("A" & UltLinea - 1)) & "-" & Day(Range("A" & UltLinea - 1)) & "','AGREGADO BMK','TWRR1yr. BMK','" & Range("JO" & UltLinea - 1).Value & "';")
        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - 1)) & "-" & Month(Range("A" & UltLinea - 1)) & "-" & Day(Range("A" & UltLinea - 1)) & "','AGREGADO BMK','TWRR3yr.Ann. BMK','" & Range("JP" & UltLinea - 1).Value & "';")
        
    conn.Close
End Sub


Sub cargue_bmk_diario()

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sConnString As String
    Dim test As String
    Dim hoja_retornos As Worksheet
 
    ' Create the connection string.
    sConnString = "Provider=SQLOLEDB;Data Source=FLAR-TSQL2017EX\SQL2017EX;" & _
                  "Initial Catalog=RiesgoDB;" & _
                  "Integrated Security=SSPI;"
    
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    conn.Open sConnString
    
    Set hoja_retornos = ThisWorkbook.Sheets("índices BBrg")
    
    hoja_retornos.Activate
    
    UltLinea = Range("A1048576").End(xlUp).Row

    'Debug.Print (UltLinea)


For i = 0 To 11

'        Debug.Print ("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','INVECOMPOSITE BMK','TWRRM-T-D BMK','" & Range("AL" & UltLinea - i).Value * 10000 & "';")
'        Debug.Print ("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','INVECOMPOSITE BMK','TWRRY-T-D BMK','" & Range("AU" & UltLinea - i).Value * 10000 & "';")
'        Debug.Print ("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','INVECOMPOSITE BMK','TWRR BMK','" & Range("BD" & UltLinea - i).Value * 10000 & "';")
'
'        Debug.Print ("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','ADMINSITRADORES BMK','TWRRM-T-D BMK','" & Range("AM" & UltLinea - i).Value * 10000 & "';")
'        Debug.Print ("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','ADMINSITRADORES BMK','TWRRY-T-D BMK','" & Range("AV" & UltLinea - i).Value * 10000 & "';")
'        Debug.Print ("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','ADMINSITRADORES BMK','TWRR BMK','" & Range("BE" & UltLinea - i).Value * 10000 & "';")
        
        

        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','INVECOMPOSITE BMK','TWRRM-T-D BMK','" & Range("AL" & UltLinea - i).Value * 10000 & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','INVECOMPOSITE BMK','TWRRY-T-D BMK','" & Range("AU" & UltLinea - i).Value * 10000 & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','INVECOMPOSITE BMK','TWRR BMK','" & Range("BD" & UltLinea - i).Value * 10000 & "';")

        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','ADMINISTRADORES BMK','TWRRM-T-D BMK','" & Range("AM" & UltLinea - i).Value * 10000 & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','ADMINISTRADORES BMK','TWRRY-T-D BMK','" & Range("AV" & UltLinea - i).Value * 10000 & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','ADMINISTRADORES BMK','TWRR BMK','" & Range("BE" & UltLinea - i).Value * 10000 & "';")

        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','TOTAL BMK','TWRRM-T-D BMK','" & Range("AN" & UltLinea - i).Value * 10000 & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','TOTAL BMK','TWRRY-T-D BMK','" & Range("AW" & UltLinea - i).Value * 10000 & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','TOTAL BMK','TWRR BMK','" & Range("BF" & UltLinea - i).Value * 10000 & "';")

        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','AGREGADO BMK','TWRRM-T-D BMK','" & Range("AO" & UltLinea - i).Value * 10000 & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','AGREGADO BMK','TWRRY-T-D BMK','" & Range("AX" & UltLinea - i).Value * 10000 & "';")
        Set rs = conn.Execute("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','AGREGADO BMK','TWRR BMK','" & Range("BG" & UltLinea - i).Value * 10000 & "';")

Next


    conn.Close
End Sub



Sub cargue_bmk_diario_nuevo()

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sConnString As String
    Dim test As String
    Dim hoja_retornos As Worksheet
 
    ' Create the connection string.
    sConnString = "Provider=SQLOLEDB;Data Source=FLAR-PSQL2017\PSQL2017;" & _
                  "Initial Catalog=RiesgoDB;" & _
                  "User ID=sql_ops;Password=e^!n3+@eu?y0Cz5:"
    
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    conn.Open sConnString
    
    Set hoja_retornos = ThisWorkbook.Sheets("índices BBrg")
    
    hoja_retornos.Activate
    
    UltLinea = Range("A1048576").End(xlUp).Row

    'Debug.Print (UltLinea)


For i = 0 To 11
        Debug.Print UltLinea - i
       Debug.Print ("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','INVECOMPOSITE BMK','TWRRM-T-D BMK','" & Range("AL" & UltLinea - i).Value * 10000 & "';")
        Debug.Print ("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','INVECOMPOSITE BMK','TWRRY-T-D BMK','" & Range("AU" & UltLinea - i).Value * 10000 & "';")
        Debug.Print ("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','INVECOMPOSITE BMK','TWRR BMK','" & Range("BD" & UltLinea - i).Value * 10000 & "';")

        Debug.Print ("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','ADMINSITRADORES BMK','TWRRM-T-D BMK','" & Range("AM" & UltLinea - i).Value * 10000 & "';")
        Debug.Print ("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','ADMINSITRADORES BMK','TWRRY-T-D BMK','" & Range("AV" & UltLinea - i).Value * 10000 & "';")
        Debug.Print ("Exec dbo.Insert_MDF_II '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','ADMINSITRADORES BMK','TWRR BMK','" & Range("BE" & UltLinea - i).Value * 10000 & "';")
        
        

        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','INVECOMPOSITE BMK','TWRRM-T-D BMK','" & Range("AL" & UltLinea - i).Value * 10000 & "';")
        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','INVECOMPOSITE BMK','TWRRY-T-D BMK','" & Range("AU" & UltLinea - i).Value * 10000 & "';")
        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','INVECOMPOSITE BMK','TWRR BMK','" & Range("BD" & UltLinea - i).Value * 10000 & "';")

        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','ADMINISTRADORES BMK','TWRRM-T-D BMK','" & Range("AM" & UltLinea - i).Value * 10000 & "';")
        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','ADMINISTRADORES BMK','TWRRY-T-D BMK','" & Range("AV" & UltLinea - i).Value * 10000 & "';")
        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','ADMINISTRADORES BMK','TWRR BMK','" & Range("BE" & UltLinea - i).Value * 10000 & "';")

        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','TOTAL BMK','TWRRM-T-D BMK','" & Range("AN" & UltLinea - i).Value * 10000 & "';")
        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','TOTAL BMK','TWRRY-T-D BMK','" & Range("AW" & UltLinea - i).Value * 10000 & "';")
        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','TOTAL BMK','TWRR BMK','" & Range("BF" & UltLinea - i).Value * 10000 & "';")

        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','AGREGADO BMK','TWRRM-T-D BMK','" & Range("AO" & UltLinea - i).Value * 10000 & "';")
        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','AGREGADO BMK','TWRRY-T-D BMK','" & Range("AX" & UltLinea - i).Value * 10000 & "';")
        Set rs = conn.Execute("Exec dwh.Insert_Metricas_Agregadas '" & Year(Range("A" & UltLinea - i)) & "-" & Month(Range("A" & UltLinea - i)) & "-" & Day(Range("A" & UltLinea - i)) & "','AGREGADO BMK','TWRR BMK','" & Range("BG" & UltLinea - i).Value * 10000 & "';")

Next


    conn.Close
End Sub


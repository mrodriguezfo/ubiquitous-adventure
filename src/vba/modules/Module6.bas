Attribute VB_Name = "Module6"
Option Explicit

' ========= CONFIG =========
Private Const REPO_ROOT_FOLDER_NAME As String = "repo_export"
' ==========================

Public Sub Export_Repo_Sources()
    Dim repoRoot As String
    repoRoot = GetRepoRootPath()
    
    EnsureFolder repoRoot
    EnsureFolder repoRoot & "\src"
    EnsureFolder repoRoot & "\src\vba"
    EnsureFolder repoRoot & "\src\vba\modules"
    EnsureFolder repoRoot & "\src\vba\classes"
    EnsureFolder repoRoot & "\src\vba\forms"
    EnsureFolder repoRoot & "\src\vba\document"
    
    EnsureFolder repoRoot & "\src\powerquery"
    EnsureFolder repoRoot & "\src\powerquery\queries"
    
    ' Export Power Query (M)
    ExportPowerQueries repoRoot & "\src\powerquery\queries"
    
    ' Export VBA project components
    ExportVBAComponents repoRoot
    
    ' Write a small README
    WriteReadme repoRoot
    
    MsgBox "Export complete:" & vbCrLf & repoRoot, vbInformation
End Sub

' -------------------------
' Power Query export
' -------------------------
Private Sub ExportPowerQueries(ByVal targetFolder As String)
    On Error GoTo NoPQ
    
    Dim q As Object ' WorkbookQuery
    For Each q In ThisWorkbook.Queries
        Dim filePath As String
        filePath = targetFolder & "\" & SanitizeFileName(CStr(q.Name)) & ".m"
        WriteTextFile filePath, CStr(q.Formula)
    Next q
    
    Exit Sub
    
NoPQ:
    ' If workbook has no Queries collection (older Excel) or none exist, just ignore.
    ' You can inspect Err.Number here if you want.
    Err.Clear
End Sub

' -------------------------
' VBA export
' -------------------------
Private Sub ExportVBAComponents(ByVal repoRoot As String)
    ' Requires: Trust access to the VBA project object model
    ' Excel: File -> Options -> Trust Center -> Macro Settings -> "Trust access..."
    
    On Error GoTo VBAFail
    
    Dim comp As Object ' VBIDE.VBComponent
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Dim outPath As String
        
        Select Case comp.Type
            Case 1 ' vbext_ct_StdModule
                outPath = repoRoot & "\src\vba\modules\" & SanitizeFileName(comp.Name) & ".bas"
                comp.Export outPath
                
            Case 2 ' vbext_ct_ClassModule
                outPath = repoRoot & "\src\vba\classes\" & SanitizeFileName(comp.Name) & ".cls"
                comp.Export outPath
                
            Case 3 ' vbext_ct_MSForm
                outPath = repoRoot & "\src\vba\forms\" & SanitizeFileName(comp.Name) & ".frm"
                comp.Export outPath
                
            Case 100 ' vbext_ct_Document (ThisWorkbook / Sheet modules)
                outPath = repoRoot & "\src\vba\document\" & SanitizeFileName(comp.Name) & ".cls"
                comp.Export outPath
                
            Case Else
                ' Unknown type: export as text in modules folder
                outPath = repoRoot & "\src\vba\modules\" & SanitizeFileName(comp.Name) & ".txt"
                WriteTextFile outPath, GetComponentCode(comp)
        End Select
    Next comp
    
    Exit Sub
    
VBAFail:
    MsgBox _
        "Could not export VBA. Most likely you need to enable:" & vbCrLf & _
        "Trust Center -> Macro Settings -> 'Trust access to the VBA project object model'." & vbCrLf & _
        "Error: " & Err.Description, vbExclamation
    Err.Clear
End Sub

Private Function GetComponentCode(ByVal comp As Object) As String
    On Error GoTo Fail
    GetComponentCode = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
    Exit Function
Fail:
    GetComponentCode = ""
End Function

' -------------------------
' Helpers: repo root, IO
' -------------------------
Private Function GetRepoRootPath() As String
    ' Export next to workbook in a folder named repo_export_YYYYMMDD_HHMMSS
    Dim basePath As String
    If Len(ThisWorkbook.Path) > 0 Then
        basePath = ThisWorkbook.Path
    Else
        basePath = Environ$("USERPROFILE") & "\Desktop"
    End If
    
    GetRepoRootPath = basePath & "\" & REPO_ROOT_FOLDER_NAME & "_" & Format(Now, "yyyymmdd_hhnnss")
End Function

Private Sub EnsureFolder(ByVal folderPath As String)
    If Dir(folderPath, vbDirectory) = vbNullString Then
        MkDir folderPath
    End If
End Sub

Private Sub WriteTextFile(ByVal filePath As String, ByVal text As String)
    Dim f As Integer
    f = FreeFile
    Open filePath For Output As #f
    Print #f, text
    Close #f
End Sub

Private Function SanitizeFileName(ByVal s As String) As String
    Dim badChars As Variant, c As Variant
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each c In badChars
        s = Replace(s, CStr(c), "_")
    Next c
    SanitizeFileName = Trim$(s)
End Function

Private Sub WriteReadme(ByVal repoRoot As String)
    Dim content As String
    content = ""
    content = content & "# Excel Source Export" & vbCrLf & vbCrLf
    content = content & "Generated: " & Format(Now, "yyyy-mm-dd HH:nn:ss") & vbCrLf & vbCrLf
    content = content & "## Contents" & vbCrLf
    content = content & "- src/vba: Exported VBA components (.bas/.cls/.frm)" & vbCrLf
    content = content & "- src/powerquery/queries: Exported Power Query M scripts (.m)" & vbCrLf & vbCrLf
    content = content & "## Notes" & vbCrLf
    content = content & "- To export VBA, enable: Trust Center -> 'Trust access to the VBA project object model'." & vbCrLf
    
    WriteTextFile repoRoot & "\README.md", content
End Sub



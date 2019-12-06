Attribute VB_Name = "ExportFiles"
Sub exportVBSourceCodeFiles()
    Dim expPath As String
    Dim extName As String
    
    expPath = "F:\svn\V3R19C00\Tools_BaseStationInner\VBAMacroModules\tanjun_test - base\creation\SRAN\Summary_IUB_Macro_SRAN\"
    If Dir(expPath, vbDirectory) = "" Then
        MkDir (expPath)
    End If
    If Dir(expPath) <> "" Then
        Kill (expPath & "*.*")
    End If
    
    Dim vbc
    For Each vbc In Application.VBE.ActiveVBProject.VBComponents
        'ThisWorkbook/Sheet1 Type都是100，模块是1，类是2，窗体是3
        Select Case vbc.Type
        Case 100    'class module
            extName = ".cls"
        Case 3    'form
            extName = ".frm"
        Case 1    'module
            extName = ".bas"
        Case 2
            extName = ".cls"
        End Select
        
        Dim basename As String
        basename = vbc.name & extName
        
        Application.VBE.ActiveVBProject.VBComponents(vbc.name).Export (expPath & "\" & basename)
    Next
End Sub



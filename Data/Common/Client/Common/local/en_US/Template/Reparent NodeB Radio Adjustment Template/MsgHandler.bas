Attribute VB_Name = "MsgHandler"
'MsgHandler

Dim mbCut As Boolean
Dim mrngSource As Range

'Initialise cell copy-paste
Public Sub InitCutCopyPaste()
  'Hook all the cut, copy and paste keystrokes
  Application.OnKey "^X", "DoCut"
  Application.OnKey "^x", "DoCut"
  Application.OnKey "+{DEL}", "DoCut"

  Application.OnKey "^C", "DoCopy"
  Application.OnKey "^c", "DoCopy"
  Application.OnKey "^{INSERT}", "DoCopy"

  Application.OnKey "^V", "DoPaste"
  Application.OnKey "^v", "DoPaste"
  Application.OnKey "+{INSERT}", "DoPaste"

  Application.OnKey "~", "DoPaste"

  'Switch off drag/drop
  Application.CellDragAndDrop = False
End Sub

Public Sub DoCut()
  If TypeOf Selection Is Range Then
    mbCut = True
    Set mrngSource = Selection
    Selection.Copy
  Else
    Set mrngSource = Nothing
  End If
End Sub

Public Sub DoCopy()
  If TypeOf Selection Is Range Then
    mbCut = False
    Set mrngSource = Selection
  Else
    Set mrngSource = Nothing
  End If
  
  Selection.Copy
End Sub

Public Sub DoPaste()
  On Error GoTo ErrorHandle
  
  If Application.CutCopyMode And Not mrngSource Is Nothing Then
    Selection.PasteSpecial xlValues
    If mbCut Then
      mrngSource.ClearContents
    End If
    
    Application.CutCopyMode = False
  Else
    ActiveSheet.Paste               '这句代码在从别的应用程序中paste进Excel时，或一个Excel中paste两次总出错，先将
  End If                            '异常吃掉，规避这个问题
  
ErrorExit:
  Exit Sub
  
ErrorHandle:
  'MsgBox "No contents to be pasted.", vbOKOnly, "Information"
  Resume ErrorExit
End Sub

'**********************************************************
'快捷键: Ctrl+Shift+H，第一次隐藏所有配置表，第二次取消隐藏
'**********************************************************
Sub SetTabVisibilityMacro()
Attribute SetTabVisibilityMacro.VB_ProcData.VB_Invoke_Func = "H\n14"
  If TableSht.Visible = xlSheetVeryHidden Then
    InitFieldSht.Visible = xlSheetVisible
    InitTableSht.Visible = xlSheetVisible
    TableSht.Visible = xlSheetVisible
    ValidSht.Visible = xlSheetVisible
  Else
    InitFieldSht.Visible = xlSheetVeryHidden
    InitTableSht.Visible = xlSheetVeryHidden
    TableSht.Visible = xlSheetVeryHidden
    ValidSht.Visible = xlSheetVeryHidden
  End If
End Sub

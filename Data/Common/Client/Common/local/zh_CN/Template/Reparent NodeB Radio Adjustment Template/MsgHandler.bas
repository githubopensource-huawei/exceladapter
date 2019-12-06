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
    ActiveSheet.Paste               '�������ڴӱ��Ӧ�ó�����paste��Excelʱ����һ��Excel��paste�����ܳ����Ƚ�
  End If                            '�쳣�Ե�������������
  
ErrorExit:
  Exit Sub
  
ErrorHandle:
  'MsgBox "No contents to be pasted.", vbOKOnly, "Information"
  Resume ErrorExit
End Sub

'**********************************************************
'��ݼ�: Ctrl+Shift+H����һ�������������ñ��ڶ���ȡ������
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

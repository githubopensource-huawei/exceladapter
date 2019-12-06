Attribute VB_Name = "ConfigureSmallTemplate"
Option Explicit
Const TableDefSht_BeginCol = 12
Const TableDefSht_EndCol = 13
Const TableDefSht_BeginRow = 15
Public Const ATMTrans = 1
Public Const IPTrans = 2
Public TemplateSheet As Worksheet
Public CurrentSheetName As String
Public ProcessSheetName As String
Sub ShowCfgForm()
    If ActiveSheet.Name = "Cover" Then Exit Sub
    CurrentSheetName = ActiveSheet.Name
    BuidOneInterfaceTemplate ("IUCS")
    BuidOneInterfaceTemplate ("IUPS")
    BuidOneInterfaceTemplate ("IUR")
    TemplateCMForm.Show vbModeless
    Sheets(CurrentSheetName).Select
End Sub

Private Sub BuidOneInterfaceTemplate(SheetName As String)
    ProcessSheetName = SheetName
    Set TemplateSheet = Sheets(SheetName)
    Call InitTemplateCMForm
End Sub

Public Sub InitTemplateCMForm()
    If ProcessSheetName = "IUCS" Then
        Select Case GetTransportType
            Case ATMTrans
               TemplateCMForm.IUCSATMOptionButton.Value = 1
            Case IPTrans
               TemplateCMForm.IUCSIPOptionButton.Value = 1
        End Select
        Call InitIUCSCommonMOComboBox
        Call InitIUCSATMMOComboBox
        Call InitIUCSIPMOComboBox
    End If
     
    If ProcessSheetName = "IUPS" Then
        Select Case GetTransportType
            Case ATMTrans
               TemplateCMForm.IUPSATMOptionButton.Value = 1
            Case IPTrans
               TemplateCMForm.IUPSIPOptionButton.Value = 1
        End Select
        Call InitIUPSCommonMOComboBox
        Call InitIUPSATMMOComboBox
        Call InitIUPSIPMOComboBox
    End If
    
    If ProcessSheetName = "IUR" Then
        Select Case GetTransportType
            Case ATMTrans
               TemplateCMForm.IURATMOptionButton.Value = 1
            Case IPTrans
               TemplateCMForm.IURIPOptionButton.Value = 1
        End Select
        Call InitIURCommonMOComboBox
        Call InitIURATMMOComboBox
        Call InitIURIPMOComboBox
    End If
    
End Sub

Public Sub InitIUCSCommonMOComboBox()
    Call InitComBoBoxValue(TemplateCMForm.IUCSCNDOMAINComboBox, "CNDOMAIN")
    Call InitComBoBoxValue(TemplateCMForm.IUCSN7DPCComboBox, "N7DPC")
    Call InitComBoBoxValue(TemplateCMForm.IUCSCNNODEComboBox, "CNNODE")
    Call InitComBoBoxValue(TemplateCMForm.IUCSAdjNodeComboBox, "ADJNODE")
    Call InitComBoBoxValue(TemplateCMForm.IUCSAdjMapComboBox, "ADJMAP")
End Sub

Public Sub InitIUCSATMMOComboBox()
    Call InitComBoBoxValue(TemplateCMForm.IUCSMTP3LKSComboBox, "MTP3LKS")
    Call InitComBoBoxValue(TemplateCMForm.IUCSMTP3LNKComboBox, "MTP3LNK")
    Call InitComBoBoxValue(TemplateCMForm.IUCSAAL2RTComboBox, "AAL2RT")
    Call InitComBoBoxValue(TemplateCMForm.IUCSAAL2PathComboBox, "AAL2PATH")
    Call InitComBoBoxValue(TemplateCMForm.IUCSMTP3RTComboBox, "MTP3RT")
End Sub

Public Sub InitIUCSIPMOComboBox()
    Call InitComBoBoxValue(TemplateCMForm.IUCSM3LKSComboBox, "M3LKS")
    Call InitComBoBoxValue(TemplateCMForm.IUCSM3LNKComboBox, "M3LNK")
    Call InitComBoBoxValue(TemplateCMForm.IUCSM3RTComboBox, "M3RT")
    Call InitComBoBoxValue(TemplateCMForm.IUCSIPPathComBoBox, "IPPATH")
End Sub

Public Sub InitIUPSCommonMOComboBox()
    Call InitComBoBoxValue(TemplateCMForm.IUPSCNDOMAINComboBox, "CNDOMAIN")
    Call InitComBoBoxValue(TemplateCMForm.IUPSN7DPCComboBox, "N7DPC")
    Call InitComBoBoxValue(TemplateCMForm.IUPSCNNODEComboBox, "CNNODE")
    Call InitComBoBoxValue(TemplateCMForm.IUPSADJNODEComboBox, "ADJNODE")
    Call InitComBoBoxValue(TemplateCMForm.IUPSADJMAPComboBox, "ADJMAP")
    Call InitComBoBoxValue(TemplateCMForm.IUPSIPPATHComboBox, "IPPATH")
End Sub

Public Sub InitIUPSATMMOComboBox()
    Call InitComBoBoxValue(TemplateCMForm.IUPSMTP3LKSComboBox, "MTP3LKS")
    Call InitComBoBoxValue(TemplateCMForm.IUPSMTP3LNKComboBox, "MTP3LNK")
    Call InitComBoBoxValue(TemplateCMForm.IUPSIPOAPVCComboBox, "IPOAPVC")
    Call InitComBoBoxValue(TemplateCMForm.IUPSMTP3RTComboBox, "MTP3RT")
End Sub

Public Sub InitIUPSIPMOComboBox()
    Call InitComBoBoxValue(TemplateCMForm.IUPSM3LKSComboBox, "M3LKS")
    Call InitComBoBoxValue(TemplateCMForm.IUPSM3LNKComboBox, "M3LNK")
    Call InitComBoBoxValue(TemplateCMForm.IUPSM3RTComboBox, "M3RT")
    'Call InitComBoBoxValue(TemplateCMForm.IUPSIPPATHComboBox, "IPPATH")
End Sub

Public Sub InitIURCommonMOComboBox()
    Call InitComBoBoxValue(TemplateCMForm.IURN7DPCComboBox, "N7DPC")
    Call InitComBoBoxValue(TemplateCMForm.IURNRNCComboBox, "NRNC")
    Call InitComBoBoxValue(TemplateCMForm.IURADJNODEComboBox, "ADJNODE")
    Call InitComBoBoxValue(TemplateCMForm.IURADJMAPComboBox, "ADJMAP")
End Sub

Public Sub InitIURATMMOComboBox()
    Call InitComBoBoxValue(TemplateCMForm.IURMTP3LKSComboBox, "MTP3LKS")
    Call InitComBoBoxValue(TemplateCMForm.IURMTP3LNKComboBox, "MTP3LNK")
    Call InitComBoBoxValue(TemplateCMForm.IURAAL2RTComboBox, "AAL2RT")
    Call InitComBoBoxValue(TemplateCMForm.IURAAL2PATHComboBox, "AAL2PATH")
    Call InitComBoBoxValue(TemplateCMForm.IURMTP3RTComboBox, "MTP3RT")
End Sub

Public Sub InitIURIPMOComboBox()
    Call InitComBoBoxValue(TemplateCMForm.IURM3LKSComboBox, "M3LKS")
    Call InitComBoBoxValue(TemplateCMForm.IURM3LNKComboBox, "M3LNK")
    Call InitComBoBoxValue(TemplateCMForm.IURM3RTComboBox, "M3RT")
    Call InitComBoBoxValue(TemplateCMForm.IURIPPATHComboBox, "IPPATH")
End Sub

Public Sub InitComBoBoxValue(Box As ComboBox, MOName As String)
  Dim ValueIndex As Integer
  Dim LineIndex As Integer
  Dim MOBeginRow As Integer
  Dim MOEndRow As Integer
  Dim BoxCurrentValue As String

  Box.Clear
  Call GetMOBeginAndEndRow(MOName, MOBeginRow, MOEndRow)
  For ValueIndex = 0 To MOEndRow - MOBeginRow + 1
     Box.AddItem (ValueIndex)
   Next ValueIndex
  
  BoxCurrentValue = MOEndRow - MOBeginRow + 1
  For LineIndex = MOBeginRow To MOEndRow
     If TemplateSheet.Rows(LineIndex).Hidden Then
       BoxCurrentValue = BoxCurrentValue - 1
     End If
  Next LineIndex
  Box.Text = BoxCurrentValue
End Sub

Private Function GetTransportType() As Integer
  Dim MOBeginRow As Integer
  Dim MOEndRow As Integer
  Dim TransportType As Integer
  
  '1~ATM , 2~IP
  
  Call GetMOBeginAndEndRow("M3LKS", MOBeginRow, MOEndRow)
  If Not TemplateSheet.Rows(MOBeginRow - 1).Hidden Then
      TransportType = 2
  Else
      TransportType = 1
  End If
  
  GetTransportType = TransportType
End Function


Public Sub HideLine(MOName As String, SaveCount As Integer)
  Dim MOBeginRow As Integer
  Dim MOEndRow As Integer
  Dim Index As Integer

  Call GetMOBeginAndEndRow(MOName, MOBeginRow, MOEndRow)
  Call UnprotectSheet(TemplateSheet)
   
  For Index = MOBeginRow To MOEndRow - SaveCount
     TemplateSheet.Rows(Index).Hidden = True
  Next Index
  If SaveCount = 0 Then
     TemplateSheet.Rows(MOBeginRow - 1).Hidden = True
  End If
End Sub

Public Sub HideTemplateEmptyRow()
    If ActiveSheet.Name = "Cover" Then Exit Sub
    
    Set TemplateSheet = ActiveSheet
    ProcessSheetName = ActiveSheet.Name
    
    Dim StartRow As Integer
    Dim EndRow As Integer
    StartRow = 7
    EndRow = TemplateSheet.UsedRange.Rows.Count
    
    Dim Index As Integer
    For Index = StartRow To EndRow
        If TemplateSheet.Range("A" + CStr(Index)).Cells(1, 1) <> "" Then
            'MsgBox ActiveSheet.Range("A" + CStr(Index)).Cells(1, 1), vbOKOnly
            Call HideEmptyLine(TemplateSheet.Range("A" + CStr(Index)).Cells(1, 1))
        End If
    Next Index
    
    TemplateSheet.Select
End Sub

Public Sub HideEmptyLine(MOName As String)
  Dim MOBeginRow As Integer
  Dim MOEndRow As Integer
  Dim Index As Integer
  Dim ExitstData As Boolean

  Call GetMOBeginAndEndRow(MOName, MOBeginRow, MOEndRow)
  Call UnprotectSheet(TemplateSheet)
   
  For Index = MOBeginRow To MOEndRow
    If TemplateSheet.Cells(Index, 2) = "" Then
       TemplateSheet.Rows(Index).Hidden = True
     Else
       ExitstData = True
     End If
  Next Index
  
  If Not ExitstData Then
     TemplateSheet.Rows(MOBeginRow - 1).Hidden = True
  End If
  
End Sub

'Search MO Begin and End Line Num of the Template
Private Sub GetMOBeginAndEndRow(MOName As String, ByRef MinLineNum As Integer, ByRef MaxLineNum As Integer)
  Dim MOBeginRow As Integer
  Dim TableDefSht As Worksheet
  Dim InterfaceBeginRow As Integer
  Dim InterfaceEndRow As Integer

  Application.ScreenUpdating = False
  Set TableDefSht = Sheets("TableDef")
  TableDefSht.Activate
  
  Call GetInterfaceBeginAndEndRow(InterfaceBeginRow, InterfaceEndRow)
  TableDefSht.UsedRange.Range("C" + CStr(InterfaceBeginRow) + ":C" + CStr(InterfaceEndRow)).Select
  MOBeginRow = Selection.Find(What:=MOName, After:=ActiveCell, LookAt:=xlWhole).Row
  MinLineNum = Cells(MOBeginRow, TableDefSht_BeginCol) + 1
  MaxLineNum = Cells(MOBeginRow, TableDefSht_EndCol)
  Application.ScreenUpdating = True
End Sub

Public Sub GetInterfaceBeginAndEndRow(ByRef BeginRow As Integer, ByRef EndRow As Integer)
  If ProcessSheetName = "IUCS" Then
      Call GetIUCSBeginAndEndRow(BeginRow, EndRow)
  ElseIf ProcessSheetName = "IUPS" Then
      Call GetIUPSBeginAndEndRow(BeginRow, EndRow)
  ElseIf ProcessSheetName = "IUR" Then
      Call GetIURBeginAndEndRow(BeginRow, EndRow)
  ElseIf ProcessSheetName = "COMMON" Then
      Call GetCOMMONBeginAndEndRow(BeginRow, EndRow)
  End If
End Sub

Private Sub GetIUCSBeginAndEndRow(BeginRow As Integer, EndRow As Integer)
    BeginRow = GetInterfaceBeginRow("IUCS")
    EndRow = GetInterfaceBeginRow("IUPS") - 1
End Sub

Private Sub GetIUPSBeginAndEndRow(BeginRow As Integer, EndRow As Integer)
    BeginRow = GetInterfaceBeginRow("IUPS")
    EndRow = GetInterfaceBeginRow("IUR") - 1
End Sub

Private Sub GetIURBeginAndEndRow(BeginRow As Integer, EndRow As Integer)
    BeginRow = GetInterfaceBeginRow("IUR")
    EndRow = GetInterfaceBeginRow("COMMON") - 1
End Sub

Private Sub GetCOMMONBeginAndEndRow(BeginRow As Integer, EndRow As Integer)
    BeginRow = GetInterfaceBeginRow("COMMON")
    EndRow = TblRows + TableDefSht_BeginRow
End Sub

Public Sub ResetTemplate()
  Call UnHideTemplate(TemplateSheet)
End Sub

Public Sub ToolButtoUnHideTemplate()
    If ActiveSheet.Name = "Cover" Then Exit Sub
  Call UnHideTemplate(ActiveSheet)
End Sub

Public Sub UnHideTemplate(Template As Worksheet)
   Template.Activate
   Call UnprotectSheet(Template)
   Template.Cells.Select
   Template.Cells.EntireRow.Hidden = False
   Template.Rows("1:1").Select
   Selection.EntireRow.Hidden = True
   Call ProtectSheet(Template)
End Sub

Public Sub ReProtectTemplate()
   TemplateSheet.Activate
   Call ProtectSheet(TemplateSheet)
End Sub

Public Sub CheckIfExistData(MOName As String, SaveCount As Integer, ByRef CheckReport As String)
  Dim MOBeginRow As Integer
  Dim MOEndRow As Integer
  
  Call GetMOBeginAndEndRow(MOName, MOBeginRow, MOEndRow)
  Call CheckRow(MOBeginRow, MOEndRow - MOBeginRow - SaveCount, CheckReport)
End Sub

Private Sub CheckRow(MinLineNum As Integer, HideRowCount As Integer, ByRef CheckReport As String)
   Dim Index As Integer
    
   For Index = MinLineNum To MinLineNum + HideRowCount
     If TemplateSheet.Cells(Index, 2) <> "" Then
       CheckReport = CheckReport + "     Row:" + CStr(Index) + " exist data." & vbCrLf
     End If
   Next Index
End Sub

Public Sub SetEnglishVersion()
  ThisWorkbook.Worksheets("Cover").Activate
  Call SwitchEnglishVersion
  ThisWorkbook.Worksheets("Cover").Select
End Sub

Public Sub SetChineseVersion()
  ThisWorkbook.Worksheets("Cover").Activate
  Call SwitchChineseVersion
  ThisWorkbook.Worksheets("Cover").Select
End Sub

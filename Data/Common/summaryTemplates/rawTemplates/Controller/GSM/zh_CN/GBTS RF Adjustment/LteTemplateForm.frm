VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LteTemplateForm 
   Caption         =   "Add/Delete User-defined Template"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8160
   OleObjectBlob   =   "LteTemplateForm.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "LteTemplateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








'����������Ӻ�ɾ����MappingSiteTemplate��,��MappingCellTemplate���͡�MappingRadioTemplate��ҳ��ģ��

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''                         '
'1.��Ӻ�ɾ����MappingSiteTemplate��ҳ��ģ�塣
'��*Site Type����Cabinet Type����FDD/TDD Mode����*Site Patten���к�ѡֵ�Ĵ���    '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'�����¼�,Ĭ��������ʾSite��ǩҳ
Private Sub UserForm_Activate()
    MultiPage.Font.Size = 10
    Me.Caption = getResByKey("Bar_Template")
'    If MultiPage.value = 0 Then
'        'Me.SitePattern.SetFocus
'        'SetSiteType
'    Else
    MultiPage.value = 1
    SetFddTddIdd
'    End If
    Dim func As ToolBarFunction
    Set func = New ToolBarFunction
    If func.siteAddSupport = False Then
        MultiPage.Pages.Item(0).Visible = False
    End If
    If func.cellAddSupport = False Then
        MultiPage.Pages.Item(1).Visible = False
    End If
    If func.radioAddSupport = False Then
        MultiPage.Pages.Item(2).Visible = False
    End If
End Sub
'ѡ���¼�,ѡ��ͬ�ı�ǩ(����Site,Cell��Radio)
Private Sub MultiPage_Change()
 If MultiPage.value = 0 Then
    'SetSiteType
 ElseIf MultiPage.value = 1 Then
    SetFddTddIdd
 ElseIf MultiPage.value = 2 Then
    'SetFTMode
    'SetRSA
 End If
End Sub
'ѡ���¼�,ѡ��Add ѡ��
Private Sub AddSiteTemplate_Click()
    Me.SitePattern.Visible = True
    Me.SitePatternList.Visible = False
    Me.AddSiteButton.Caption = getResByKey("Add")
End Sub
'ѡ���¼�,ѡ��Deleteѡ��
Private Sub DeleteSiteTemplate_Click()
    Me.SitePattern.Visible = False
    Me.SitePatternList.Visible = True
    Me.AddSiteButton.Caption = getResByKey("Delete")
    Set_Template_Related
End Sub
'�ύ����Add/Delete����
Private Sub AddSiteButton_Click()
    If Me.AddSiteTemplate.value = True Then
        AddSite
    Else
        DeleteSite
    End If
    Call refreshCell
End Sub
'ȡ���˴β���
Private Sub CancelSiteButton_Click()
    Unload Me
End Sub

'��DeleteSite����ť�¼�,ɾ��ģ��
Private Sub DeleteSite()
    Dim templatename As String
    Dim rowscount As Long
    Dim templateStr As String
    Dim iIndex As Long
    Dim existflg As Long
    
    templatename = SitePatternList.text
    
    '��ǰ��������
    rowscount = MappingSiteTemplate.Range("a1048576").End(xlUp).row
    
    '����flag
    existflg = 0
      
    '��Site Pattern��Ϊ��ʱ����
    If templatename = "" Then
       MsgBox templatename & getResByKey("sitePatternIsEmpty"), vbExclamation, getResByKey("Warning")
        existflg = 2
        Exit Sub
    End If
    
    '��������Ĭ�Ϻ�ѡֵ������������ֵ�ظ���ʱɾ����
    For m_rowNum = 2 To rowscount
        If MappingSiteTemplate.Cells(m_rowNum, 4) = templatename _
                And MappingSiteTemplate.Cells(m_rowNum, 3) = FDDTDDMode.value _
                And MappingSiteTemplate.Cells(m_rowNum, 2) = CabinetType.value _
                And MappingSiteTemplate.Cells(m_rowNum, 1) = SiteType.value Then
            MappingSiteTemplate.Rows(m_rowNum).Delete
            existflg = 1
        End If
    Next
    
    If existflg = 0 Then
        MsgBox templatename & getResByKey("NotExist"), vbExclamation, getResByKey("Warning")
    End If
    
    If existflg = 1 Then
        SitePatternList.text = ""
        For iIndex = 0 To SitePatternList.ListCount - 1
                    If (SitePatternList.List(iIndex, 0) = templatename) Then
                        SitePatternList.RemoveItem (iIndex)
                        Exit For
                    End If
        Next
        Load Me
                MsgBox getResByKey("is deleted"), vbInformation, getResByKey("Information")
    End If
    
End Sub

'��AddSite����ť�¼�,���ģ��
Private Sub AddSite()
    Dim templatename As String
    Dim rowscount As Long
    Dim templateStr As String
    Dim iIndex As Long
    Dim existflg As Long
    Dim lastLineofGroup As Long
    
    '�û�����ֵ
    templatename = SitePattern.text
    
    '��ǰ��������
    rowscount = MappingSiteTemplate.Range("a1048576").End(xlUp).row
    
    '�ظ�����flag
    existflg = 0
      
    '��Site Pattern��Ϊ��ʱ����
    If templatename = "" Then
        MsgBox templatename & getResByKey("sitePatternIsEmpty"), vbExclamation, getResByKey("Warning")
        existflg = 1
        Exit Sub
    End If
    
    '��������Ĭ�Ϻ�ѡֵ������������ֵ�ظ���ʱ����
    For m_rowNum = 2 To rowscount
        If MappingSiteTemplate.Cells(m_rowNum, 4) = templatename _
                And MappingSiteTemplate.Cells(m_rowNum, 3) = FDDTDDMode.value _
                And MappingSiteTemplate.Cells(m_rowNum, 2) = CabinetType.value _
                And MappingSiteTemplate.Cells(m_rowNum, 1) = SiteType.value Then
            MsgBox templatename & getResByKey("AlreadyExists"), vbExclamation, getResByKey("Warning")
            SitePattern.SetFocus
            existflg = 1
            Exit For
        End If
    Next
        
    '���ҵ�ǰ������һ��
    lastLineofGroup = rowscount + 1
    Worksheets("MappingSiteTemplate").Rows(CStr(lastLineofGroup) & ":" & CStr(lastLineofGroup)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    '���ظ�ʱ������ֵ׷�ӵ���ѡֵ�б����������ƶ�����eNodeB Transport Data��ҳ�ĵ�һ����¼
    If existflg = 0 Then
        MappingSiteTemplate.Cells(lastLineofGroup, 1).value = SiteType.value
        MappingSiteTemplate.Cells(lastLineofGroup, 2).value = CabinetType.value
        MappingSiteTemplate.Cells(lastLineofGroup, 3).value = FDDTDDMode.value
        MappingSiteTemplate.Cells(lastLineofGroup, 4).value = templatename
        Load Me
         MsgBox getResByKey("is added"), vbInformation, getResByKey("Information")
    End If
End Sub

'��Site Type��ѡ���¼�
Private Sub SiteType_Change()
    SetCabinetType (Me.SiteType.text)
    
    'DBS3900_LTE֧��FDD/TDD/FDDTDD���֣�����վ��ֻ֧��FDD
    Me.FDDTDDMode.Clear
    If Me.SiteType.text = "DBS3900_LTE" Then
        Me.FDDTDDMode.AddItem ("FDD")
        Me.FDDTDDMode.AddItem ("TDD")
        Me.FDDTDDMode.AddItem ("FDDTDD")
    Else
        Me.FDDTDDMode.AddItem ("FDD")
    End If
    
    FDDTDDMode.value = "FDD"
    Set_Template_Related
    
End Sub

'��Cabinet Type��ѡ���¼�
Private Sub CabinetType_Change()
    Set_Template_Related
End Sub

'��FDDTDD Mode��ѡ���¼�
Private Sub FDDTDDMode_Change()
    Set_Template_Related
End Sub

'�ӡ�MappingSiteTypeCabinetType��ҳ��ȡ��*Site Type���к�ѡֵ
Private Sub SetSiteType()
    Dim m_rowNum As Long
    Dim m_RowNum_Inner As Long
    Dim m_Str As String
    Dim flag As Boolean
    
    Me.SiteType.Clear
    Me.SiteType.AddItem (MappingSiteTypeCabinetType.Cells(2, 1).value)
    SiteType.value = MappingSiteTypeCabinetType.Cells(2, 1).value
    For m_rowNum = 3 To MappingSiteTypeCabinetType.Range("a1048576").End(xlUp).row
        For m_RowNum_Inner = 2 To m_rowNum - 1
            If MappingSiteTypeCabinetType.Cells(m_rowNum, 1) <> MappingSiteTypeCabinetType.Cells(m_RowNum_Inner, 1) Then
                flag = False
            Else
                flag = True
                Exit For
            End If
        Next
        If flag = False Then
            Me.SiteType.AddItem (MappingSiteTypeCabinetType.Cells(m_rowNum, 1).value)
        End If
    Next
    
End Sub

'�ӡ�MappingSiteTypeCabinetType��ҳ��ȡ��Cabinet Type���к�ѡֵ
Private Sub SetCabinetType(SiteType As String)
    Dim m_rowNum As Long
    Dim m_Str As String
    Dim flag As Boolean
    flag = True
    
    '�����Cabinet Type���ɵĺ�ѡֵ
    Me.CabinetType.Clear
    
    For m_rowNum = 2 To MappingSiteTypeCabinetType.Range("a1048576").End(xlUp).row
        If SiteType = MappingSiteTypeCabinetType.Cells(m_rowNum, 1).value Then
            Me.CabinetType.AddItem (MappingSiteTypeCabinetType.Cells(m_rowNum, 2).value)
            If flag = True Then
                CabinetType.value = MappingSiteTypeCabinetType.Cells(m_rowNum, 2).value
                flag = False
            End If
        End If
    Next
    
End Sub
'�ӡ�MappingSiteTemplate��ҳ��ȡ��*Site Patten���к�ѡֵ
Private Sub Set_Template_Related()
    Dim m_rowNum As Long
    
    '�����Cabinet Type���ɵĺ�ѡֵ
    SitePatternList.Clear
    For m_rowNum = 2 To MappingSiteTemplate.Range("a1048576").End(xlUp).row
        If SiteType.text = MappingSiteTemplate.Cells(m_rowNum, 1).value _
            And CabinetType.text = MappingSiteTemplate.Cells(m_rowNum, 2).value _
            And FDDTDDMode.text = MappingSiteTemplate.Cells(m_rowNum, 3).value Then
            SitePatternList.AddItem (MappingSiteTemplate.Cells(m_rowNum, 4).value)
        End If
    Next
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'��Ӻ�ɾ����MappingCellTemplate��ҳ��ģ��                             '
'��*Band Width����FddTddIdd����FDD/TDD Mode����*Cell Mode���к�ѡֵ�Ĵ���    '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddCellTemplate_Click()
    Me.CellPatternText.Visible = True
    Me.CellPattern.Visible = False
    Me.AddCellButton.Caption = getResByKey("Add")
End Sub
Private Sub DeleteCellTemplate_Click()
    Me.CellPattern.Visible = True
    Me.CellPatternText.Visible = False
    Me.AddCellButton.Caption = getResByKey("Delete")
    SetCellPattern
End Sub
'�ύ����Add/Delete����
Private Sub AddCellButton_Click()
    If Me.AddCellTemplate.value = True Then
        AddCell
    Else
        DeleteCell
    End If
    Call refreshCell
End Sub
'ȡ���˴β���
Private Sub CancelCellButton_Click()
    Unload Me
End Sub

'��Delete����ť�¼�,ɾ��ģ��
Private Sub DeleteCell()
    Dim templatename As String
    Dim rowscount As Long
    Dim templateStr As String
    Dim iIndex As Long
    Dim existflg As Long
    Dim MappingCellTemplate As Worksheet
    Set MappingCellTemplate = ThisWorkbook.Worksheets("MappingCellTemplate")
        
        
    '�û�����ֵ
    templatename = CellPattern.text
    
    '��ǰ��������
    rowscount = MappingCellTemplate.Range("a1048576").End(xlUp).row
    
    '����flag
    existflg = 0
      
    '��Cell Pattern��Ϊ��ʱ����
    If templatename = "" Then
        MsgBox templatename & getResByKey("cellPatternIsEmpty"), vbExclamation, getResByKey("Warning")
        existflg = 2
        Exit Sub
    End If
    
    '��������Ĭ�Ϻ�ѡֵ���ҵ���Ӧ��ʱɾ����
    For m_rowNum = 2 To rowscount
        If MappingCellTemplate.Cells(m_rowNum, 5) = templatename _
                And (MappingCellTemplate.Cells(m_rowNum, 4) = SA.value Or Me.SA.Enabled = False) _
                And MappingCellTemplate.Cells(m_rowNum, 3) = FddTddIdd.value _
                And MappingCellTemplate.Cells(m_rowNum, 2) = TxRxMode.value _
                And MappingCellTemplate.Cells(m_rowNum, 1) = BandWidth.value Then
            MappingCellTemplate.Rows(m_rowNum).Delete
            SetCellPattern
            existflg = 1
        End If
    Next
    
    If existflg = 0 Then
         MsgBox templatename & getResByKey("NotExist"), vbExclamation, getResByKey("Warning")
    End If
    
    If existflg = 1 Then
        'CellPattern.text = ""
        For iIndex = 0 To CellPattern.ListCount - 1
                    If (CellPattern.List(iIndex, 0) = templatename) Then
                        CellPattern.RemoveItem (iIndex)
                        Exit For
                    End If
        Next
        Load Me
                MsgBox getResByKey("is deleted"), vbInformation, getResByKey("Information")
    End If
End Sub

'��Add����ť�¼�,���ģ��
Private Sub AddCell()
    Dim templatename As String
    Dim rowscount As Long
    Dim templateStr As String
    Dim iIndex As Long
    Dim existflg As Long
    Dim lastLineofGroup As Long
    Dim flag4, flag3, flag2, flag1 As Boolean
    
    Dim MappingCellTemplate As Worksheet
    Set MappingCellTemplate = ThisWorkbook.Worksheets("MappingCellTemplate")
    
    flag4 = False
    flag3 = False
    flag2 = False
    falg1 = False
    
    '�û�����ֵ
    templatename = CellPatternText.text
    
    '��ǰ��������
    rowscount = MappingCellTemplate.Range("E1048576").End(xlUp).row
    
    '�ظ�����flag
    existflg = 0
      
    '��Cell Pattern��Ϊ��ʱ����
    If templatename = "" Then
        MsgBox templatename & getResByKey("cellPatternIsEmpty"), vbExclamation, getResByKey("Warning")
        existflg = 1
        Exit Sub
    End If
    
    '��������Ĭ�Ϻ�ѡֵ������������ֵ�ظ���ʱ����
    For m_rowNum = rowscount To 2 Step -1
        If MappingCellTemplate.Cells(m_rowNum, 5) = templatename _
                And (MappingCellTemplate.Cells(m_rowNum, 4) = SA.value Or Me.SA.Enabled = False) _
                And MappingCellTemplate.Cells(m_rowNum, 3) = FddTddIdd.value _
                And MappingCellTemplate.Cells(m_rowNum, 2) = TxRxMode.value _
                And MappingCellTemplate.Cells(m_rowNum, 1) = BandWidth.value Then
            MsgBox templatename & getResByKey("AlreadyExists"), vbExclamation, getResByKey("Warning")
            CellPatternText.SetFocus
            existflg = 1
            Exit For
        End If
    Next
        
    '���ҵ�ǰ������һ��
    lastLineofGroup = rowscount + 1
    Worksheets("MappingCellTemplate").Rows(CStr(lastLineofGroup) & ":" & CStr(lastLineofGroup)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        
    '���ظ�ʱ������ֵ׷�ӵ���ѡֵ�б����������ƶ�����Base Station Transport Data��ҳ�ĵ�һ����¼
    If existflg = 0 Then
        MappingCellTemplate.Cells(lastLineofGroup, 1).value = BandWidth.value
        MappingCellTemplate.Cells(lastLineofGroup, 2).value = TxRxMode.value
        MappingCellTemplate.Cells(lastLineofGroup, 3).value = FddTddIdd.value
        MappingCellTemplate.Cells(lastLineofGroup, 4).value = SA.value
        MappingCellTemplate.Cells(lastLineofGroup, 5).value = templatename
        Me.CellPatternText.value = ""
        Me.CellPatternText.SetFocus
        Load Me
        MsgBox getResByKey("is added"), vbInformation, getResByKey("Information")
    End If
End Sub

Private Sub FddTddIdd_Change()
    SetBandWidth
    SetTxRxMode
    SetSA
    SetCellPattern
End Sub
Private Sub BandWidth_Change()
    SetCellPattern
End Sub
Private Sub TxRxMode_Change()
    SetCellPattern
End Sub
Private Sub SA_Change()
    SetCellPattern
End Sub
'���á�*FddTddIdd���к�ѡֵ
Private Sub SetFddTddIdd()

    Me.FddTddIdd.Clear
    Me.FddTddIdd.AddItem ("FDD")
    Me.FddTddIdd.AddItem ("TDD")
    Me.FddTddIdd.text = "TDD"
    
End Sub

'���á�*BandWidth���к�ѡֵ
Private Sub SetBandWidth()
    
    Me.BandWidth.Clear
    
    If Me.FddTddIdd.text = "FDD" Then
        Me.BandWidth.AddItem ("1.4M")
        Me.BandWidth.AddItem ("3M")
        Me.BandWidth.AddItem ("5M")
        Me.BandWidth.AddItem ("10M")
        Me.BandWidth.AddItem ("15M")
        Me.BandWidth.AddItem ("20M")
        Me.BandWidth.text = ("1.4M")
    ElseIf Me.FddTddIdd.text = "TDD" Then
        Me.BandWidth.AddItem ("5M")
        Me.BandWidth.AddItem ("10M")
        Me.BandWidth.AddItem ("15M")
        Me.BandWidth.AddItem ("20M")
        Me.BandWidth.text = "10M"
    End If

End Sub
'���á�*TxRxMode���к�ѡֵ
Private Sub SetTxRxMode()
    
    Me.TxRxMode.Clear
    
    If Me.FddTddIdd.text = "FDD" Then
        Me.TxRxMode.AddItem ("1T1R")
        Me.TxRxMode.AddItem ("1T2R")
        Me.TxRxMode.AddItem ("2T2R")
        Me.TxRxMode.AddItem ("2T4R")
        Me.TxRxMode.AddItem ("4T4R")
        Me.TxRxMode.text = "1T1R"
    ElseIf Me.FddTddIdd.text = "TDD" Then
        Me.TxRxMode.AddItem ("1T1R")
        Me.TxRxMode.AddItem ("2T2R")
        Me.TxRxMode.AddItem ("4T4R")
        Me.TxRxMode.AddItem ("8T8R")
        Me.TxRxMode.text = "1T1R"
    End If
    
End Sub

'���á�*SA���к�ѡֵ
Private Sub SetSA()
    
    Me.SA.Clear
    If Me.FddTddIdd.text = "TDD" Then
        Me.SA.Enabled = True
        Me.SA.AddItem ("SA0")
        Me.SA.AddItem ("SA1")
        Me.SA.AddItem ("SA2")
        Me.SA.AddItem ("SA3")
        Me.SA.AddItem ("SA4")
        Me.SA.AddItem ("SA5")
        Me.SA.AddItem ("SA6")
        Me.SA.text = "SA0"
    Else
        'Me.SA.text = " "
        Me.SA.Enabled = False
    End If
    
End Sub
'�ӡ�MappingCellTemplate��ҳ��ȡ��*Cell Pattern���к�ѡֵ
Private Sub SetCellPattern()
    Dim m_rowNum As Long
    Dim m_Str As String
    Dim flag As Boolean
    flag = True
    
    Dim MappingCellTemplate As Worksheet
    Set MappingCellTemplate = ThisWorkbook.Worksheets("MappingCellTemplate")
    
    FddTddIdd = Me.FddTddIdd.text
    BandWidth = Me.BandWidth.text
    TxRxMode = Me.TxRxMode.text
    SA = Me.SA.text
    
    '�����Cell Pattern���ɵĺ�ѡֵ
    Me.CellPattern.Clear
    
    For m_rowNum = 2 To MappingCellTemplate.Range("a1048576").End(xlUp).row
         If BandWidth = MappingCellTemplate.Cells(m_rowNum, 1).value And FddTddIdd = MappingCellTemplate.Cells(m_rowNum, 3).value And TxRxMode = MappingCellTemplate.Cells(m_rowNum, 2).value Then
            If (Me.SA.Enabled = False) Or (SA = MappingCellTemplate.Cells(m_rowNum, 4).value) Then
                Me.CellPattern.AddItem (MappingCellTemplate.Cells(m_rowNum, 5).value)
                If flag = True Then
                    CellPattern.value = MappingCellTemplate.Cells(m_rowNum, 5).value
                    flag = False
                End If
             End If
        End If
    Next
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'��Ӻ�ɾ����MappingRadioTemplate��ҳ��ģ��                             '
'��FTMode����RSA����Radio Pattern���к�ѡֵ�Ĵ���    '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddRadioTemplate_Click()
    Me.RadioPatternText.Visible = True
    Me.RadioPattern.Visible = False
    Me.AddRadioButton.Caption = "Add"
End Sub
Private Sub DeleteRadioTemplate_Click()
    Me.RadioPattern.Visible = True
    Me.RadioPatternText.Visible = False
    Me.AddRadioButton.Caption = "Delete"
    SetRadioPattern
End Sub
'�ύ����Add/Delete����
Private Sub AddRadioButton_Click()
    If Me.AddRadioTemplate.value = True Then
        AddRadio
    Else
        DeleteRadio
    End If
End Sub
Private Sub CancelRadioButton_Click()
    Unload Me
End Sub
Private Sub SetFTMode()
    
    Me.FTMode.Clear
    Me.FTMode.AddItem ("FDD")
    Me.FTMode.AddItem ("TDD")
    Me.FTMode.AddItem ("FDDTDD")
    Me.FTMode.text = "TDD"

End Sub
Private Sub SetRSA()
    
    Me.RSA.Clear
    If Me.FTMode.text = "TDD" Then
        Me.RSA.Enabled = True
        Me.RSA.AddItem ("SA0")
        Me.RSA.AddItem ("SA1")
        Me.RSA.AddItem ("SA2")
        Me.RSA.AddItem ("SA3")
        Me.RSA.AddItem ("SA4")
        Me.RSA.AddItem ("SA5")
        Me.RSA.AddItem ("SA6")
        Me.RSA.text = "SA0"
    Else
        Me.RSA.text = ""
        Me.RSA.Enabled = False
    End If

End Sub
Private Sub FTMode_Change()
    SetRSA
    SetRadioPattern
End Sub

Private Sub RSA_Change()
    SetRadioPattern
End Sub
Private Sub SetRadioPattern()
    Dim m_rowNum As Long
    Dim m_Str As String
    Dim flag As Boolean
    flag = True
    
    FTMode = Me.FTMode.text
    RSA = Me.RSA.text
    
    '�����Radio Pattern���ɵĺ�ѡֵ
    Me.RadioPattern.Clear
    
    For m_rowNum = 2 To MappingRadioTemplate.Range("a1048576").End(xlUp).row
         If FTMode = MappingRadioTemplate.Cells(m_rowNum, 1).value Then
            If ((Me.RSA.Enabled = False) Or (RSA = "") Or (RSA = MappingRadioTemplate.Cells(m_rowNum, 2).value)) Then
                Me.RadioPattern.AddItem (MappingRadioTemplate.Cells(m_rowNum, 3).value)
                If flag = True Then
                    RadioPattern.value = MappingRadioTemplate.Cells(m_rowNum, 3).value
                    flag = False
                End If
             End If
        End If
    Next

End Sub
'��Delete����ť�¼�,ɾ��ģ��
Private Sub DeleteRadio()
    Dim templatename As String
    Dim rowscount As Long
    Dim templateStr As String
    Dim iIndex As Long
    Dim existflg As Long
    
    '�û�����ֵ
    templatename = RadioPattern.text
    
    '��ǰ��������
    rowscount = MappingRadioTemplate.Range("a1048576").End(xlUp).row
    
    '����flag
    existflg = 0
      
    '��Radio Pattern��Ϊ��ʱ����
    If templatename = "" Then
        MsgBox templatename & getResByKey("radioPatternIsEmpty"), vbExclamation, getResByKey("Warning")

        existflg = 2
        Exit Sub
    End If
    
    '��������Ĭ�Ϻ�ѡֵ���ҵ�ɾ�������ɾ����
    For m_rowNum = 2 To rowscount
        If MappingRadioTemplate.Cells(m_rowNum, 3) = templatename _
                And MappingRadioTemplate.Cells(m_rowNum, 1) = FTMode.value _
                And (Me.RSA.Enabled = False Or MappingRadioTemplate.Cells(m_rowNum, 2) = RSA.value) Then
            MappingRadioTemplate.Rows(m_rowNum).Delete
            existflg = 1
        End If
    Next
    
    If existflg = 0 Then
         MsgBox templatename & getResByKey("NotExist"), vbExclamation, getResByKey("Warning")
    End If
    
    If existflg = 1 Then
        RadioPattern.text = ""
        For iIndex = 0 To RadioPattern.ListCount - 1
                    If (RadioPattern.List(iIndex, 0) = templatename) Then
                        RadioPattern.RemoveItem (iIndex)
                        Exit For
                    End If
        Next
        RadioPattern.SetFocus
        Load Me
        MsgBox getResByKey("is deleted"), vbInformation, getResByKey("Information")
    End If
End Sub

'��Add����ť�¼�,���ģ��
Private Sub AddRadio()
    Dim templatename As String
    Dim rowscount As Long
    Dim templateStr As String
    Dim iIndex As Long
    Dim existflg As Long
    Dim lastLineofGroup As Long
    Dim flag As Boolean
    flag = False
    
    '�û�����ֵ
    templatename = RadioPatternText.text
    
    '��ǰ��������
    rowscount = MappingRadioTemplate.Range("a1048576").End(xlUp).row
    
    '�ظ�����flag
    existflg = 0
      
    '��Radio Pattern��Ϊ��ʱ����
    If templatename = "" Then
        MsgBox templatename & getResByKey("radioPatternIsEmpty"), vbExclamation, getResByKey("Warning")
        existflg = 1
        Exit Sub
    End If
    
    '��������Ĭ�Ϻ�ѡֵ��������������ֵ�ظ���ʱ����
    For m_rowNum = 2 To rowscount
        If MappingRadioTemplate.Cells(m_rowNum, 3) = templatename _
                And MappingRadioTemplate.Cells(m_rowNum, 2) = RSA.value _
                And MappingRadioTemplate.Cells(m_rowNum, 1) = FTMode.value Then
            MsgBox templatename & getResByKey("AlreadyExists"), vbExclamation, getResByKey("Warning")
            RadioPatternText.SetFocus
            existflg = 1
            Exit For
        End If
    Next
        
   lastLineofGroup = rowscount + 1
    Worksheets("MappingRadioTemplate").Rows(CStr(lastLineofGroup) & ":" & CStr(lastLineofGroup)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
 
    '���ظ�ʱ������ֵ׷�ӵ���ѡֵ�б����������ƶ�����eNodeB Transport Data��ҳ�ĵ�һ����¼
    If existflg = 0 Then
        MappingRadioTemplate.Cells(lastLineofGroup, 1).value = FTMode.value
        MappingRadioTemplate.Cells(lastLineofGroup, 2).value = RSA.value
        MappingRadioTemplate.Cells(lastLineofGroup, 3).value = templatename
        RadioPatternText.value = ""
        Load Me
         MsgBox getResByKey("is added"), vbInformation, getResByKey("Information")
    End If
End Sub

Private Sub refreshCell()
    Dim rangeHis As Range
    Dim row, columen As Long
    Dim selectionSheet As Worksheet
    Dim cellsheet As Worksheet
    Dim sheetName As String
    
    Set selectionSheet = ThisWorkbook.ActiveSheet
    
    For Each sheet In ThisWorkbook.Worksheets
        sheetName = sheet.name
        If isCellSheet(sheetName) Then
            Set cellsheet = ThisWorkbook.Worksheets(sheetName)
            cellsheet.Select
            Set rangeHis = Selection
            ActiveSheet.Cells(Selection.row + 1, Selection.column).Select
            rangeHis.Select
        End If
    Next sheet
    
    selectionSheet.Select
    
End Sub


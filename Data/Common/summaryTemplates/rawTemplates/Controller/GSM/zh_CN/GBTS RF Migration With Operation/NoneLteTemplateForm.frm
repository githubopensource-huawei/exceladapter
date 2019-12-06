VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NoneLteTemplateForm 
   Caption         =   "Add/Delete User-defined Template"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8160
   OleObjectBlob   =   "NoneLteTemplateForm.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "NoneLteTemplateForm"
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
    If MultiPage.value = 0 Then
        Me.SitePattern.SetFocus
        SetSiteType
    Else
        MultiPage.value = 0
    End If
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
    SetSiteType
 ElseIf MultiPage.value = 1 Then
    setCellTypePattern
 ElseIf MultiPage.value = 2 Then
    setRadioType
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
    Dim neType As String
    neType = getNeType()
    
    templatename = Trim(SitePatternList.text)
    
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
                And MappingSiteTemplate.Cells(m_rowNum, 1) = SiteType.value _
                And MappingSiteTemplate.Cells(m_rowNum, 5) = neType Then
            MappingSiteTemplate.Rows(m_rowNum).Delete
            existflg = 1
        End If
    Next
    
    If existflg = 0 Then
        MsgBox templatename & getResByKey("NotExist"), vbExclamation, getResByKey("Warning")
    End If
    
    'ɾ���ɹ�ʱ�������ƶ�����Base Station Transport Data��ҳ�ĵ�һ����¼
    If existflg = 1 Then
        SitePatternList.text = ""
        For iIndex = 0 To SitePatternList.ListCount - 1
                    If (SitePatternList.List(iIndex, 0) = templatename) Then
                        SitePatternList.RemoveItem (iIndex)
                        Exit For
                    End If
        Next
        

        'SitePatternList.SetFocus
        'ThisWorkbook.Worksheets("Base Station Transport Data").Activate
        'ThisWorkbook.Worksheets("Base Station Transport Data").range("A3").Select
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
    Dim neType As String
    
    neType = getNeType()

    '�û�����ֵ
    templatename = Trim(SitePattern.text)
    
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
                And MappingSiteTemplate.Cells(m_rowNum, 1) = SiteType.value _
                And MappingSiteTemplate.Cells(m_rowNum, 5) = neType Then
            MsgBox templatename & getResByKey("AlreadyExists"), vbExclamation, getResByKey("Warning")
            SitePattern.SetFocus
            existflg = 1
            Exit Sub
        End If
    Next
        
    '���ҵ�ǰ������һ��
    'lastLineofGroup = rowscount
    'For n_RowNum = 2 To rowscount
    '    If (MappingSiteTemplate.Cells(n_RowNum, 1) = SiteType.value And MappingSiteTemplate.Cells(n_RowNum, 5) = neType) _
    '        Or MappingSiteTemplate.Cells(n_RowNum, 1) = "" Then
    '        lastLineofGroup = n_RowNum
    '     End If
    'Next
    'lastLineofGroup = lastLineofGroup + 1
    lastLineofGroup = rowscount + 1
    Worksheets("MappingSiteTemplate").Rows(CStr(lastLineofGroup) & ":" & CStr(lastLineofGroup)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    '���ظ�ʱ������ֵ׷�ӵ���ѡֵ�б����������ƶ�����Base Station Transport Data��ҳ�ĵ�һ����¼
    If existflg = 0 Then
        MappingSiteTemplate.Cells(lastLineofGroup, 1).value = SiteType.value
        MappingSiteTemplate.Cells(lastLineofGroup, 2).value = ""
        MappingSiteTemplate.Cells(lastLineofGroup, 3).value = ""
        MappingSiteTemplate.Cells(lastLineofGroup, 4).value = templatename
        MappingSiteTemplate.Cells(lastLineofGroup, 5).value = neType
        'Worksheets("Base Station Transport Data").Activate
        'Worksheets("Base Station Transport Data").range("A3").Select
        
        SitePattern.value = ""
        SitePattern.SetFocus
        Load Me
        MsgBox getResByKey("is added"), vbInformation, getResByKey("Information")
    End If
End Sub

'��Site Type��ѡ���¼�
Private Sub SiteType_Change()
    Set_Template_Related
End Sub


'�ӡ�MappingSiteTemplate��ҳ��ȡ��*Site Type���к�ѡֵ
Private Sub SetSiteType()
    Dim m_rowNum As Long
    Dim neType As String
    neType = getNeType()
    Me.SiteType.Clear
    For m_rowNum = 2 To ProductType.Range("a1048576").End(xlUp).row
            If neType = ProductType.Cells(m_rowNum, 2).value Then
                Me.SiteType.AddItem (ProductType.Cells(m_rowNum, 1).value)
            End If
    Next
    If Me.SiteType.ListCount Then
         Me.SiteType.ListIndex = 0
    End If
End Sub


'�ӡ�MappingSiteTemplate��ҳ��ȡ��*Site Patten���к�ѡֵ
Private Sub Set_Template_Related()
    Dim m_rowNum As Long
    Dim neType As String
    neType = getNeType()
    '�����Cabinet Type���ɵĺ�ѡֵ
    SitePatternList.Clear
    For m_rowNum = 2 To MappingSiteTemplate.Range("a1048576").End(xlUp).row
        If SiteType.text = MappingSiteTemplate.Cells(m_rowNum, 1).value _
        And neType = MappingSiteTemplate.Cells(m_rowNum, 5).value Then
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
    Dim CellType, neType As String
    Dim rowscount As Long
    Dim templateStr As String
    Dim iIndex As Long
    Dim existflg As Long
    
    neType = getNeType()
    '�û�����ֵ
    templatename = Trim(CellPattern.text)
    CellType = Me.CellType.value
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
        If MappingCellTemplate.Cells(m_rowNum, 1) = templatename _
        And MappingCellTemplate.Cells(m_rowNum, 2) = CellType _
        And MappingCellTemplate.Cells(m_rowNum, 3) = neType Then
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
        'ThisWorkbook.Worksheets("Base Station Transport Data").Activate
        'ThisWorkbook.Worksheets("Base Station Transport Data").range("A3").Select
        Load Me
        MsgBox getResByKey("is deleted"), vbInformation, getResByKey("Information")
    End If
End Sub

'��Add����ť�¼�,���ģ��
Private Sub AddCell()
    Dim templatename As String
    Dim CellType, neType As String
    
    Dim rowscount As Long
    Dim templateStr As String
    Dim iIndex As Long
    Dim existflg As Long
    Dim lastLineofGroup As Long
    Dim flag4, flag3, flag2, flag1 As Boolean
    
    flag4 = False
    flag3 = False
    flag2 = False
    falg1 = False
    
    neType = getNeType()
    
    '�û�����ֵ
    templatename = Trim(CellPatternText.text)
    
    CellType = Me.CellType.value
    '��ǰ��������
    rowscount = MappingCellTemplate.Range("a1048576").End(xlUp).row
    
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
        If MappingCellTemplate.Cells(m_rowNum, 1) = templatename _
        And MappingCellTemplate.Cells(m_rowNum, 2) = CellType _
        And MappingCellTemplate.Cells(m_rowNum, 3) = neType Then
            MsgBox templatename & getResByKey("AlreadyExists"), vbExclamation, getResByKey("Warning")
            CellPatternText.SetFocus
            existflg = 1
            Exit Sub
        End If
    Next
    
    'lastLineofGroup = rowscount
    'For n_RowNum = 2 To rowscount
    '    If (MappingCellTemplate.Cells(n_RowNum, 3) = neType And MappingCellTemplate.Cells(n_RowNum, 2) = CellType) _
    '        Or MappingCellTemplate.Cells(n_RowNum, 1) = "" Then
     '       lastLineofGroup = n_RowNum
    '    End If
    'Next
    
    '���ҵ�ǰ������һ��
    'lastLineofGroup = lastLineofGroup + 1
    lastLineofGroup = rowscount + 1
    Worksheets("MappingCellTemplate").Rows(CStr(lastLineofGroup) & ":" & CStr(lastLineofGroup)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        
    '���ظ�ʱ������ֵ׷�ӵ���ѡֵ�б����������ƶ�����Base Station Transport Data��ҳ�ĵ�һ����¼
    If existflg = 0 Then
        MappingCellTemplate.Cells(lastLineofGroup, 1).value = templatename
        MappingCellTemplate.Cells(lastLineofGroup, 2).value = CellType
        MappingCellTemplate.Cells(lastLineofGroup, 3).value = neType
        Me.CellPatternText.value = ""
        Me.CellPatternText.SetFocus
        'ThisWorkbook.Worksheets("Base Station Transport Data").Activate
        'ThisWorkbook.Worksheets("Base Station Transport Data").range("A3").Select
        Load Me
        MsgBox getResByKey("is added"), vbInformation, getResByKey("Information")
    End If
End Sub

Private Sub setCellTypePattern()
        Dim neType As String
        neType = getNeType()
        CellType.Clear
        If neType = "MRAT" Then
            With CellType
                        If isContainBaseStation() Then
                            .AddItem getResByKey("GSM Local Cell")
                            .AddItem getResByKey("UMTS Local Cell")
                            .AddItem getResByKey("LTE Cell")
                        End If
                        If isContainGsmControl() Then
                            .AddItem getResByKey("GSM Logic Cell")
                        End If
                        If isContainUmtsControl() Then
                            .AddItem getResByKey("UMTS Logic Cell")
                        End If
            End With
        ElseIf neType = "UMTS" Then
            With CellType
                        If isContainBaseStation() Then
                            .AddItem getResByKey("UMTS Local Cell")
                         End If
                         If isContainUmtsControl() Then
                            .AddItem getResByKey("UMTS Logic Cell")
                         End If
            End With
        ElseIf neType = "GSM" Then
            With CellType
                         If isContainBaseStation() Then
                            .AddItem getResByKey("GSM Local Cell")
                        End If
                        If isContainGsmControl() Then
                            .AddItem getResByKey("GSM Logic Cell")
                        End If
            End With
        ElseIf neType = "LTE" Then
            With CellType
                        .AddItem getResByKey("LTE Cell")
            End With
        End If
        If Me.CellType.ListCount Then
            CellType.ListIndex = 0
        End If
End Sub

Private Sub CellType_Change()
        SetCellPattern
End Sub

'�ӡ�MappingCellTemplate��ҳ��ȡ��*Cell Pattern���к�ѡֵ
Private Sub SetCellPattern()
    Dim m_rowNum As Long
    Dim m_Str As String
    Dim flag As Boolean
    Dim CellType, neType As String
    
    neType = getNeType()
    CellType = Me.CellType.value
    
    flag = True
    
    '�����Cell Pattern���ɵĺ�ѡֵ
    Me.CellPattern.Clear
    
    For m_rowNum = 2 To MappingCellTemplate.Range("a1048576").End(xlUp).row
            If CellType = MappingCellTemplate.Cells(m_rowNum, 2).value _
            And neType = MappingCellTemplate.Cells(m_rowNum, 3).value Then
                Me.CellPattern.AddItem (MappingCellTemplate.Cells(m_rowNum, 1).value)
                 If flag = True Then
                     CellPattern.value = MappingCellTemplate.Cells(m_rowNum, 1).value
                     flag = False
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
    Me.AddRadioButton.Caption = getResByKey("Add")
End Sub
Private Sub DeleteRadioTemplate_Click()
    Me.RadioPattern.Visible = True
    Me.RadioPatternText.Visible = False
    Me.AddRadioButton.Caption = getResByKey("Delete")
    SetRadioPattern
End Sub
'�ύ����Add/Delete����
Private Sub AddRadioButton_Click()
    If Me.AddRadioTemplate.value = True Then
        AddRadio
    Else
        DeleteRadio
    End If
    Call refreshCell
End Sub
Private Sub CancelRadioButton_Click()
    Unload Me
End Sub



Private Sub setRadioType()
        Dim neType As String
        neType = getNeType()

        Me.RadioType.Clear
        If neType = "MRAT" Then
            With RadioType
                        RadioType.AddItem getResByKey("GSM Radio Template")
                        RadioType.AddItem getResByKey("UMTS Radio Template")
                        RadioType.AddItem getResByKey("LTE Radio Template")
            End With
        ElseIf neType = "UMTS" Then
            With RadioType
                        RadioType.AddItem getResByKey("UMTS Radio Template")
            End With
        ElseIf neType = "GSM" Then
            With RadioType
                        RadioType.AddItem getResByKey("GSM Radio Template")
            End With
        ElseIf neType = "LTE" Then
            With RadioType
                        RadioType.AddItem getResByKey("LTE Radio Template")
            End With
        End If
        
        If Me.RadioType.ListCount > 0 Then
            Me.RadioType.ListIndex = 0
        End If
        
End Sub

Private Sub RadioType_Change()
        SetRadioPattern
End Sub

Private Sub SetRadioPattern()
    Dim m_rowNum As Long
    Dim m_Str As String
    Dim flag As Boolean
    Dim radioTp, neType As String
    flag = True
    
    neType = getNeType()
    radioTp = Me.RadioType.value
    
    '�����Radio Pattern���ɵĺ�ѡֵ
    Me.RadioPattern.Clear
    For m_rowNum = 2 To MappingRadioTemplate.Range("a1048576").End(xlUp).row
            If MappingRadioTemplate.Cells(m_rowNum, 2).value = radioTp _
                And MappingRadioTemplate.Cells(m_rowNum, 3).value = neType Then
                    Me.RadioPattern.AddItem (MappingRadioTemplate.Cells(m_rowNum, 1).value)
                    If flag = True Then
                        RadioPattern.value = MappingRadioTemplate.Cells(m_rowNum, 1).value
                        flag = False
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
    Dim radioTp, neType As String
    
    neType = getNeType()
    
    '�û�����ֵ
    templatename = Trim(RadioPattern.text)
    radioTp = RadioType.value
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
        If MappingRadioTemplate.Cells(m_rowNum, 1) = templatename _
        And MappingRadioTemplate.Cells(m_rowNum, 2) = radioTp _
        And MappingRadioTemplate.Cells(m_rowNum, 3) = neType Then
            MappingRadioTemplate.Rows(m_rowNum).Delete
            existflg = 1
        End If
    Next
    
    If existflg = 0 Then
        MsgBox templatename & getResByKey("NotExist"), vbExclamation, getResByKey("Warning")
    End If
    
    'ɾ���ɹ�ʱ�������ƶ�����Base Station Transport Data��ҳ�ĵ�һ����¼
    If existflg = 1 Then
        'ThisWorkbook.Worksheets("Base Station Transport Data").Activate
        'ThisWorkbook.Worksheets("Base Station Transport Data").range("A3").Select
        
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
    Dim radioTp, neType As String
    radioTp = RadioType.value
    flag = False
    
    neType = getNeType()
    '�û�����ֵ
    templatename = Trim(RadioPatternText.text)
    
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
        If MappingRadioTemplate.Cells(m_rowNum, 1) = templatename _
           And MappingRadioTemplate.Cells(m_rowNum, 2) = radioTp _
           And MappingRadioTemplate.Cells(m_rowNum, 3) = neType Then
            MsgBox templatename & getResByKey("AlreadyExists"), vbExclamation, getResByKey("Warning")
            RadioPatternText.SetFocus
            existflg = 1
            Exit Sub
        End If
    Next

    'lastLineofGroup = rowscount
    'For n_RowNum = 2 To rowscount
    '    If (MappingRadioTemplate.Cells(n_RowNum, 3) = neType And MappingRadioTemplate.Cells(n_RowNum, 2) = radioTp) _
    '    Or MappingRadioTemplate.Cells(n_RowNum, 1) = "" Then
    '        lastLineofGroup = n_RowNum
    '    End If
    ' Next
    
    '���ҵ�ǰ������һ��
   ' lastLineofGroup = lastLineofGroup + 1
   lastLineofGroup = rowscount + 1
    Worksheets("MappingRadioTemplate").Rows(CStr(lastLineofGroup) & ":" & CStr(lastLineofGroup)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    '���ظ�ʱ������ֵ׷�ӵ���ѡֵ�б����������ƶ�����Base Station Transport Data��ҳ�ĵ�һ����¼
    If existflg = 0 Then
        MappingRadioTemplate.Cells(lastLineofGroup, 1).value = templatename
        MappingRadioTemplate.Cells(lastLineofGroup, 2).value = radioTp
        MappingRadioTemplate.Cells(lastLineofGroup, 3).value = neType
        RadioPatternText.value = ""
        'ThisWorkbook.Worksheets("Base Station Transport Data").Activate
        'ThisWorkbook.Worksheets("Base Station Transport Data").range("A3").Select
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













VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TemplateForm 
   Caption         =   "Add/Delete User-defined Template"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5805
   OleObjectBlob   =   "TemplateForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "TemplateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





'「Cancel」按钮事件
Private Sub CancelButton_Click()
    Unload Me
End Sub

'「Delete」按钮事件
Private Sub Delete()
    Dim TemplateName As String
    Dim rowscount As Integer
    Dim templateStr As String
    Dim iIndex As Integer
    Dim existflg As Integer
    
    '用户输入值
    TemplateName = SitePattenList.Text
    
    '当前数据行数
    rowscount = MappingSiteTemplate.Range("a65536").End(xlUp).row
    
    '存在flag
    existflg = 0
      
    '「template」为空时报错
    If Trim(TemplateName) = "" Then
        MsgBox gMsg_DelEmpty, vbExclamation, gMsg_OperWarning
        existflg = 2
        Exit Sub
    End If
    
    '遍历已有默认侯选值，存在与输入值重复项时删除。
    For m_rowNum = 2 To rowscount
        If Trim(MappingSiteTemplate.Cells(m_rowNum, 2)) = Trim(TemplateName) _
                And Trim(MappingSiteTemplate.Cells(m_rowNum, 1)) = Trim(SiteType.value) Then
            MappingSiteTemplate.Rows(m_rowNum).Delete
            Set_Template_Related
            'SitePattenList.Text = ""
            MsgBox gMsg_DelSuccH & " [" & TemplateName & "] " & gMsg_DelSuccE, vbExclamation, gMsg_OperWarning
            existflg = 1
        End If
    Next
    
    If existflg = 0 Then
        MsgBox gMsg_DelExistH & " [" & TemplateName & "] " & gMsg_DelExistE, vbExclamation, gMsg_OperWarning
    End If
    
End Sub

'「Add」按钮事件
Private Sub Add()
    Dim TemplateName As String
    Dim rowscount As Integer
    Dim templateStr As String
    Dim iIndex As Integer
    Dim existflg As Boolean
    Dim lastLineofGroup As Integer
    
    '用户输入值
    TemplateName = SitePatten.Text
    
    '当前数据行数
    rowscount = MappingSiteTemplate.Range("a65536").End(xlUp).row
    
    '重复存在flag
    existflg = False
      
    '「template」为空时报错
    If Trim(TemplateName) = "" Then
        MsgBox gMsg_AddEmpty, vbExclamation, gMsg_OperWarning
        existflg = True
        Exit Sub
    End If
    
    '遍历已有默认侯选值，存在与输入值重复项时报错。
    For m_rowNum = 2 To rowscount
        If Trim(MappingSiteTemplate.Cells(m_rowNum, 2)) = Trim(TemplateName) _
                And Trim(MappingSiteTemplate.Cells(m_rowNum, 1)) = Trim(SiteType.value) Then
            MsgBox gMsg_AddExistH & " [" & TemplateName & "] " & gMsg_AddExistE, vbExclamation, gMsg_OperWarning
            SitePatten.SetFocus
            existflg = True
            Exit Sub
        End If
    Next
        
    '查找当前组的最后一行
    lastLineofGroup = rowscount
    For n_RowNum = 2 To rowscount
        If Trim(MappingSiteTemplate.Cells(n_RowNum, 1)) = Trim(SiteType.value) Then
            lastLineofGroup = n_RowNum
        End If
    Next
    lastLineofGroup = lastLineofGroup + 1
    ThisWorkbook.Worksheets("MappingSiteTemplate").Rows(CStr(lastLineofGroup) & ":" & CStr(lastLineofGroup)).Insert shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    '不重复时将输入值追加到候选值列表
    If existflg = False Then
        MappingSiteTemplate.Cells(lastLineofGroup, 1).value = SiteType.value
        MappingSiteTemplate.Cells(lastLineofGroup, 2).value = TemplateName
        MsgBox gMsg_AddSuccH & " [" & TemplateName & "] " & gMsg_AddSuccE, vbExclamation, gMsg_OperInfo
    End If
End Sub

Private Sub OptionButton2_Click()
    TemplateForm.SitePatten.Visible = False
    TemplateForm.SitePattenList.Visible = True
    
    TemplateForm.SubmitButton.Caption = gCaption_SubmitDelete
    'TemplateForm.SitePatten.Value = ""
    Set_Template_Related
End Sub

Private Sub OptionButton1_Click()
    TemplateForm.SitePatten.Visible = True
    TemplateForm.SitePattenList.Visible = False
    
    TemplateForm.SubmitButton.Caption = gCaption_SubmitAdd
End Sub

'「Submit」按钮事件
Private Sub SubmitButton_Click()
    If TemplateForm.OptionButton1.value = True Then
        Add
    Else
        Delete
    End If
    Call refreshCell
End Sub

Private Sub refreshCell()
    Dim rangeHis As Range
    Dim row, columen As Long
    Set rangeHis = Selection
    ActiveSheet.Cells(Selection.row + 1, Selection.Column).Select
    rangeHis.Select
End Sub

'激活事件
Private Sub UserForm_Activate()
    SetSiteType
End Sub

'「BTS Type」选择事件
Private Sub SiteType_Change()
    Set_Template_Related
End Sub

'从「MappingSiteTemplate」页获取「BTS Type」列侯选值
Private Sub SetSiteType()
    Dim m_rowNum As Integer
    Dim m_RowNum_Inner As Integer
    Dim m_Str As String
    Dim flag As Boolean
    TemplateForm.SiteType.Clear
    TemplateForm.SiteType.AddItem (MappingSiteTemplate.Cells(2, 1).value)
    SiteType.value = MappingSiteTemplate.Cells(2, 1).value
    For m_rowNum = 3 To MappingSiteTemplate.Range("a65536").End(xlUp).row
        For m_RowNum_Inner = 2 To m_rowNum - 1
            If Trim(MappingSiteTemplate.Cells(m_rowNum, 1)) <> Trim(MappingSiteTemplate.Cells(m_RowNum_Inner, 1)) Then
                flag = False
            Else
                flag = True
                Exit For
            End If
        Next
        If flag = False Then
            TemplateForm.SiteType.AddItem (MappingSiteTemplate.Cells(m_rowNum, 1).value)
        End If
    Next
    
End Sub

'从「MappingSiteTemplate」页获取「BTS Template Name」列侯选值
Private Sub Set_Template_Related()
    Dim m_rowNum As Integer
    Dim m_Str As String
    
    SiteType = TemplateForm.SiteType.Text

    '旧的侯选值
    TemplateForm.SitePattenList.Clear
    
    For m_rowNum = 2 To MappingSiteTemplate.Range("a65536").End(xlUp).row
        If Trim(SiteType) = Trim(MappingSiteTemplate.Cells(m_rowNum, 1).value) And Trim(MappingSiteTemplate.Cells(m_rowNum, 2).value) <> "" Then
            TemplateForm.SitePattenList.AddItem (MappingSiteTemplate.Cells(m_rowNum, 2).value)
        End If
    Next
    
End Sub

Public Sub InitGUI()
    init ThisWorkbook
    TemplateForm.Caption = gCaption_TemplateForm
    Label2.Caption = gCaption_Label1
    Label1.Caption = gCaption_Label2
    OptionButton1.Caption = gCaption_OptionButton1
    OptionButton2.Caption = gCaption_OptionButton2
    SubmitButton.Caption = gCaption_SubmitAdd
    CancelButton.Caption = gCaption_CancelButton

End Sub











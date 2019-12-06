VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ListHyperlinksForm 
   Caption         =   "Add Ref"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9615.001
   OleObjectBlob   =   "ListHyperlinksForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "ListHyperlinksForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Dim HasHistoryData As Boolean
Private Sub AddButton_Click()
    On Error Resume Next
    '没选中单元格时应报错，选中多个单元格时只填第一个单元格
    Dim sheetName As String
    Dim GroupName As String
    Dim columnName As String
     
    sheetName = Trim(ListHyperlinksForm.SheetNameListRef.List(SheetNameListRef.ListIndex))
    GroupName = Trim(ListHyperlinksForm.ListGroupName.text)
    columnName = Trim(ListHyperlinksForm.ListColumnName.text)
    If sheetName = "" Or GroupName = "" Or columnName = "" Then
            MsgBox getResByKey("SelectValideName"), vbExclamation + vbOKCancel, getResByKey("Warning")
            Exit Sub
    End If
    Call addListRef(sheetName, GroupName, columnName)
     Unload Me
End Sub



Function checkValueIfInteger(strValue As String) As Boolean
    Dim nLoop As Long
    Dim sItem As String
    
    checkValueIfInteger = True
    For nLoop = 1 To Len(Trim(strValue))
        sItem = Right(Left(Trim(strValue), nLoop), 1)
        If sItem < "0" Or sItem > "9" Then
            checkValueIfInteger = False
            Exit Function
        End If
    Next
    
End Function

Sub setMappingDefIsRefValue(sheet As Worksheet, range As range)
    Dim m_rowNum As Long
    Dim GroupName As String
    Dim columnName As String
    Dim mappingDef As Worksheet
    Set mappingDef = Worksheets("MAPPING DEF")
    Call getGroupAndColumnName(sheet, range, GroupName, columnName)
    If GroupName <> "" And columnName <> "" Then
        For m_rowNum = 2 To mappingDef.range("a1048576").End(xlUp).row
            If sheet.name = mappingDef.Cells(m_rowNum, 1).value _
                And GroupName = mappingDef.Cells(m_rowNum, 2).value _
                And columnName = mappingDef.Cells(m_rowNum, 3).value Then
                mappingDef.Cells(m_rowNum, 6).value = "TRUE"
                Exit For
            End If
        Next
    End If
End Sub

Private Sub CancelButton_Click()
    HasHistoryData = True
    Unload Me
End Sub

Private Sub UserForm_Activate()
    ListHyperlinksForm.Caption = getResByKey("Bar_Refrence")
    If HasHistoryData = False Or ListHyperlinksForm.SheetNameListRef.ListCount < 1 Then
        SetSheetNameList
    End If
End Sub
Private Sub SetSheetNameList()
    ListHyperlinksForm.SheetNameListRef.Clear
    ListHyperlinksForm.SheetNameListRef.AddItem (getResByKey("BaseTransPort"))
    ListHyperlinksForm.SheetNameListRef.value = getResByKey("BaseTransPort")
End Sub
Private Sub SheetNameListRef_Change()
    Dim sheetName As String
    Dim ws As Worksheet
    Dim m_rowNum, m_colNum As Long
    
    GroupNameListRef.Clear
    If ListHyperlinksForm.SheetNameListRef.ListIndex <> -1 Then
        sheetName = SheetNameListRef.List(SheetNameListRef.ListIndex)
        ListSheetName.value = sheetName
        Set ws = Worksheets(sheetName)
        If sheetName = getResByKey("Comm Data") Then
            For m_rowNum = 1 To ws.range("a1048576").End(xlUp).row
                If ws.Cells(m_rowNum, 1).Interior.colorIndex = 34 Then
                    ListHyperlinksForm.GroupNameListRef.AddItem (ws.Cells(m_rowNum, 1).value)
                End If
            Next
        Else
            For m_colNum = 1 To ws.range("XFD1").End(xlToLeft).column
                If Trim(ws.Cells(1, m_colNum).value) <> "" Then
                    GroupNameListRef.AddItem (ws.Cells(1, m_colNum).value)
                End If
            Next
        End If
    End If
End Sub
Private Sub GroupNameListRef_Change()
    Dim sheetName, GroupName As String
    Dim m_rowNum, m_colNum, m_colNum1, columnsNum As Long
    ColumnNameListRef.Clear
    Dim ws As Worksheet
    
    If (ListHyperlinksForm.SheetNameListRef.ListIndex <> -1) And (GroupNameListRef.ListIndex <> -1) Then
        sheetName = SheetNameListRef.List(SheetNameListRef.ListIndex)
        GroupName = GroupNameListRef.List(GroupNameListRef.ListIndex)
        ListGroupName.value = GroupName
        Set ws = Worksheets(sheetName)
        If sheetName = getResByKey("Comm Data") Then
            For m_rowNum = 1 To ws.range("a1048576").End(xlUp).row
                If GroupName = ws.Cells(m_rowNum, 1).value Then
                    For m_colNum = 1 To ws.range("XFD" + CStr(m_rowNum + 1)).End(xlToLeft).column
                        ListHyperlinksForm.ColumnNameListRef.AddItem (ws.Cells(m_rowNum + 1, m_colNum).value)
                    Next
                    Exit For
                End If
            Next
        Else
            For m_colNum = 1 To ws.range("XFD1").End(xlToLeft).column
                If GroupName = ws.Cells(1, m_colNum).value Then
                    columnsNum = ws.Cells(1, m_colNum).MergeArea.columns.count
                    For m_colNum1 = m_colNum To m_colNum + columnsNum - 1
                        ColumnNameListRef.AddItem (ws.Cells(2, m_colNum1).value)
                    Next
                    Exit For
                End If
            Next
        End If
    End If
End Sub
Private Sub ColumnNameListRef_Change()
    Dim sheetName, GroupName As String
    Dim columnName As String
        
    If (SheetNameListRef.ListIndex <> -1) And (GroupNameListRef.ListIndex <> -1) And (ColumnNameListRef.ListIndex <> -1) Then
        sheetName = SheetNameListRef.List(SheetNameListRef.ListIndex)
        GroupName = GroupNameListRef.List(GroupNameListRef.ListIndex)
        columnName = ColumnNameListRef.List(ColumnNameListRef.ListIndex)
        ListColumnName.value = columnName
       ' If sheetName = getResByKey("Comm Data") Then
       '     LinkValueText.value = sheetName + "\" + groupName + "\" + ColumnName + "[0]"
       ' Else
       '     LinkValueText.value = sheetName + "\" + groupName + "\" + ColumnName
        'End If
    End If
End Sub
Private Sub ColumnNameListRef_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    AddButton_Click
End Sub

Private Sub setAllIubSheetCellStyle(address As String, text As String)
    Dim index As Long
    Dim count As Long
    Dim mainSheet As String
    Dim ws As Worksheet
    Dim iRange As range
    mainSheet = GetMainSheetName()
    Set ws = Worksheets(mainSheet)
    count = ws.UsedRange.rows.count

    For index = 4 To count
        If ws.Cells(index, 1).value <> "" Then
            If IsExistsSheet(ws.Cells(index, 1).value) Then
                Set iRange = Worksheets(ws.Cells(index, 1).value).range(address)
                Call setCellStyle(iRange, text)
            End If
        End If
    Next
End Sub


Private Sub setCellStyle(iRange As range, text As String)
        Call addValidation(iRange)
        With iRange.Validation
                .inputTitle = getResByKey("Reference Address")
                .inputMessage = text
                .ShowInput = True
                .ShowError = False
        End With
        With iRange.Interior
            .colorIndex = HyperLinkColorIndex
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
        End With
End Sub

Private Sub addValidation(iRange As range)
On Error Resume Next
        With iRange.Validation
                .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertInformation
        End With
End Sub


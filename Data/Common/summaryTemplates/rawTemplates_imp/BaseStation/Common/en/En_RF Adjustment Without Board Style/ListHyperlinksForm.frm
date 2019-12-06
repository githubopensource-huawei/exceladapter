VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ListHyperlinksForm 
   Caption         =   "Add Ref"
   ClientHeight    =   6750
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
    Dim groupName As String
    Dim columnName As String
    Dim rowIndexStr  As String
    Dim rowNum As Long
    Dim columnNum As Long
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim BaseStationsheet As Worksheet
    Dim target As range
    Dim maxRowNum As Long
    Dim linkCellRange As String
    Dim selectRange As range
    Set selectRange = Selection
    
    
    If Selection.Areas.count = 1 And Selection.Areas.Item(1).count = 1 And Selection.Interior.colorIndex <> SolidColorIdx And Selection.Interior.Pattern <> SolidPattern And Selection.Interior.colorIndex <> 34 And Selection.Interior.colorIndex <> 40 Then
        sheetName = Trim(ListHyperlinksForm.SheetNameListRef.List(SheetNameListRef.ListIndex))
        groupName = Trim(ListHyperlinksForm.ListGroupName.text)
        columnName = Trim(ListHyperlinksForm.ListColumnName.text)
        
        If sheetName = "" Or groupName = "" Or columnName = "" Then
            MsgBox getResByKey("SelectValideName"), vbExclamation + vbOKCancel, getResByKey("Warning")
            Exit Sub
        End If
        
        If InStr(Trim(groupName), "[") <> 0 Or InStr(Trim(columnName), "[") <> 0 Then
            MsgBox getResByKey("ReferenceInvalidCharacter"), vbExclamation + vbOKCancel, getResByKey("Warning")
            Exit Sub
        End If
        
        If Trim(LinkValueText.value) = "" Or UBound(Split(LinkValueText.value, "\")) <> 2 Then
            MsgBox getResByKey("ReferenceInvalidCharacter2"), vbExclamation + vbOKCancel, getResByKey("Warning")
            Exit Sub
        End If
        
        '如果需新增，新增GroupName和columnName到sheetName上，并得到新增字段的行号、列号
        
        Dim cell As range
        For Each cell In selectRange
            Call addGroupAndColoum(sheetName, groupName, columnName, cell.row, cell.column)
            Call getRowNumAndColumnNum(sheetName, groupName, columnName, rowNum, columnNum)
            Set BaseStationsheet = ThisWorkbook.Worksheets(sheetName)
            'maxRowNum = BaseStationsheet.range("a1048576").End(xlUp).row
            maxRowNum = BaseStationsheet.UsedRange.rows.count
            Set target = BaseStationsheet.range(getColumnNameFromColumnNum(columnNum) + "3" + ":" + getColumnNameFromColumnNum(columnNum) + CStr(maxRowNum))
            'Target.Validation.Delete
            '如果没有映射到MOC信息，则不联动修改基站传输页格式
            If mocIsNotNull(groupName, columnName) Then
                Call setRefStyle(ws, cell.row, cell.column, target)
            End If
        Next
       
        '在当前单元格上设置超链接
        linkCellRange = "R" + CStr(rowNum) + "C" + CStr(columnNum)

        ActiveSheet.Hyperlinks.Add Anchor:=Selection, address:="", SubAddress:= _
        "'" + sheetName + "'!" + linkCellRange, TextToDisplay:=LinkValueText.value
        
        Call setHyperlinkRangeFont(Selection)
        ActiveSheet.Columns(Selection.column).WrapText = False
        ActiveSheet.Columns(Selection.column).AutoFit
        
        '修改MappingDef
        Call setMappingDefIsRefValue(ActiveSheet, Selection)
        HasHistoryData = True
        HyperlinksForm.Hide
    Else
        MsgBox getResByKey("SelectCell"), vbExclamation + vbOKCancel, getResByKey("Warning")
    End If

    Unload Me
End Sub

'添加单元格格式引用
Sub setRefStyle(ws As Worksheet, rowNumber As Long, columnNumber As Long, target As range)
    Dim controlDef  As CControlDef
    Dim m_Str As String
    Dim groupName As String
    Dim columnName As String
    
    If getRangeGroupAndColumnName(ws, rowNumber, columnNumber, groupName, columnName) = True Then
        Set controlDef = getControlDefine(getResByKey("Board Style"), groupName, columnName)
        If controlDef Is Nothing Then
            Exit Sub
        End If
        m_Str = controlDef.lstValue
        If Len(m_Str) > 255 Then
            Dim valideDef As CValideDef
            Set valideDef = initDefaultDataSub.getInnerValideDef(getResByKey("Board Style") + "," + groupName + "," + columnName)
            If valideDef Is Nothing Then
                Set valideDef = addInnerValideDef(getResByKey("Board Style"), groupName, columnName, m_Str)
            End If
            m_Str = valideDef.getValidedef
        End If
        
        If Not controlDef Is Nothing Then
            On Error Resume Next
            If UCase(controlDef.dataType) = "ENUM" And controlDef.lstValue <> "" Then
                If target.Validation Is Nothing Then
                    With target.Validation
                       .Delete
                       .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=m_Str
                    End With
                Else
                    If Not targetHasFormula1(target) Then
                        With target.Validation
                            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=m_Str
                            .Modify Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=m_Str
                        End With
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Function mocIsNotNull(ByRef groupName As String, ByRef columnName As String) As Boolean
    Dim sheetDef As Worksheet
    Dim index As Long
    Dim defCollection As New Collection
    Dim mappingdefgroupName As String
    Dim mappingdefcolumnName As String
    Dim mocName As String
    Dim attributeName As String
    
    mocIsNotNull = False
    
    Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For index = 2 To sheetDef.range("a1048576").End(xlUp).row
        mappingdefgroupName = sheetDef.Cells(index, 2)
        mappingdefcolumnName = sheetDef.Cells(index, 3)
        mocName = sheetDef.Cells(index, 4)
        attributeName = sheetDef.Cells(index, 5)
        If mappingdefgroupName = groupName And columnName = mappingdefcolumnName And mocName <> "" And attributeName <> "" Then
            mocIsNotNull = True
            Exit For
        End If
    Next
End Function

Private Function targetHasFormula1(ByRef target As range) As Boolean
    On Error GoTo ErrorHandler
    targetHasFormula1 = True
    If target.Validation Is Nothing Then '没有有效性，则没有formula1
        targetHasFormula1 = False
        Exit Function
    End If
    
    Dim formula1 As String
    formula1 = target.Validation.formula1 '如果有formula1，则赋值成功，如果没有，则赋值出错，进入ErrorHandler
    If formula1 = "" Then targetHasFormula1 = False
    Exit Function
ErrorHandler:
    targetHasFormula1 = False
End Function

'在MappingDef中设置当前单元格所在的Group、Column:Is Reference=TRUE
Sub setMappingDefIsRefValue(sheet As Worksheet, range As range)
    Dim m_rowNum As Long
    Dim groupName As String
    Dim columnName As String
    Dim mappingDef As Worksheet
    
    Set mappingDef = Worksheets("MAPPING DEF")
    Call getGroupAndColumnName(sheet, range, groupName, columnName)
    
    If groupName <> "" And columnName <> "" Then
        For m_rowNum = 2 To mappingDef.range("a1048576").End(xlUp).row
            If sheet.name = mappingDef.Cells(m_rowNum, 1).value _
                And groupName = mappingDef.Cells(m_rowNum, 2).value _
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

Private Sub ListColumnName_Change()
    Call displayLinkValueText
End Sub

Private Sub ListGroupName_Change()
    Call displayLinkValueText
End Sub

Private Sub displayLinkValueText()
    Dim sheetName As String
    Dim groupName As String
    Dim columnName As String
    
    sheetName = Trim(ListHyperlinksForm.ListSheetName.text)
    groupName = Trim(ListHyperlinksForm.ListGroupName.text)
    columnName = Trim(ListHyperlinksForm.ListColumnName.text)
    
    If groupName = "" And columnName = "" Then
         LinkValueText.value = sheetName
         Exit Sub
    End If
    
    If columnName <> "" Then
        LinkValueText.value = sheetName + "\" + groupName + "\" + columnName
    Else
        LinkValueText.value = sheetName + "\" + groupName
    End If
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

'刷新ListSheetName.value和GroupNameListRef
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

'刷新ListGroupName.value和GroupNameListRef
Private Sub GroupNameListRef_Change()
    Dim sheetName, groupName As String
    Dim m_rowNum, m_colNum, m_colNum1, columnsNum As Long
    Dim ws As Worksheet
    
    ColumnNameListRef.Clear
    If (ListHyperlinksForm.SheetNameListRef.ListIndex <> -1) And (GroupNameListRef.ListIndex <> -1) Then
        sheetName = SheetNameListRef.List(SheetNameListRef.ListIndex)
        groupName = GroupNameListRef.List(GroupNameListRef.ListIndex)

        ListGroupName.value = groupName
        
        Set ws = Worksheets(sheetName)
        If sheetName = getResByKey("Comm Data") Then
            For m_rowNum = 1 To ws.range("a1048576").End(xlUp).row
                If groupName = ws.Cells(m_rowNum, 1).value Then
                    For m_colNum = 1 To ws.range("XFD" + CStr(m_rowNum + 1)).End(xlToLeft).column
                        ListHyperlinksForm.ColumnNameListRef.AddItem (ws.Cells(m_rowNum + 1, m_colNum).value)
                    Next
                    Exit For
                End If
            Next
        Else
            For m_colNum = 1 To ws.range("XFD1").End(xlToLeft).column
                If groupName = ws.Cells(1, m_colNum).value Then
                    columnsNum = ws.Cells(1, m_colNum).MergeArea.Columns.count
                    For m_colNum1 = m_colNum To m_colNum + columnsNum - 1
                        ColumnNameListRef.AddItem (ws.Cells(2, m_colNum1).value)
                    Next
                    Exit For
                End If
            Next
        End If
    End If
End Sub

'刷新ColumnNameListRef
Private Sub ColumnNameListRef_Change()
    Dim sheetName, groupName As String
    Dim columnName As String
        
    If (SheetNameListRef.ListIndex <> -1) And (GroupNameListRef.ListIndex <> -1) And (ColumnNameListRef.ListIndex <> -1) Then
        columnName = ColumnNameListRef.List(ColumnNameListRef.ListIndex)
        ListColumnName.value = columnName
    End If
End Sub

Private Sub ColumnNameListRef_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    AddButton_Click
End Sub
















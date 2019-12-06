VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ListHyperlinksForm 
   Caption         =   "添加引用"
   ClientHeight    =   6975
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
    Dim Target As Range
    Dim maxRowNum As Long
    Dim linkCellRange As String
    Dim selectRange As Range
    Set selectRange = Selection
    
    If Selection.Areas.count = 1 And Selection.Areas.Item(1).count = 1 And Selection.Interior.colorIndex <> SolidColorIdx And Selection.Interior.Pattern <> SolidPattern And Selection.Interior.colorIndex <> 34 And Selection.Interior.colorIndex <> 40 Then
        sheetName = Trim(ListHyperlinksForm.SheetNameListRef.List(SheetNameListRef.ListIndex))
        groupName = Trim(ListHyperlinksForm.ListGroupName.text)
        columnName = Trim(ListHyperlinksForm.ListColumnName.text)
        If sheetName = "" Or groupName = "" Or columnName = "" Then
            MsgBox getResByKey("SelectValideName"), vbExclamation + vbOKCancel, getResByKey("Warning")
            Exit Sub
        End If
        
        If Trim(LinkValueText.value) = "" Or UBound(Split(LinkValueText.value, "\")) <> 2 Then
            MsgBox getResByKey("SelectValideData"), vbExclamation + vbOKCancel, getResByKey("Warning")
            Exit Sub
        End If
        
                If InStr(Trim(groupName), "[") <> 0 Or InStr(Trim(columnName), "[") <> 0 Then
            MsgBox getResByKey("ReferenceInvalidCharacter"), vbExclamation + vbOKCancel, getResByKey("Warning")
            ListHyperlinksForm.Hide
            Exit Sub
        End If

        If InStr(Trim(LinkValueText.value), "[") <> 0 Then
            rowIndexStr = Mid(LinkValueText.value, InStr(LinkValueText.value, "[") + 1, InStr(LinkValueText.value, "]") - InStr(LinkValueText.value, "[") - 1)
            If (rowIndexStr <> "") And (checkValueIfInteger(rowIndexStr) = True) Then
                rowNum = rowNum + 1 + CLng(rowIndexStr)
            Else
                MsgBox getResByKey("SelectValideIndex"), vbExclamation + vbOKCancel, getResByKey("Warning")
                Exit Sub
            End If
        End If
        
        Dim Cell As Range
        For Each Cell In selectRange
            Call addGroupAndColoum(sheetName, groupName, columnName, Cell.row, Cell.column)
            Call getRowNumAndColumnNum(sheetName, groupName, columnName, rowNum, columnNum)
            Set BaseStationsheet = ThisWorkbook.Worksheets(sheetName)
            'maxRowNum = BaseStationsheet.range("a1048576").End(xlUp).row
            maxRowNum = BaseStationsheet.UsedRange.Rows.count
            Set Target = BaseStationsheet.Range(getColumnNameFromColumnNum(columnNum) + "3" + ":" + getColumnNameFromColumnNum(columnNum) + CStr(maxRowNum))
            'Target.Validation.Delete
            '???????MOC??,?????????????
            If mocIsNotNull(groupName, columnName) Then
                Call setRefStyle(ws, Cell.row, Cell.column, Target)
            End If
        Next
        linkCellRange = "R" + CStr(rowNum) + "C" + CStr(columnNum)
        'Selection.value = ""
        ActiveSheet.Hyperlinks.Add Anchor:=Selection, address:="", SubAddress:= _
        "'" + sheetName + "'!" + linkCellRange, TextToDisplay:=LinkValueText.value
        Call setHyperlinkRangeFont(Selection)
        ActiveSheet.columns(Selection.column).WrapText = False
        ActiveSheet.columns(Selection.column).AutoFit
        Call setMappingDefIsRefValue(ActiveSheet, Selection)
        HasHistoryData = True
        HyperlinksForm.Hide
    Else
        MsgBox getResByKey("SelectCell"), vbExclamation + vbOKCancel, getResByKey("Warning")
    End If
    'Call addListRef(sheetName, GroupName, columnName)
    Unload Me
End Sub

'?????????
Sub setRefStyle(ws As Worksheet, rowNumber As Long, columnNumber As Long, Target As Range)
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
                If Target.Validation Is Nothing Then
                    With Target.Validation
                       .Delete
                       .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=m_Str
                    End With
                Else
                    If Not targetHasFormula1(Target) Then
                        With Target.Validation
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
    For index = 2 To sheetDef.Range("a1048576").End(xlUp).row
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

Private Function targetHasFormula1(ByRef Target As Range) As Boolean
    On Error GoTo ErrorHandler
    targetHasFormula1 = True
    If Target.Validation Is Nothing Then '?????,???formula1
        targetHasFormula1 = False
        Exit Function
    End If
    
    Dim formula1 As String
    formula1 = Target.Validation.formula1 '???formula1,?????,????,?????,??ErrorHandler
    If formula1 = "" Then targetHasFormula1 = False
    Exit Function
ErrorHandler:
    targetHasFormula1 = False
End Function

Public Sub setHyperlinkRangeFont(ByRef certainRange As Range)
    With certainRange.Font
        .name = "Arial"
        .Size = 10
    End With
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

Sub setMappingDefIsRefValue(sheet As Worksheet, Range As Range)
    Dim m_rowNum As Long
    Dim groupName As String
    Dim columnName As String
    Dim mappingDef As Worksheet
    Set mappingDef = Worksheets("MAPPING DEF")
    Call getGroupAndColumnName(sheet, Range, groupName, columnName)
    If groupName <> "" And columnName <> "" Then
        For m_rowNum = 2 To mappingDef.Range("a1048576").End(xlUp).row
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
    LinkValueText.value = ListHyperlinksForm.ListSheetName.text + "\" + ListHyperlinksForm.ListGroupName.text + "\" + ListHyperlinksForm.ListColumnName.text
End Sub

Private Sub ListGroupName_Change()
    LinkValueText.value = ListHyperlinksForm.ListSheetName.text + "\" + ListHyperlinksForm.ListGroupName.text + "\" + ListHyperlinksForm.ListColumnName.text
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
            For m_rowNum = 1 To ws.Range("a1048576").End(xlUp).row
                If ws.Cells(m_rowNum, 1).Interior.colorIndex = 34 Then
                    ListHyperlinksForm.GroupNameListRef.AddItem (ws.Cells(m_rowNum, 1).value)
                End If
            Next
        Else
            For m_colNum = 1 To ws.Range("XFD1").End(xlToLeft).column
                If Trim(ws.Cells(1, m_colNum).value) <> "" Then
                    GroupNameListRef.AddItem (ws.Cells(1, m_colNum).value)
                End If
            Next
        End If
    End If
End Sub
Private Sub GroupNameListRef_Change()
    Dim sheetName, groupName As String
    Dim m_rowNum, m_colNum, m_colNum1, columnsNum As Long
    ColumnNameListRef.Clear
    Dim ws As Worksheet
    
    If (ListHyperlinksForm.SheetNameListRef.ListIndex <> -1) And (GroupNameListRef.ListIndex <> -1) Then
        sheetName = SheetNameListRef.List(SheetNameListRef.ListIndex)
        groupName = GroupNameListRef.List(GroupNameListRef.ListIndex)
        LinkValueText.value = sheetName + "\" + groupName
        ListGroupName.value = groupName
        Set ws = Worksheets(sheetName)
        If sheetName = getResByKey("Comm Data") Then
            For m_rowNum = 1 To ws.Range("a1048576").End(xlUp).row
                If groupName = ws.Cells(m_rowNum, 1).value Then
                    For m_colNum = 1 To ws.Range("XFD" + CStr(m_rowNum + 1)).End(xlToLeft).column
                        ListHyperlinksForm.ColumnNameListRef.AddItem (ws.Cells(m_rowNum + 1, m_colNum).value)
                    Next
                    Exit For
                End If
            Next
        Else
            For m_colNum = 1 To ws.Range("XFD1").End(xlToLeft).column
                If groupName = ws.Cells(1, m_colNum).value Then
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
    Dim sheetName, groupName As String
    Dim columnName As String
        
    If (SheetNameListRef.ListIndex <> -1) And (GroupNameListRef.ListIndex <> -1) And (ColumnNameListRef.ListIndex <> -1) Then
        sheetName = SheetNameListRef.List(SheetNameListRef.ListIndex)
        groupName = GroupNameListRef.List(GroupNameListRef.ListIndex)
        columnName = ColumnNameListRef.List(ColumnNameListRef.ListIndex)
        ListColumnName.value = columnName
        LinkValueText.value = sheetName + "\" + ListHyperlinksForm.ListGroupName.text + "\" + ListHyperlinksForm.ListColumnName.text
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

















VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HyperlinksForm 
   Caption         =   "Add Reference"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9615.001
   OleObjectBlob   =   "HyperlinksForm.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "HyperlinksForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim HasHistoryData As Boolean
Private Sub AddButton_Click()
    On Error Resume Next
    'ûѡ�е�Ԫ��ʱӦ����ѡ�ж����Ԫ��ʱֻ���һ����Ԫ��
    Dim sheetName As String
    Dim GroupName As String
    Dim columnName As String
    'Dim rangeStr, colStr As String
    Dim rowNum As Long
    Dim columnNum As Long
    Dim rowIndexStr  As String
    Dim linkCellRange As String
    
    If Selection.Areas.count = 1 And Selection.Areas.Item(1).count = 1 Then
        If (HyperlinksForm.SheetNameList.ListIndex <> -1) _
            And (HyperlinksForm.GroupNameList.ListIndex <> -1) _
            And (HyperlinksForm.ColumnNameList.ListIndex <> -1) _
            And Trim(LinkValueText.value) <> "" _
            And UBound(Split(LinkValueText.value, "\")) = 2 Then
            If isPatternSheet(ActiveSheet.name) = False Then
                MsgBox getResByKey("SelectPattern"), vbExclamation + vbOKCancel, getResByKey("Comm Data")
                HyperlinksForm.Hide
                Exit Sub
            End If
            sheetName = SheetNameList.List(SheetNameList.ListIndex)
            GroupName = GroupNameList.List(GroupNameList.ListIndex)
            columnName = ColumnNameList.List(ColumnNameList.ListIndex)
            
             If InStr(Trim(GroupName), "[") <> 0 Or InStr(Trim(columnName), "[") <> 0 Then
                 MsgBox getResByKey("ReferenceInvalidCharacter"), vbExclamation + vbOKCancel, getResByKey("Warning")
                 HyperlinksForm.Hide
                 Exit Sub
            End If

            Call getRowNumAndColumnNum(sheetName, GroupName, columnName, rowNum, columnNum)
            If sheetName = getResByKey("Comm Data") Then
                If InStr(Trim(LinkValueText.value), "[") <> 0 Then
                    rowIndexStr = Mid(LinkValueText.value, InStr(LinkValueText.value, "[") + 1, InStr(LinkValueText.value, "]") - InStr(LinkValueText.value, "[") - 1)
                    If (rowIndexStr <> "") And (checkValueIfInteger(rowIndexStr) = True) Then
                        rowNum = rowNum + 1 + CLng(rowIndexStr)
                    Else
                        MsgBox getResByKey("SelectValideIndex"), vbExclamation + vbOKCancel, getResByKey("Warning")
                        Exit Sub
                    End If
                Else
                    MsgBox getResByKey("SelectValideIndex"), vbExclamation + vbOKCancel, getResByKey("Warning")
                    Exit Sub
                End If
            End If
            'If columnNum \ 26 > 0 Then
            '    colStr = Chr(64 + columnNum \ 26) + Chr(64 + columnNum Mod 26)
            'Else
            '    colStr = Chr(64 + columnNum)
            'End If
            
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
        ElseIf Trim(LinkValueText.value) = "" Or UBound(Split(LinkValueText.value, "\")) <> 2 Then
            MsgBox getResByKey("SelectValideData"), vbExclamation + vbOKCancel, getResByKey("Warning")
        Else
            MsgBox getResByKey("SelectValideName"), vbExclamation + vbOKCancel, getResByKey("Warning")
        End If
    Else
        MsgBox getResByKey("SelectCell"), vbExclamation + vbOKCancel, getResByKey("Warning")
    End If
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
    
    Call getGroupAndColumnName(sheet, range, GroupName, columnName)
    If GroupName <> "" And columnName <> "" Then
        For m_rowNum = 2 To Worksheets("MAPPING DEF").range("a1048576").End(xlUp).row
            If sheet.name = Worksheets("MAPPING DEF").Cells(m_rowNum, 1).value _
                And GroupName = Worksheets("MAPPING DEF").Cells(m_rowNum, 2).value _
                And columnName = Worksheets("MAPPING DEF").Cells(m_rowNum, 3).value Then
                Worksheets("MAPPING DEF").Cells(m_rowNum, 6).value = "TRUE"
                Exit For
            End If
        Next
    End If
End Sub
Sub getRowNumAndColumnNum(sheetName As String, GroupName As String, columnName As String, rowNum As Long, columnNum As Long)
    If sheetName = getResByKey("Comm Data") Then
        For m_rowNum = 1 To Worksheets(sheetName).range("a1048576").End(xlUp).row
            If GroupName = Worksheets(sheetName).Cells(m_rowNum, 1).value Then
                For m_colNum = 1 To Worksheets(sheetName).range("XFD" + CStr(m_rowNum + 1)).End(xlToLeft).column
                    If columnName = Worksheets(sheetName).Cells(m_rowNum + 1, m_colNum).value Then
                        rowNum = m_rowNum + 1
                        columnNum = m_colNum
                        Exit For
                    End If
                Next
                Exit For
            End If
        Next
    Else
        For m_colNum = 1 To Worksheets(sheetName).range("XFD1").End(xlToLeft).column
            If GroupName = Worksheets(sheetName).Cells(1, m_colNum).value Then
                columnsNum = Worksheets(sheetName).Cells(1, m_colNum).MergeArea.columns.count
                For m_colNum1 = m_colNum To m_colNum + columnsNum - 1
                    If columnName = Worksheets(sheetName).Cells(2, m_colNum1).value Then
                        rowNum = 2
                        columnNum = m_colNum1
                        Exit For
                    End If
                Next
                Exit For
            End If
        Next
    End If
End Sub
Private Sub CancelButton_Click()
    HasHistoryData = True
    HyperlinksForm.Hide
End Sub
Private Sub UserForm_Activate()
    HyperlinksForm.Caption = getResByKey("Bar_Refrence")
    If HasHistoryData = False Or HyperlinksForm.SheetNameList.ListCount < 1 Then
        SetSheetNameList
    End If
End Sub
Private Sub SetSheetNameList()
    HyperlinksForm.SheetNameList.Clear
    If GetCommonSheetName <> "" Then
         HyperlinksForm.SheetNameList.AddItem (getResByKey("Comm Data"))
    End If
    
    HyperlinksForm.SheetNameList.AddItem (getResByKey("BaseTransPort"))
End Sub
Private Sub SheetNameList_Change()
    Dim sheetName As String
    Dim m_rowNum, m_colNum As Long
    
    HyperlinksForm.GroupNameList.Clear
    If HyperlinksForm.SheetNameList.ListIndex <> -1 Then
        sheetName = SheetNameList.List(SheetNameList.ListIndex)
        LinkValueText.value = sheetName
        If sheetName = getResByKey("Comm Data") Then
            For m_rowNum = 1 To Worksheets(sheetName).range("a1048576").End(xlUp).row
                If Worksheets(sheetName).Cells(m_rowNum, 1).Interior.colorIndex = 34 Then
                    HyperlinksForm.GroupNameList.AddItem (Worksheets(sheetName).Cells(m_rowNum, 1).value)
                End If
            Next
        Else
            For m_colNum = 1 To Worksheets(sheetName).range("XFD1").End(xlToLeft).column
                If Trim(Worksheets(sheetName).Cells(1, m_colNum).value) <> "" Then
                    HyperlinksForm.GroupNameList.AddItem (Worksheets(sheetName).Cells(1, m_colNum).value)
                End If
            Next
        End If
    End If
End Sub
Private Sub GroupNameList_Change()
    Dim sheetName, GroupName As String
    Dim m_rowNum, m_colNum, m_colNum1, columnsNum As Long
    
    HyperlinksForm.ColumnNameList.Clear
    If (HyperlinksForm.SheetNameList.ListIndex <> -1) And (HyperlinksForm.GroupNameList.ListIndex <> -1) Then
        sheetName = SheetNameList.List(SheetNameList.ListIndex)
        GroupName = GroupNameList.List(GroupNameList.ListIndex)
        LinkValueText.value = sheetName + "\" + GroupName
        If sheetName = getResByKey("Comm Data") Then
            For m_rowNum = 1 To Worksheets(sheetName).range("a1048576").End(xlUp).row
                If GroupName = Worksheets(sheetName).Cells(m_rowNum, 1).value Then
                    For m_colNum = 1 To Worksheets(sheetName).range("XFD" + CStr(m_rowNum + 1)).End(xlToLeft).column
                        HyperlinksForm.ColumnNameList.AddItem (Worksheets(sheetName).Cells(m_rowNum + 1, m_colNum).value)
                    Next
                    Exit For
                End If
            Next
        Else
            For m_colNum = 1 To Worksheets(sheetName).range("XFD1").End(xlToLeft).column
                If GroupName = Worksheets(sheetName).Cells(1, m_colNum).value Then
                    columnsNum = Worksheets(sheetName).Cells(1, m_colNum).MergeArea.columns.count
                    For m_colNum1 = m_colNum To m_colNum + columnsNum - 1
                        HyperlinksForm.ColumnNameList.AddItem (Worksheets(sheetName).Cells(2, m_colNum1).value)
                    Next
                    Exit For
                End If
            Next
        End If
    End If
End Sub
Private Sub ColumnNameList_Change()
    Dim sheetName, GroupName As String
    Dim columnName As String
        
    If (HyperlinksForm.SheetNameList.ListIndex <> -1) And (HyperlinksForm.GroupNameList.ListIndex <> -1) And (HyperlinksForm.ColumnNameList.ListIndex <> -1) Then
        sheetName = SheetNameList.List(SheetNameList.ListIndex)
        GroupName = GroupNameList.List(GroupNameList.ListIndex)
        columnName = ColumnNameList.List(ColumnNameList.ListIndex)
        If sheetName = getResByKey("Comm Data") Then
            LinkValueText.value = sheetName + "\" + GroupName + "\" + columnName + "[0]"
        Else
            LinkValueText.value = sheetName + "\" + GroupName + "\" + columnName
        End If
    End If
End Sub
Private Sub ColumnNameList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    AddButton_Click
End Sub


Attribute VB_Name = "Util"
Option Explicit

Global Const g_strInvalidChar4Sql As String = "'"
Global Const g_strInvalidChar4PureFileName  As String = "\/:*?<>|"""
Global Const g_strInvalidChar4Path          As String = "/*?<>|"""

Private Const SW_SHOWNORMAL = 1
Public HasHistoryData As Boolean



'装载用于添加「Base Station Transport Data」页「*Site Template」列侯选值的窗体
Sub addTemplate()

    Load NoneLteTemplateForm
    NoneLteTemplateForm.Show

End Sub

'从指定sheet页的指定行，查找指定列，返回列号,通过attrName和MocName获取到
Public Function getColNum(sheetName As String, RecordRow As Long, attrName As String, mocName As String) As Long
    On Error Resume Next
    Dim m_ColNum As Long
    Dim m_rowNum As Long
    Dim ColName As String
    Dim colGroupName As String
    
    Dim flag As Boolean
    Dim MAPPINGDEF As Worksheet
    Dim ws As Worksheet
    
    Set MAPPINGDEF = ThisWorkbook.Worksheets("MAPPING DEF")
    flag = False
    getColNum = -1
    For m_rowNum = 2 To MAPPINGDEF.Range("a1048576").End(xlUp).row
        If UCase(attrName) = UCase(MAPPINGDEF.Cells(m_rowNum, 5).value) _
           And UCase(sheetName) = UCase(MAPPINGDEF.Cells(m_rowNum, 1).value) _
           And UCase(mocName) = UCase(MAPPINGDEF.Cells(m_rowNum, 4).value) Then
            ColName = MAPPINGDEF.Cells(m_rowNum, 3).value
            colGroupName = MAPPINGDEF.Cells(m_rowNum, 2).value
            flag = True
            Exit For
        End If
    Next
    If flag = True Then
        Set ws = ThisWorkbook.Worksheets(sheetName)
        For m_ColNum = 1 To ws.Range("XFD" + CStr(RecordRow)).End(xlToLeft).column
            If get_GroupName(sheetName, m_ColNum) = colGroupName Then
                If GetDesStr(ColName) = GetDesStr(ws.Cells(RecordRow, m_ColNum).value) Then
                    getColNum = m_ColNum
                    Exit For
                End If
            End If
        Next
    End If
End Function
Public Function GetMainSheetName() As String
       On Error Resume Next
        Dim name As String
        Dim RowNum As Long
        Dim sheetDef As Worksheet
        Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
        For RowNum = 1 To sheetDef.Range("a1048576").End(xlUp).row
            If sheetDef.Cells(RowNum, 2).value = "MAIN" Then
                name = sheetDef.Cells(RowNum, 1).value
                Exit For
            End If
        Next
        GetMainSheetName = name
End Function



'从普通页取得Group name
Public Function get_GroupName(sheetName As String, colNum As Long) As String
        Dim index As Long
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(sheetName)
        For index = colNum To 1 Step -1
            If Not isEmpty(ws.Cells(1, index).value) And ws.Cells(1, index).value <> "" Then
                get_GroupName = ws.Cells(1, index).value
                Exit Function
            End If
        Next
        get_GroupName = ""
End Function

'从普通页取得Colum name
Public Function get_ColumnName(ByVal sheetName As String, colNum As Long) As String
        Dim index As Long
        get_ColumnName = ThisWorkbook.Worksheets(sheetName).Cells(2, colNum)
End Function

Public Function isNum(contedType As String) As Boolean
    If contedType = "Integer" Or contedType = "UInteger" Then
        isNum = True
    Else
        isNum = False
    End If
End Function

Sub destroyMenuStatus()
    With Application
        .CommandBars("Row").Reset
        .CommandBars("Column").Reset
        .CommandBars("Cell").Reset
        .CommandBars("Ply").Reset
    End With
End Sub
Sub insertAndDeleteControl(ByRef flag As Boolean)
    On Error Resume Next
    With Application
        With .CommandBars("Row")
            .FindControl(ID:=3183).Enabled = flag '插入
            .FindControl(ID:=296).Enabled = flag '行
            .FindControl(ID:=293).Enabled = flag '删除
        End With
        With .CommandBars("Column")
            .FindControl(ID:=3183).Enabled = flag '插入
            .FindControl(ID:=297).Enabled = flag '行
            .FindControl(ID:=294).Enabled = flag '删除
        End With
        With .CommandBars("Cell")
            .FindControl(ID:=3181).Enabled = flag '插入
            .FindControl(ID:=295).Enabled = flag '行
            .FindControl(ID:=292).Enabled = flag '删除
        End With
    End With
End Sub

Sub initMenuStatus(sh As Worksheet)

    Call initTempSheetControl(True)
    Call insertAndDeleteControl(True)
End Sub

Function getSheetType(sheetName As String) As String
    Dim m_rowNum As Long
    Dim sheetDef As Worksheet
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
    For m_rowNum = 2 To sheetDef.Range("a1048576").End(xlUp).row
        If sheetName = sheetDef.Cells(m_rowNum, 1).value Then
            getSheetType = UCase(sheetDef.Cells(m_rowNum, 2).value)
            Exit Function
        End If
    Next
    getSheetType = ""
End Function

'将比较字符串整形
Function GetDesStr(srcStr As String) As String
    GetDesStr = UCase(Trim(srcStr))
End Function

Public Function containsASheet(ByRef wb As Workbook, ByRef sheetName As String) As Boolean
    On Error GoTo ErrorHandler
    containsASheet = True
    Dim sheet As Worksheet
    Set sheet = wb.Worksheets(sheetName)
    Exit Function
ErrorHandler:
    containsASheet = False
End Function

Public Sub popUpWbCannotSaveMsgbox()
    Call MsgBox(getResByKey("CannotSaveWb"), vbExclamation)
End Sub


Public Function hasFreqColumn() As Boolean
    Dim sheetDef As Worksheet
    Dim index As Long
    Dim attrName As String
    
    hasFreqColumn = False
    
    Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For index = 2 To sheetDef.Range("a1048576").End(xlUp).row
        attrName = sheetDef.Cells(index, 5)
        If attrName = "BCCHFREQ" Then
            hasFreqColumn = True
            Exit For
        End If
    Next
End Function

Public Function hasNonFreqColumn() As Boolean
    Dim sheetDef As Worksheet
    Dim index As Long
    Dim attrName As String
    
    hasNonFreqColumn = False
    
    Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For index = 2 To sheetDef.Range("a1048576").End(xlUp).row
        attrName = sheetDef.Cells(index, 5)
        If attrName = "NONBCCHFREQLIST" Then
            hasNonFreqColumn = True
            Exit For
        End If
    Next
End Function

Public Function isInCollection(ByRef collect As Collection, ByRef str As String) As Boolean
    isInCollection = False
    
    If collect Is Nothing Then Exit Function
    
    Dim var As Variant
    For Each var In collect
        If CStr(var) = str Then
            isInCollection = True
            Exit Function
        End If
    Next
End Function

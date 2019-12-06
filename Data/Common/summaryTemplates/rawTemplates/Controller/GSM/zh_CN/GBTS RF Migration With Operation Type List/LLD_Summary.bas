Attribute VB_Name = "LLD_Summary"
Option Explicit
Public lldModelFlag As String

Function getEndRow(readRowNum As Long) As Long
        Dim rowCount, colIndex As Long
        Dim curStatus, preStatus As Boolean
        preStatus = False
        For rowCount = 1 To Worksheets("COMMON").Range("a1048576").End(xlUp).row
                curStatus = True
                For colIndex = 1 To Worksheets("COMMON").Range("XFD2").End(xlToLeft).column
                    If Worksheets("COMMON").Cells(readRowNum + rowCount, colIndex).value <> "" Then
                        curStatus = False
                    End If
                Next
                If preStatus = True And curStatus = False Then
                    Exit For
                End If
                preStatus = curStatus
        Next
        getEndRow = rowCount - 1
End Function

'将比较字符串整形
Function GetDesStr(srcStr As String) As String
    GetDesStr = UCase(Trim(srcStr))
End Function

'从指定sheet页的指定行，查找指定列，返回列号
Function Get_Col(sheetName As String, recordRow As Long, ColValue As String) As Long
    Dim m_colNum As Long
    
    Get_Col = -1
    For m_colNum = 1 To Worksheets(sheetName).Range("XFD" + CStr(recordRow)).End(xlToLeft).column
        If GetDesStr(ColValue) = GetDesStr(Worksheets(sheetName).Cells(recordRow, m_colNum).value) Then
                Get_Col = m_colNum
            Exit For
        End If
    Next
End Function

'从指定sheet页查找group所在行
Function Get_GroupRow(sheetName As String, groupName As String) As Long
    Dim m_rowNum As Long
    Dim f_flag As Boolean
    
    f_flag = False
    For m_rowNum = 1 To Worksheets(sheetName).Range("a1048576").End(xlUp).row
        If GetDesStr(groupName) = GetDesStr(Worksheets(sheetName).Cells(m_rowNum, 1).value) Then
            f_flag = True
            Exit For
        End If
    Next
    If f_flag = False Then
        'MsgBox sheetName & "缺少Group：" & groupName, vbExclamation, "Warning"
    End If
    
    Get_GroupRow = m_rowNum
    
End Function






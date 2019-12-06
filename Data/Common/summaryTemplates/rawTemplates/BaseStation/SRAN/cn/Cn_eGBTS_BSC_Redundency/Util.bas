Attribute VB_Name = "Util"
Option Explicit

Public Function GetMainSheetName() As String
       On Error Resume Next
        Dim name As String
        Dim rowNum As Long
        Dim sheetDef As Worksheet
        Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
        For rowNum = 1 To sheetDef.range("a65536").End(xlUp).row
            If sheetDef.Cells(rowNum, 2).value = "MAIN" Then
                name = sheetDef.Cells(rowNum, 1).value
                Exit For
            End If
        Next
        GetMainSheetName = name
End Function

Public Function isNum(contedType As String) As Boolean
    If contedType = "Integer" Or contedType = "UInteger" Then
        isNum = True
    Else
        isNum = False
    End If
End Function

Function getColStr(ByVal NumVal As Long) As String
    Dim str As String
    Dim strs() As String
    
    If NumVal > 256 Or NumVal < 1 Then
        getColStr = ""
    Else
        str = Cells(NumVal).address
        strs = Split(str, "$", -1)
        getColStr = strs(1)
    End If
End Function

Function GetDesStr(srcStr As String) As String
    GetDesStr = UCase(Trim(srcStr))
End Function

Function calculateColumnName(ByRef columnNumber As Long) As String
    Dim myRange As range
    Set myRange = Cells(1, columnNumber)    '指定该列标号的任意单元格
    calculateColumnName = Left(myRange.range("A1").address(True, False), _
        InStr(1, myRange.range("A1").address(True, False), "$", 1) - 1)
    Set myRange = Nothing
End Function

Function Contains(coll As Collection, key As String) As Boolean
        On Error GoTo NotFound
        Call coll(key)
        Contains = True
        Exit Function
NotFound:
        Contains = False
End Function

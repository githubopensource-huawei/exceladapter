Attribute VB_Name = "C_UtilityModule"
Option Explicit

Public Function IsInCollection(Value As String, col As Collection) As Boolean
    On Error GoTo E
    col.Item (Value)
    IsInCollection = True
    Exit Function
E:
    IsInCollection = False
End Function

Public Function IsInNameCollection(Value As String, col As NameCollectionClass) As Boolean
    On Error GoTo E
    col.Item (Value)
    IsInNameCollection = True
    Exit Function
E:
    IsInNameCollection = False
End Function

Public Function FileExists(FileName As String) As Boolean
    FileExists = False
    
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(FileName) Then
        FileExists = True
    End If

    Set fs = Nothing
End Function

Public Function GetColName(ColIndex As Integer) As String
    GetColName = ""
    Dim iTemp As Integer, i As Integer
    iTemp = ColIndex Mod 26
    If iTemp > 0 Then
        GetColName = Chr(64 + iTemp)
    ElseIf iTemp = 0 Then
        GetColName = Chr(64 + iTemp + 26)
    End If
    
    If ColIndex > 26 Then
        i = 1
        Do While i < 10
            If i * 26 + iTemp = ColIndex Then
                If iTemp = 0 Then
                    i = i - 1
                End If
                Exit Do
            End If
            i = i + 1
        Loop
        GetColName = Chr(64 + i) + GetColName
    End If
End Function

Public Function IsEmptyRow(Sheet As Worksheet, Row As Long)
    IsEmptyRow = True
    Dim i As Long, iEnd As Long
    iEnd = Sheet.UsedRange.Columns.Count
    iEnd = Sheet.UsedRange.Columns(iEnd).Column
    For i = 1 To iEnd
        If Sheet.Cells(Row, i).Value <> "" Then
            IsEmptyRow = False
            Exit Function
        End If
    Next i
End Function

Public Function GetWorkbook(FileName As String) As Workbook
    Set GetWorkbook = Nothing
    Dim w As Workbook
    For Each w In Application.Workbooks
        If w.FullName = FileName Then
            Set GetWorkbook = w
            Exit For
        End If
    Next w
End Function

Public Function FormatStr(SourceStr As String, ParamArray args() As Variant)
    FormatStr = SourceStr
    Dim i As Long, s As String
    For i = 0 To UBound(args)
        s = args(i)
        FormatStr = Replace(FormatStr, GetFormatIndex(i + 1), s)
    Next i
End Function

Public Function GetFormatIndex(index As Long) As String
    Select Case index
        Case Is < 10
            GetFormatIndex = "@@" + CStr(index)
        Case Is < 36
            GetFormatIndex = "@@" + Chr(65 - 10 + index)
        Case Else
            GetFormatIndex = ""
    End Select
End Function

Public Function GetLastRowIndex(sht As Worksheet) As Long
    Dim i As Long, j As Integer, iRow As Long, iCol As Integer, IsEmptyRow As Boolean, s As String
    iCol = sht.UsedRange.Columns.Count
    iCol = sht.UsedRange.Columns(iCol).Column
    iRow = sht.UsedRange.Rows.Count
    iRow = sht.UsedRange.Rows(iRow).Row
    For i = iRow To 1 Step -1
        IsEmptyRow = True
        For j = 1 To iCol
            s = sht.Cells(i, j).Value
            If s <> "" Then
                IsEmptyRow = False
                Exit For
            End If
        Next j

        If Not IsEmptyRow Then
            GetLastRowIndex = i
            Exit Function
        End If
    Next i
    GetLastRowIndex = 0
End Function
        
Public Sub ClearCollection(col As Collection)
    Do While col.Count > 0
        col.Remove (1)
    Loop
End Sub

Public Sub SetBorderWeight(Range As Range, Top As Integer, Bottom As Integer, Left As Integer, Right As Integer)
    With Range
        .Borders(xlEdgeTop).Weight = Top
        .Borders(xlEdgeBottom).Weight = Bottom
        .Borders(xlEdgeLeft).Weight = Left
        .Borders(xlEdgeRight).Weight = Right
    End With
End Sub

Public Sub SetRangeValidation(r As Range, Attr As AttrClass)
    Dim s As String, isAdded As Boolean
    r.Validation.Delete
    isAdded = False
    Select Case Attr.DataType
        Case atInteger
            If IsNumeric(Attr.MinValue) And IsNumeric(Attr.MaxValue) Then
                s = FormatStr(ERR_MSG_RANGE, Attr.MinValue, Attr.MaxValue)
                r.Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                    Operator:=xlBetween, Formula1:=Attr.MinValue, Formula2:=Attr.MaxValue
                isAdded = True
            End If
        Case atEnum
            If Attr.ValueList <> "" Then
                s = FormatStr(ERR_MSG_ENUM, Attr.ValueList)
                r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                    Operator:=xlBetween, Formula1:=Attr.ValueList
                isAdded = True
            End If
        Case atString
            If IsNumeric(Attr.MinValue) And IsNumeric(Attr.MaxValue) Then
                s = FormatStr(ERR_MSG_LENGTH, Attr.MinValue, Attr.MaxValue)
                r.Validation.Add Type:=xlValidateTextLength, AlertStyle:=xlValidAlertStop, _
                    Operator:=xlBetween, Formula1:=Attr.MinValue, Formula2:=Attr.MaxValue
                isAdded = True
            End If
            r.NumberFormatLocal = "@"
    End Select

    If isAdded Then
        If r.Validation.Formula1 <> "" Then
            With r.Validation
                .ErrorTitle = ERR_TITLE_PROMPT
                .ErrorMessage = s
            End With
        End If
    Else
        MsgBox "Attr " + Attr.Name + " has no validation."
    End If
End Sub

Public Sub ExportCode()
    Dim Modules As Variant
    Modules = Array("Common", "Template", "Worksheet", "Form", "ConvertTemplate", "DoubleFrequency")

    Dim wk As Workbook
    Set wk = Application.ActiveWorkbook
    Dim c As Variant, s As String, strHomePath As String
    strHomePath = wk.Path + "\code_" + Format(Date, "yyyy-mm-dd") + "_" + Format(Time, "hh-mm-ss")
    Call MkDir(strHomePath)
    strHomePath = strHomePath + "\"
    Dim i As Integer, Paths(5) As String
    For i = 0 To 5
        Paths(i) = strHomePath + Modules(i) + "\"
        Call MkDir(Paths(i))
    Next i

    For Each c In wk.VBProject.VBComponents
        Application.StatusBar = "Exporting " + c.Name + "..."
        Select Case c.Type
            Case 1
                If InStr(c.Name, "C_") = 1 Then
                    s = Paths(0) + c.Name + ".bas"
                ElseIf InStr(c.Name, "T_") = 1 Then
                    s = Paths(1) + c.Name + ".bas"
                ElseIf InStr(c.Name, "CT_") = 1 Then
                    s = Paths(4) + c.Name + ".bas"
                ElseIf InStr(c.Name, "DF_") = 1 Then
                    s = Paths(5) + c.Name + ".bas"
                Else
                    s = strHomePath + c.Name + ".bas"
                End If
            Case 2
                s = Paths(0) + c.Name + ".cls"
            Case 3
                s = Paths(3) + c.Name + ".frm"
            Case Else
                s = Paths(2) + c.Name + ".cls"
        End Select

        c.Export FileName:=s
    Next c

    Application.StatusBar = "Finished."
End Sub

Public Sub DeleteEmptyRowInSheetEnding(ByVal MocName As String)
    Dim sht As Worksheet
    Set sht = Sheets(MocName)
    
    Dim iEnd1 As Long, iEnd2 As Long
    iEnd1 = GetLastRowIndex(sht)
    If iEnd1 <= 0 Then
        Exit Sub
    End If
    iEnd2 = sht.UsedRange.Rows.Count
    iEnd2 = sht.UsedRange.Rows(iEnd2).Row
    
    Dim i As Long
    For i = iEnd2 To iEnd1 + 1 Step -1
        sht.Rows(i).Delete
    Next i
End Sub

Public Function SheetExist(wk As Workbook, ByVal SheetName As String) As Boolean
    On Error GoTo E
    Dim sht As Worksheet
    Set sht = wk.Sheets(SheetName)
    SheetExist = True
    Exit Function
E:
    SheetExist = False
End Function

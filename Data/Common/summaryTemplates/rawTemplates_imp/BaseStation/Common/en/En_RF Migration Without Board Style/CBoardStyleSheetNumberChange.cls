VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBoardStyleSheetNumberChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const delimiter_ As String = "_"
Private boardStyleSheetPrefix_ As String
Private baseBoardStyleSheetName_ As String
Private temlateBrdStyleShtName_ As String
'

Public Sub keepOneBaseBoardStyleSheet()
    Dim sheet As Worksheet, sheetName As String
    Dim temShtExistFlage As Boolean
    Dim sheetNameArr() As String
    temShtExistFlage = False
    
    For Each sheet In ThisWorkbook.Worksheets
        If sheetName = "" Then
            sheetName = sheet.name
        Else
            sheetName = sheetName + "," + sheet.name
        End If
    Next
    
    If sheetName <> "" Then
        sheetNameArr = Split(sheetName, ",")
        Call QuickSort(sheetNameArr())
    End If
    
    Dim index As Long
'    For Each sheet In ThisWorkbook.Worksheets
    For index = 0 To UBound(sheetNameArr())
        sheetName = sheetNameArr(index)
        '判断页签名是否包含Board Style
        If InStr(sheetName, boardStyleSheetPrefix_) <> 0 Then
            '判断页签名称是否包含Board Style_T%，不包含直接改名为tem_boardStyle，退出For循环
            If InStr(sheetName, temlateBrdStyleShtName_) = 0 Then
                If temShtExistFlage = False Then
                    Call renameCertainSheet(ThisWorkbook, sheetName, "tem_boardStyle")
                    temShtExistFlage = True
                Else
                    Call deleteCertainSheet(ThisWorkbook, sheetName)
                End If
            End If
        End If
    Next
    
    '临时tem_boardStyle页签未生成，再尝试拷贝Board Style_T%开头页签为模板
    If Not WorksheetExists(ThisWorkbook, "tem_boardStyle") Then
'        For Each sheet In ThisWorkbook.Worksheets
        For index = 0 To UBound(sheetNameArr())
            sheetName = sheetNameArr(index)
            'Board Style_T%页签存在直接复制改名为Board Style_1
            If InStr(sheetName, temlateBrdStyleShtName_) <> 0 Then
                ThisWorkbook.Worksheets(sheetName).Copy after:=ThisWorkbook.Worksheets(sheetName)
                ThisWorkbook.ActiveSheet.name = baseBoardStyleSheetName_
                Exit For
            End If
        Next
    Else
        Call renameCertainSheet(ThisWorkbook, "tem_boardStyle", baseBoardStyleSheetName_)
    End If
    
    
    '将调整批注大小的开关关闭，提高后续插入删除行的效率
    Dim baseBoardStyleSheet As Worksheet
    Set baseBoardStyleSheet = ThisWorkbook.Worksheets(baseBoardStyleSheetName_)
    Dim maxRowNumber As Long, maxColumnNumber As Long
    maxRowNumber = baseBoardStyleSheet.range("A1048576").End(xlUp).row
    maxColumnNumber = baseBoardStyleSheet.UsedRange.Columns.count
    Call refreshComment(baseBoardStyleSheet.range(baseBoardStyleSheet.Cells(1, 1), baseBoardStyleSheet.Cells(maxRowNumber, maxColumnNumber)))
End Sub

Public Sub addBoardStyleSheets(ByVal sheetNumber As Long)
    If sheetNumber >= 2 Then
        Call copyBoardStyleSheet(sheetNumber)
        'Call copyBoardStyleDefSheet(sheetNumber, ThisWorkbook.Worksheets("MAPPING DEF"), "A")
        'Call copyBoardStyleDefSheet(sheetNumber, ThisWorkbook.Worksheets("SHEET DEF"), "A")
        'Call copyBoardStyleDefSheet(sheetNumber, ThisWorkbook.Worksheets("CONTROL DEF"), "G")
    End If
End Sub

Private Sub copyBoardStyleSheet(ByVal sheetNumber As Long)
    Dim k As Long
    Dim baseSheet As Worksheet
    Set baseSheet = ThisWorkbook.Worksheets(baseBoardStyleSheetName_)
    For k = 2 To sheetNumber
            baseSheet.Copy after:=baseSheet
            Set baseSheet = ThisWorkbook.ActiveSheet
            baseSheet.name = boardStyleSheetPrefix_ & delimiter_ & k
    Next k
End Sub

Private Sub copyBoardStyleDefSheet(ByVal sheetNumber As Long, ByRef ws As Worksheet, ByRef columnLetter As String)
    Dim maxRowNumber As Long, rowNumber As Long, currentRowNumber As Long
    maxRowNumber = ws.range("a1048576").End(xlUp).row
    currentRowNumber = maxRowNumber + 1
    For rowNumber = 2 To maxRowNumber
        If baseBoardStyleSheetName_ = ws.range(columnLetter & rowNumber).value Then
            Call copyBaseBoardStyleRowAndRename(ws, rowNumber, currentRowNumber, columnLetter, sheetNumber)
            currentRowNumber = currentRowNumber + sheetNumber - 1
        End If
    Next rowNumber
End Sub

Private Sub copyBaseBoardStyleRowAndRename(ByRef ws As Worksheet, ByVal srcRowNumber As Long, _
    ByVal dstRowNumber As Long, ByRef columnLetter As String, ByVal sheetNumber As Long)
    Dim k As Long
    ws.Rows(srcRowNumber).Copy
    ws.Paste Destination:=ws.Rows(dstRowNumber & ":" & dstRowNumber + sheetNumber - 2) '一次性拷贝所有页签的行，再依次重命名
    For k = 2 To sheetNumber
        ws.range(columnLetter & dstRowNumber).value = boardStyleSheetPrefix_ & delimiter_ & k
        dstRowNumber = dstRowNumber + 1
    Next k
End Sub

Private Sub deleteRowsInCertainSheet(ByRef ws As Worksheet, ByRef sheetName, ByRef columnLetter As String)
    Dim rowNumber As Long
    For rowNumber = 2 To ws.UsedRange.Rows.count
        If sheetName = ws.range(columnLetter & rowNumber).value Then
            ws.Rows(rowNumber).Delete
            rowNumber = rowNumber - 1
        End If
    Next rowNumber
End Sub

Private Sub deleteCertainSheet(ByRef wb As Workbook, ByRef sheetName As String)
    Dim ws As Worksheet
    Set ws = wb.Worksheets(sheetName)
    ws.Delete
    'Call deleteRowsInCertainSheet(wb.Worksheets("MAPPING DEF"), sheetName, "A")
    'Call deleteRowsInCertainSheet(wb.Worksheets("SHEET DEF"), sheetName, "A")
    'Call deleteRowsInCertainSheet(wb.Worksheets("CONTROL DEF"), sheetName, "G")
End Sub

Private Sub renameSheetInCertainSheet(ByRef ws As Worksheet, ByRef oldSheetName As String, ByRef newSheetName As String, ByRef columnLetter As String)
    Dim rowNumber As Long
    Dim sheetNameRange As range
    For rowNumber = 2 To ws.UsedRange.Rows.count
        Set sheetNameRange = ws.range(columnLetter & rowNumber)
        If oldSheetName = sheetNameRange.value Then
            sheetNameRange.value = newSheetName
        End If
    Next rowNumber
End Sub

Private Sub renameCertainSheet(ByRef wb As Workbook, ByRef oldSheetName As String, ByRef newSheetName As String)
    wb.Worksheets(oldSheetName).name = newSheetName
    'Call renameSheetInCertainSheet(wb.Worksheets("MAPPING DEF"), oldSheetName, newSheetName, "A")
    'Call renameSheetInCertainSheet(wb.Worksheets("SHEET DEF"), oldSheetName, newSheetName, "A")
    'Call renameSheetInCertainSheet(wb.Worksheets("CONTROL DEF"), oldSheetName, newSheetName, "G")
End Sub

Private Function WorksheetExists(wb As Workbook, strName As String) As Boolean
    Dim str As String
    On Error GoTo worksheetExistsErr
    str = wb.Worksheets(strName).name
    WorksheetExists = True
    Exit Function
worksheetExistsErr:
    WorksheetExists = False
End Function

Private Sub QuickSort(ByRef lngArray() As String)
    Dim iLBound As Long
    Dim iUBound As Long
    Dim iTemp As Variant
    Dim iOuter As Long
    Dim iMax As Long
    
    iLBound = LBound(lngArray)
    iUBound = UBound(lngArray)

    If (iUBound - iLBound) Then
        For iOuter = iLBound To iUBound
            If lngArray(iOuter) > lngArray(iMax) Then iMax = iOuter
        Next iOuter

        iTemp = lngArray(iMax)
        lngArray(iMax) = lngArray(iUBound)
        lngArray(iUBound) = iTemp

        Call InnerQuickSort(lngArray, iLBound, iUBound)
    End If
End Sub

Private Sub InnerQuickSort(ByRef lngArray() As String, ByVal iLeftEnd As Long, ByVal iRightEnd As Long)
    Dim iLeftCur As Long
    Dim iRightCur As Long
    Dim iPivot As Variant
    Dim iTemp As Variant

    If iLeftEnd >= iRightEnd Then Exit Sub

    iLeftCur = iLeftEnd
    iRightCur = iRightEnd + 1
    iPivot = lngArray(iLeftEnd)
    
    Do
        Do
            iLeftCur = iLeftCur + 1
        Loop While lngArray(iLeftCur) < iPivot

        Do
            iRightCur = iRightCur - 1
        Loop While lngArray(iRightCur) > iPivot
        
        If iLeftCur >= iRightCur Then Exit Do
        
        iTemp = lngArray(iLeftCur)
        lngArray(iLeftCur) = lngArray(iRightCur)
        lngArray(iRightCur) = iTemp
    Loop

    lngArray(iLeftEnd) = lngArray(iRightCur)
    lngArray(iRightCur) = iPivot
    Call InnerQuickSort(lngArray, iLeftEnd, iRightCur - 1)
    Call InnerQuickSort(lngArray, iRightCur + 1, iRightEnd)
End Sub
Private Sub Class_Initialize()
    boardStyleSheetPrefix_ = getResByKey("Board Style")
    baseBoardStyleSheetName_ = boardStyleSheetPrefix_ & delimiter_ & "1"
    temlateBrdStyleShtName_ = boardStyleSheetPrefix_ & delimiter_ & "T%"
End Sub

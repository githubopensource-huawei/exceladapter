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

Public Sub keepOneBaseBoardStyleSheet()
    Dim sheet As Worksheet, sheetName As String
    For Each sheet In ThisWorkbook.Worksheets
        sheetName = sheet.name
        If InStr(sheetName, boardStyleSheetPrefix_) <> 0 Then
            If sheetName <> boardStyleSheetPrefix_ And sheetName <> baseBoardStyleSheetName_ Then
                Call deleteCertainSheet(ThisWorkbook, sheetName)
            Else
                Call renameCertainSheet(ThisWorkbook, sheetName, baseBoardStyleSheetName_)
            End If
        End If
    Next sheet
    
    '将调整批注大小的开关关闭，提高后续插入删除行的效率
    Dim baseBoardStyleSheet As Worksheet
    Set baseBoardStyleSheet = ThisWorkbook.Worksheets(baseBoardStyleSheetName_)
    Dim maxRowNumber As Long, maxColumnNumber As Long
    maxRowNumber = baseBoardStyleSheet.range("A65536").End(xlUp).row
    maxColumnNumber = baseBoardStyleSheet.UsedRange.columns.count
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
    maxRowNumber = ws.range("a65536").End(xlUp).row
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
    ws.rows(srcRowNumber).Copy
    ws.Paste Destination:=ws.rows(dstRowNumber & ":" & dstRowNumber + sheetNumber - 2) '一次性拷贝所有页签的行，再依次重命名
    For k = 2 To sheetNumber
        ws.range(columnLetter & dstRowNumber).value = boardStyleSheetPrefix_ & delimiter_ & k
        dstRowNumber = dstRowNumber + 1
    Next k
End Sub

Private Sub deleteRowsInCertainSheet(ByRef ws As Worksheet, ByRef sheetName, ByRef columnLetter As String)
    Dim rowNumber As Long
    For rowNumber = 2 To ws.UsedRange.rows.count
        If sheetName = ws.range(columnLetter & rowNumber).value Then
            ws.rows(rowNumber).Delete
            rowNumber = rowNumber - 1
        End If
    Next rowNumber
End Sub

Private Sub deleteCertainSheet(ByRef wb As Workbook, ByRef sheetName As String)
    wb.Worksheets(sheetName).Delete
    'Call deleteRowsInCertainSheet(wb.Worksheets("MAPPING DEF"), sheetName, "A")
    'Call deleteRowsInCertainSheet(wb.Worksheets("SHEET DEF"), sheetName, "A")
    'Call deleteRowsInCertainSheet(wb.Worksheets("CONTROL DEF"), sheetName, "G")
End Sub

Private Sub renameSheetInCertainSheet(ByRef ws As Worksheet, ByRef oldSheetName As String, ByRef newSheetName As String, ByRef columnLetter As String)
    Dim rowNumber As Long
    Dim sheetNameRange As range
    For rowNumber = 2 To ws.UsedRange.rows.count
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

Private Sub Class_Initialize()
    boardStyleSheetPrefix_ = getResByKey("Board Style")
     baseBoardStyleSheetName_ = boardStyleSheetPrefix_ & delimiter_ & "1"
End Sub

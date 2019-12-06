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
        '�ж�ҳǩ���Ƿ����Board Style
        If InStr(sheetName, boardStyleSheetPrefix_) <> 0 Then
            '�ж�ҳǩ�����Ƿ����Board Style_T%��������ֱ�Ӹ���Ϊtem_boardStyle���˳�Forѭ��
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
    
    '��ʱtem_boardStyleҳǩδ���ɣ��ٳ��Կ���Board Style_T%��ͷҳǩΪģ��
    If Not WorksheetExists(ThisWorkbook, "tem_boardStyle") Then
'        For Each sheet In ThisWorkbook.Worksheets
        For index = 0 To UBound(sheetNameArr())
            sheetName = sheetNameArr(index)
            'Board Style_T%ҳǩ����ֱ�Ӹ��Ƹ���ΪBoard Style_1
            If InStr(sheetName, temlateBrdStyleShtName_) <> 0 Then
                ThisWorkbook.Worksheets(sheetName).Copy after:=ThisWorkbook.Worksheets(sheetName)
                ThisWorkbook.ActiveSheet.name = baseBoardStyleSheetName_
                Exit For
            End If
        Next
    Else
        Call renameCertainSheet(ThisWorkbook, "tem_boardStyle", baseBoardStyleSheetName_)
    End If
    
    
    '��������ע��С�Ŀ��عرգ���ߺ�������ɾ���е�Ч��
    Dim baseBoardStyleSheet As Worksheet
    Set baseBoardStyleSheet = ThisWorkbook.Worksheets(baseBoardStyleSheetName_)
    Dim maxRowNumber As Long, maxColumnNumber As Long
    maxRowNumber = baseBoardStyleSheet.Range("A1048576").End(xlUp).row
    maxColumnNumber = baseBoardStyleSheet.UsedRange.Columns.count
    Call refreshComment(baseBoardStyleSheet.Range(baseBoardStyleSheet.Cells(1, 1), baseBoardStyleSheet.Cells(maxRowNumber, maxColumnNumber)))
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
    
    hyperLintFlag = False
    
    For k = 2 To sheetNumber
        Set baseSheet = ThisWorkbook.Worksheets(baseBoardStyleSheetName_)
        baseSheet.Copy after:=baseSheet
        'Set baseSheet = ThisWorkbook.ActiveSheet
        
        ThisWorkbook.ActiveSheet.name = boardStyleSheetPrefix_ & delimiter_ & CStr(k)
    Next k
    
    Dim shtname As String
    For k = 2 To sheetNumber
        Dim l As Long
        l = sheetNumber - k + 2
        shtname = boardStyleSheetPrefix_ & delimiter_ & CStr(l)
        Set baseSheet = ThisWorkbook.Worksheets(shtname)
        baseSheet.name = boardStyleSheetPrefix_ & CStr(k)
    Next k
    
    For k = 2 To sheetNumber
        shtname = boardStyleSheetPrefix_ & CStr(k)
        Set baseSheet = ThisWorkbook.Worksheets(shtname)
        baseSheet.name = boardStyleSheetPrefix_ & delimiter_ & CStr(k)
    Next k
    
    hyperLintFlag = True
End Sub

Private Sub copyBoardStyleDefSheet(ByVal sheetNumber As Long, ByRef ws As Worksheet, ByRef columnletter As String)
    Dim maxRowNumber As Long, rowNumber As Long, currentRowNumber As Long
    maxRowNumber = ws.Range("a1048576").End(xlUp).row
    currentRowNumber = maxRowNumber + 1
    For rowNumber = 2 To maxRowNumber
        If baseBoardStyleSheetName_ = ws.Range(columnletter & rowNumber).value Then
            Call copyBaseBoardStyleRowAndRename(ws, rowNumber, currentRowNumber, columnletter, sheetNumber)
            currentRowNumber = currentRowNumber + sheetNumber - 1
        End If
    Next rowNumber
End Sub

Private Sub copyBaseBoardStyleRowAndRename(ByRef ws As Worksheet, ByVal srcRowNumber As Long, _
    ByVal dstRowNumber As Long, ByRef columnletter As String, ByVal sheetNumber As Long)
    Dim k As Long
    ws.Rows(srcRowNumber).Copy
    ws.Paste Destination:=ws.Rows(dstRowNumber & ":" & dstRowNumber + sheetNumber - 2) 'һ���Կ�������ҳǩ���У�������������
    For k = 2 To sheetNumber
        ws.Range(columnletter & dstRowNumber).value = boardStyleSheetPrefix_ & delimiter_ & k
        dstRowNumber = dstRowNumber + 1
    Next k
End Sub

Private Sub deleteRowsInCertainSheet(ByRef ws As Worksheet, ByRef sheetName, ByRef columnletter As String)
    Dim rowNumber As Long
    For rowNumber = 2 To ws.UsedRange.Rows.count
        If sheetName = ws.Range(columnletter & rowNumber).value Then
            ws.Rows(rowNumber).Delete
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

Private Sub renameSheetInCertainSheet(ByRef ws As Worksheet, ByRef oldSheetName As String, ByRef newSheetName As String, ByRef columnletter As String)
    Dim rowNumber As Long
    Dim sheetNameRange As Range
    For rowNumber = 2 To ws.UsedRange.Rows.count
        Set sheetNameRange = ws.Range(columnletter & rowNumber)
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
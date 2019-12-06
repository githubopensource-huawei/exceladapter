Attribute VB_Name = "CommentsAssist"
Option Explicit

Dim commentInfos As Collection
Private Const shtNameCol As Integer = 1
Private Const grpNameCol As Integer = 2
Private Const attrNameCol As Integer = 3
Private Const commentCol As Integer = 4

Private Const listShtAttrRow As Integer = 2

Private Const BluePrintSheetColor As Integer = 5

Private Const shtType_List As String = "LIST"
Private Const shtType_Pattern As String = "PATTERN"
Private Const shtType_Common As String = "COMMON"
Private Const shtType_Board As String = "BOARD"
Private Const shtType_IUB As String = "IUB"
Private Const shtType_Main As String = "MAIN"


'==============================interface===============================
Public Sub loadComments()
On Error GoTo ErrorHandler
    Debug.Print "loadComments entered..."
    Dim t As Date
    t = Timer
    
    Dim commentSht As Worksheet
    Set commentSht = ThisWorkbook.Worksheets("Comments")
    
    Set commentInfos = New Collection
    
    Dim shtname As String, shtType As String, grpName As String, attrName As String, commentText As String
    Dim preShtName As String
    
    preShtName = ""
    Dim commentsMgr As CCommentsManager
    Dim rowIdx As Integer
    With commentSht
        For rowIdx = 2 To .Range("a1048576").End(xlUp).row
            shtname = .Cells(rowIdx, shtNameCol)
            grpName = .Cells(rowIdx, grpNameCol)
            attrName = .Cells(rowIdx, attrNameCol)
            commentText = .Cells(rowIdx, commentCol)
            shtType = getShtType(shtname)
            
            If shtname <> preShtName Then
                Set commentsMgr = New CCommentsManager
                commentsMgr.sheetName = shtname
                commentsMgr.sheetType = shtType
                Call commentsMgr.insertComment(grpName, attrName, commentText)
                commentInfos.Add Item:=commentsMgr, key:=shtname
                preShtName = shtname
            Else
                Call commentsMgr.insertComment(grpName, attrName, commentText)
            End If
        Next
    End With
    
    Debug.Print "loadComments exited, time consume: " & Format(Timer - t, "0.00")
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in loadComments, " & Err.Description
End Sub

Public Sub addAllComments()
    On Error GoTo ErrorHandler
    Debug.Print "addAllComments entered..."
    Dim t As Date
    t = Timer
    
    If commentInfos Is Nothing Then loadComments
    
    Dim targetShts As New Collection
    Dim sht As Worksheet
    Dim shtname As String
    For Each sht In ThisWorkbook.Worksheets
        shtname = sht.name
        If sht.Visible = xlSheetVisible And shtname <> getResByKey("Cover") And shtname <> getResByKey("Help") Then targetShts.Add Item:=shtname, key:=shtname
    Next
    
    ProgressBar.Show vbModeless
    Dim percent As Integer
    Dim name As Variant
    Dim idx As Integer
    With ThisWorkbook
        For idx = 1 To targetShts.count
            shtname = CStr(targetShts.Item(idx))
            Call addCommentsBySheet(.Worksheets(shtname))
            percent = idx / targetShts.count * 100
            Call ProgressBar.updateProgress(percent)
        Next
    End With
    
    Debug.Print "addAllComments exited, time consume: " & Format(Timer - t, "0.00")
    
    ThisWorkbook.Save
    
    Exit Sub
ErrorHandler:
    ProgressBar.Hide
    Application.ScreenUpdating = True
    Debug.Print "some exception in addAllComments, " & Err.Description
End Sub


Public Sub addCommentsBySheet(sht As Worksheet)
    On Error GoTo ErrorHandler
    Debug.Print "addCommentsBySheet " & sht.name
    Dim t As Date
    t = Timer
    
    Dim shtType As String
    shtType = getShtType(sht.name)
    
    If commentInfos Is Nothing Then loadComments
    
    If shtType <> shtType_IUB And shtType <> shtType_Board And Not Contains(commentInfos, sht.name) Then Exit Sub
    
    Application.ScreenUpdating = False
    
    Select Case shtType
        Case shtType_List
            Call addCommentsBySheetType_List(sht)
        Case shtType_Common
            Call addCommentsBySheetType_Common(sht)
        Case shtType_Board
            Call addCommentsBySheetType_Board(sht)
        Case shtType_IUB
            Call addCommentsBySheetType_IUB(sht)
    End Select
    
    Application.ScreenUpdating = True
    
    Debug.Print "addCommentsBySheet exited, time consume: " & Format(Timer - t, "0.00")
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    Debug.Print "some exception in addCommentsBySheet, " & Err.Description
End Sub



'===============================implimentation==============================

Private Sub addCommentsBySheetType_List(sht As Worksheet)
    On Error GoTo ErrorHandler
    Dim commentsMgr As CCommentsManager, shtComments As Collection
    Set commentsMgr = commentInfos(sht.name)
    Set shtComments = commentsMgr.comments
    
    Dim attrName As String, grpName As String, key As String, newCommentText As String
    Dim oldComment As comment
    Dim colIdx As Integer
    Dim Target As Range
    
    With sht
        For colIdx = 1 To .Range("XFD" & listShtAttrRow).End(xlToLeft).column
            attrName = Trim(.Cells(listShtAttrRow, colIdx).value)
            grpName = getGrpName_List(sht, colIdx)
            key = commentsMgr.getKey(grpName, attrName)
            If commentsMgr.hasKey(key) Then
                newCommentText = shtComments(key)
                Set Target = .Cells(listShtAttrRow, colIdx)
                Set oldComment = Target.comment
                If oldComment Is Nothing Then
                    Call addNewComment(Target, newCommentText)
                ElseIf oldComment.text = "" Then
                    Target.clearComments
                    Call addNewComment(Target, newCommentText)
                End If
            End If
        Next
    End With
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in addCommentsBySheetType_List, " & Err.Description
End Sub

Private Sub addCommentsBySheetType_Common(sht As Worksheet)
    On Error GoTo ErrorHandler
    Dim shtname As String
    shtname = sht.name
    If isBoardStyleSht(shtname) Then shtname = getResByKey("Board Style")
    
    Dim commentsMgr As CCommentsManager, shtComments As Collection
    Set commentsMgr = commentInfos(shtname)
    Set shtComments = commentsMgr.comments
    
    Dim colIdx As Integer
    Dim attrRowIdx As Integer, grpRowIdx As Integer
    
    Dim attrName As String, grpName As String, key As String, newCommentText As String
    Dim oldComment As comment
    Dim Target As Range

    grpRowIdx = 1
    With sht
        Do
            attrRowIdx = grpRowIdx + 1
            For colIdx = 1 To .Range("XFD" & attrRowIdx).End(xlToLeft).column
                grpName = getGrpName_List(sht, colIdx, grpRowIdx)
                attrName = Trim(.Cells(attrRowIdx, colIdx))
                key = commentsMgr.getKey(grpName, attrName)
                If commentsMgr.hasKey(key) Then
                    newCommentText = shtComments(key)
                    Set Target = .Cells(attrRowIdx, colIdx)
                    Set oldComment = Target.comment
                    If oldComment Is Nothing Then
                        Call addNewComment(Target, newCommentText)
                    ElseIf oldComment.text = "" Then
                        Target.clearComments
                        Call addNewComment(Target, newCommentText)
                    End If
                End If
            Next
            
            grpRowIdx = findNextGrpNameRowEx(sht, attrRowIdx)
        Loop While grpRowIdx <> -1 And grpRowIdx < .Range("a1048576").End(xlUp).row
    End With
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "some exception in addCommentsBySheetType_Common, " & Err.Description
End Sub

Private Sub addCommentsBySheetType_Board(sht As Worksheet)
    On Error GoTo ErrorHandler
    
    Call addCommentsBySheetType_Common(sht)
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in addCommentsBySheetType_Board, " & Err.Description
End Sub

Private Sub addCommentsBySheetType_IUB(sht As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim colIdx As Integer
    Dim attrRowIdx As Integer
    
    Dim shtname As String, attrName As String, grpName As String, key As String, newCommentText As String
    Dim oldComment As comment
    Dim Target As Range

    attrRowIdx = 1
    With sht
        Do
            shtname = .Cells(attrRowIdx, 1)
            Dim commentsMgr As CCommentsManager, shtComments As Collection
            Set commentsMgr = commentInfos(shtname)
            Set shtComments = commentsMgr.comments
    
            For colIdx = 1 To .Range("XFD" & attrRowIdx).End(xlToLeft).column
                attrName = Trim(.Cells(attrRowIdx, colIdx))
                grpName = getGroupNameFromMappingDef(shtname, attrName)
                key = commentsMgr.getKey(grpName, attrName)
                If commentsMgr.hasKey(key) Then
                    newCommentText = shtComments(key)
                    Set Target = .Cells(attrRowIdx, colIdx)
                    Set oldComment = Target.comment
                    If oldComment Is Nothing Then
                        Call addNewComment(Target, newCommentText)
                    ElseIf oldComment.text = "" Then
                        Target.clearComments
                        Call addNewComment(Target, newCommentText)
                    End If
                End If
            Next
            
            attrRowIdx = findNextAttrNameRow(sht, attrRowIdx + 1)
        Loop While attrRowIdx <> -1 And attrRowIdx <= .Range("a1048576").End(xlUp).row
    End With
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in addCommentsBySheetType_IUB, " & Err.Description
End Sub



Private Function getShtType(shtname As String) As String
    On Error GoTo ErrorHandler
    
    If containsASheet(ThisWorkbook, shtname) Then
        If ThisWorkbook.Worksheets(shtname).Tab.colorIndex = BluePrintSheetColor Then
            getShtType = shtType_IUB
            Exit Function
        End If
    End If
    
    If InStr(shtname, getResByKey("Board Style")) = 1 Then
        getShtType = shtType_Board
        Exit Function
    End If
    
    getShtType = shtType_List
    
    Dim shtDef As Worksheet
    Set shtDef = ThisWorkbook.Worksheets("SHEET DEF")
    
    Dim shtType As String
    Dim targetRange As Range
    With shtDef
        Set targetRange = .columns(1).Find(shtname, LookIn:=xlValues, lookat:=xlWhole)
        If targetRange Is Nothing Then
            Debug.Print "cannot find sheet name : " & shtname & " in SHEET DEF"
            Exit Function
        End If
        shtType = UCase(Trim(.Cells(targetRange.row, 2)))
    End With
    
    If shtType = shtType_Main Or shtType = shtType_List Or shtType = shtType_Pattern Then Exit Function
    
    If shtType = shtType_Common Then
        getShtType = shtType_Common
    ElseIf shtType = shtType_Board Then
        getShtType = shtType_Board
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "some exception in getShtType, " & Err.Description
End Function

Private Function isBoardStyleSht(shtname As String) As Boolean
    isBoardStyleSht = False
    If InStr(shtname, getResByKey("Board Style")) = 1 Then isBoardStyleSht = True
End Function

Private Function getGrpName_List(sht As Worksheet, ByVal colIdx As Integer, Optional grpRowIdx As Integer = 1) As String
    On Error GoTo ErrorHandler
    Dim col As Integer
    getGrpName_List = ""
    For col = colIdx To 1 Step -1
        getGrpName_List = sht.Cells(grpRowIdx, col)
        If getGrpName_List <> "" Then Exit Function
    Next
    Exit Function
ErrorHandler:
    Debug.Print "some exception in getGrpName_List, " & Err.Description
End Function

'find next group name row from empty row
Private Function findNextGrpNameRow(sht As Worksheet, ByVal startRow As Long) As Long
    findNextGrpNameRow = -1
    Dim rowIdx As Long
    For rowIdx = startRow To sht.Range("a1048576").End(xlUp).row
        If Not rowIsBlank(sht, rowIdx) Then
            findNextGrpNameRow = rowIdx
            Exit Function
        End If
    Next
End Function

'find next group name row from data row
Private Function findNextGrpNameRowEx(sht As Worksheet, ByVal startRow As Long) As Long
    findNextGrpNameRowEx = -1
    Dim rowIdx As Long
    For rowIdx = startRow To sht.Range("a1048576").End(xlUp).row
        If rowIsBlank(sht, rowIdx) Then
            findNextGrpNameRowEx = findNextGrpNameRow(sht, rowIdx)
            Exit Function
        End If
    Next
End Function

'for IUB sheet
Private Function findNextAttrNameRow(iubSht As Worksheet, ByVal startRow As Long) As Long
    findNextAttrNameRow = -1
    Dim rowIdx As Long
    For rowIdx = startRow To iubSht.Range("a1048576").End(xlUp).row
        If Trim(iubSht.Cells(rowIdx, 1)) <> "" Then
            findNextAttrNameRow = rowIdx
            Exit Function
        End If
    Next
End Function

Private Sub addNewComment(Target As Range, commentText As String)
    On Error GoTo ErrorHandler
    With Target
        With .addComment
            .Visible = False
            .text commentText
        End With
        With .comment.Shape
            .TextFrame.AutoSize = True
            .TextFrame.Characters.Font.Bold = True
        End With
    End With
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in addNewComment, " & Err.Description
End Sub

'An attribute may be in different groups
Private Function getGroupNameFromMappingDef(shtname As String, attrName As String) As String
    On Error GoTo ErrorHandler
    getGroupNameFromMappingDef = ""
    
    Dim MAPPINGDEF As Worksheet
    Set MAPPINGDEF = ThisWorkbook.Worksheets("MAPPING DEF")
    
    Dim shtNameCol As Long
    Dim colNameCol As Long
    Dim grpNameCol As Long
    shtNameCol = 1
    colNameCol = 3
    grpNameCol = 2
    
    Dim tmpShtName As String
    Dim grpName As String
    Dim targetRange As Range
    Dim firstAddr As String
    
    With MAPPINGDEF.columns(colNameCol)
        Set targetRange = .Find(getPlainText(attrName), lookat:=xlWhole, LookIn:=xlValues)
        If Not targetRange Is Nothing Then
            firstAddr = targetRange.address
            Do
                tmpShtName = targetRange.Offset(0, shtNameCol - colNameCol).value
                If tmpShtName = shtname Then
                    getGroupNameFromMappingDef = targetRange.Offset(0, grpNameCol - colNameCol).value
                    Exit Function
                End If
                Set targetRange = .FindNext(targetRange)
            Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
        End If
    End With
    
    Exit Function
ErrorHandler:
    Debug.Print "some exception in getGroupNameFromMappingDef, " & Err.Description
End Function

Private Function getPlainText(ByRef strName As String) As String
    On Error GoTo ErrorHandler
    If InStr(strName, "*") > 0 Then
        getPlainText = Replace(strName, "*", "~*")
        Exit Function
    End If
ErrorHandler:
    getPlainText = strName
End Function

Private Function rowIsBlank(ByRef ws As Worksheet, ByRef RowNumber As Long) As Boolean
    If Application.WorksheetFunction.CountBlank(ws.Range("A" & RowNumber & ":IV" & RowNumber)) = 256 Then
        rowIsBlank = True
    Else
        rowIsBlank = False
    End If
End Function

Private Function containsASheet(ByRef wb As Workbook, ByRef sheetName As String) As Boolean
    On Error GoTo ErrorHandler
    containsASheet = True
    Dim sheet As Worksheet
    Set sheet = wb.Worksheets(sheetName)
    Exit Function
ErrorHandler:
    containsASheet = False
End Function

Attribute VB_Name = "Util"
Option Explicit

Public Const HyperLinkColorIndex = 6
Public Const BluePrintSheetColor = 5
Public Const MaxChosenSiteNum = 202
Private Const SW_SHOWNORMAL = 1
Public HasHistoryData As Boolean

Public Const StandardRowHeight As Double = 13.5

Public Const NeType_GSM = "GSM"
Public Const NeType_UMTS = "UMTS"
Public Const NeType_LTE = "LTE"
Public Const NeType_MRAT = "MRAT"
Public Const NeType_USU = "USU"
Public Const NeType_ICS = "ICS"
Public Const NeType_CBS = "CBS"
Public Const NeType_5G = "NR"
Public Const NeType_DSA = "DSA"

Public Const SheetType_List = "LIST"
Public Const SheetType_Pattern = "PATTERN"
Public Const SheetType_Main = "MAIN"
Public Const SheetType_Common = "COMMON"
Public Const SheetType_Board = "BOARD"
Public Const SheetType_Iub = "IUB"

Public Function getNeType() As String
    On Error Resume Next
    Dim cover As String
    Dim key As String
    Dim reValue As String
    
    cover = getResByKey("Cover")
    key = ThisWorkbook.Worksheets(cover).Cells(2, 2).value
    reValue = getResByKey(key)
    
    Select Case reValue
        Case "GSM"
            getNeType = NeType_GSM
        Case "UMTS"
            getNeType = NeType_UMTS
        Case "LTE"
            getNeType = NeType_LTE
        Case "MRAT"
            getNeType = NeType_MRAT
        Case "USU"
            getNeType = NeType_USU
        Case "ICS"
            getNeType = NeType_ICS
        Case "CBS"
            getNeType = NeType_CBS
        Case "NR"
            getNeType = NeType_5G
        Case "DSA"
            getNeType = NeType_DSA
        Case Else
            getNeType = NeType_MRAT
    End Select
End Function

Public Function GetCell(shtX As Worksheet, ByVal r, ByVal C)
    GetCell = shtX.Cells(r, C)
End Function

Public Function max(ByVal a, ByVal b)
    max = IIf(a > b, a, b)
End Function

Public Function min(ByVal a, ByVal b)
    min = IIf(a < b, a, b)
End Function

'获得垂直合并的组名代码
Public Function getVerticalGroupName(ByRef ws As Worksheet, ByRef rowNumber As Long, ByRef columnLetter As String, _
    ByRef groupStartRow As Long, ByRef groupEndRow As Long) As String
    Dim cellValue As String
    cellValue = ws.Range(columnLetter & rowNumber).value
    Dim k As Long
    If cellValue = "" Then
        For k = rowNumber To 1 Step -1
            cellValue = ws.Range(columnLetter & k).value
            If cellValue <> "" Then
                getVerticalGroupName = cellValue
                groupStartRow = k
                groupEndRow = getEndRowNumer(ws, columnLetter, k)
                Exit Function
            End If
        Next k
    Else
        getVerticalGroupName = cellValue
        groupStartRow = rowNumber
        groupEndRow = getEndRowNumer(ws, columnLetter, k)
    End If
End Function

'获得垂直合并的组结束行数代码
Public Function getEndRowNumer(ByRef ws As Worksheet, ByRef columnLetter As String, ByRef startRowNumber As Long) As Long
    Dim k As Long
    Dim maxRowNumber As Long
    maxRowNumber = ws.UsedRange.rows.count 'ws.Range("A2").End(xlToRight).Column
    For k = startRowNumber + 1 To maxRowNumber
        If ws.Range(columnLetter & k).value <> "" Then
            getEndRowNumer = k - 1
            Exit Function
        End If
    Next k
    getEndRowNumer = maxRowNumber
End Function

Public Function GetMainSheetName() As String
    On Error Resume Next
    If innerPositionMgr Is Nothing Then loadInnerPositions
    
    Dim targetRange As Range
    Set targetRange = ThisWorkbook.Worksheets("SHEET DEF").columns(innerPositionMgr.sheetDef_shtTypeColNo).Find("MAIN", lookat:=xlWhole, LookIn:=xlValues)
    If Not targetRange Is Nothing Then GetMainSheetName = targetRange.Offset(0, innerPositionMgr.sheetDef_shtNameColNo - innerPositionMgr.sheetDef_shtTypeColNo).value
End Function

Function GetCommonSheetName() As String
    On Error Resume Next
    Dim name As String
    Dim rowNum As Long
    Dim sheetDef As Worksheet
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
    If innerPositionMgr Is Nothing Then loadInnerPositions
    
    For rowNum = 1 To sheetDef.Range("a65536").End(xlUp).row
       If sheetDef.Cells(rowNum, innerPositionMgr.sheetDef_shtTypeColNo).value = "COMMON" Then
           name = sheetDef.Cells(rowNum, innerPositionMgr.sheetDef_shtNameColNo).value
           Exit For
       End If
    Next
    GetCommonSheetName = name
End Function

'从普通页取得Group name
Public Function get_GroupName(sheetName As String, column As Long) As String
        Dim index As Long
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(sheetName)
        For index = column To 1 Step -1
            'If Not isEmpty(ws.Cells(1, index).value) Then
            If ws.Cells(1, index).value <> "" Then
                get_GroupName = ws.Cells(1, index).value
                Exit Function
            End If
        Next
        get_GroupName = ""
End Function

'从普通页取得Colum name
Public Function get_ColumnName(ByVal sheetName As String, column As Long) As String
        Dim index As Long
        get_ColumnName = ThisWorkbook.Worksheets(sheetName).Cells(2, column)
End Function

Public Function isNum(contedType As String) As Boolean
    If contedType = "Integer" Or contedType = "UInteger" Then
        isNum = True
    Else
        isNum = False
    End If
End Function

'***************************************************************
'*****Interface called by JNI
'***************************************************************
Public Sub clearXLGray()
    Dim index, cloumIndex, commIndex, commCloumIndex As Long
    Dim worksh As Worksheet, sheetDef As New Worksheet
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
    If innerPositionMgr Is Nothing Then loadInnerPositions
    
    For index = 2 To sheetDef.Range("a65536").End(xlUp).row
        Set worksh = ThisWorkbook.Worksheets(sheetDef.Cells(index, innerPositionMgr.sheetDef_shtNameColNo).value)
        If sheetDef.Cells(index, innerPositionMgr.sheetDef_shtTypeColNo) = "COMMON" Then
            For commIndex = 1 To getSheetUsedRows(worksh)
                For commCloumIndex = 1 To worksh.Range("IV" + CStr(commIndex)).End(xlToLeft).column
                    If worksh.Cells(commIndex, commCloumIndex).Interior.colorIndex = 16 And _
                         worksh.Cells(commIndex, commCloumIndex).Interior.Pattern = xlGray16 Then
                            worksh.Cells(commIndex, commCloumIndex).Interior.colorIndex = xlNone
                            worksh.Cells(commIndex, commCloumIndex).Interior.Pattern = xlNone
                    End If
                Next
            Next
        ElseIf "Pattern" = sheetDef.Cells(index, innerPositionMgr.sheetDef_shtTypeColNo) Then

        Else
            For cloumIndex = 1 To worksh.Range("IV" + CStr(3)).End(xlToLeft).column
                If worksh.Cells(3, cloumIndex).Interior.colorIndex = 16 And _
                     worksh.Cells(3, cloumIndex).Interior.Pattern = xlGray16 Then
                        worksh.Cells(3, cloumIndex).Interior.colorIndex = xlNone
                        worksh.Cells(3, cloumIndex).Interior.Pattern = xlNone
                End If
            Next
        End If
    Next
    Application.DisplayAlerts = False
    ThisWorkbook.Save
End Sub

Public Function isPatternSheet(sheetName As String) As Boolean
    Dim m_rowNum As Long
    Dim sheetDef As Worksheet
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
    If innerPositionMgr Is Nothing Then loadInnerPositions
    
    For m_rowNum = 2 To sheetDef.Range("a65536").End(xlUp).row
        If sheetName = sheetDef.Cells(m_rowNum, innerPositionMgr.sheetDef_shtNameColNo).value Then
            If sheetDef.Cells(m_rowNum, innerPositionMgr.sheetDef_shtTypeColNo).value = "Pattern" Then
                isPatternSheet = True
            Else
                isPatternSheet = False
            End If
            Exit For
        End If
    Next
End Function

Public Sub setHyperlinkRangeFont(ByRef certainRange As Range)
    With certainRange.Font
        .name = "Arial"
        .Size = 10
    End With
End Sub

'某行是否为空代码
Public Function rowIsBlank(ByRef ws As Worksheet, ByRef rowNumber As Long) As Boolean
    If Application.WorksheetFunction.CountBlank(ws.Range("A" & rowNumber & ":IV" & rowNumber)) = 256 Then
        rowIsBlank = True
    Else
        rowIsBlank = False
    End If
End Function

Function is_Site(columnName As String) As Boolean
    is_Site = False
    If columnName = getResByKey("*NodeBName") Or columnName = getResByKey("*BTSName") Or _
        columnName = getResByKey("*Name") Or columnName = getResByKey("*eNodeBName") Or _
        columnName = getResByKey("*USUName") Or columnName = getResByKey("USU3900NAME") Or columnName = getResByKey("*DSAName") Or _
        columnName = getResByKey("USU3910NAME") Or columnName = getResByKey("*NBBSName") Or columnName = getResByKey("*gNodeBName") Or _
        columnName = getResByKey("*ICSName") Or columnName = getResByKey("*eLTEName") Or columnName = getResByKey("*RFAName") Then
        is_Site = True
    End If
End Function

Function is_Controller(columnName As String) As Boolean
    is_Controller = False
    If (columnName = getResByKey("*RNCName") Or columnName = getResByKey("*BSCName")) Then
        is_Controller = True
    End If
End Function

Function isOperationWs(ByRef ws As Worksheet) As Boolean
    isOperationWs = False

    If operationColNum(ws) = -1 Then Exit Function
    
    isOperationWs = True
End Function

'将比较字符串整形
Function GetDesStr(srcStr As String) As String
    GetDesStr = UCase(Trim(srcStr))
End Function

Public Function containsASheet(shtName As String) As Boolean
    On Error GoTo ErrorHandle
    Dim tmp As Worksheet
    Set tmp = ThisWorkbook.Worksheets(shtName)
    containsASheet = True
    Exit Function
ErrorHandle:
    containsASheet = False
End Function

Public Function IsBluePrintSheetName(sheetName As String) As Boolean
    IsBluePrintSheetName = (Sheets(sheetName).Tab.colorIndex = BluePrintSheetColor)
End Function

Public Function existsASheet(shtName As String) As Boolean
On Error GoTo ErrorHandler:
    existsASheet = True
    Dim sht As Worksheet
    Set sht = ThisWorkbook.Worksheets(shtName)
    Exit Function
ErrorHandler:
    existsASheet = False
End Function

Public Function isMultiVersionWb() As Boolean
    isMultiVersionWb = False
    If existsASheet(getResByKey("ModelDiffSht")) Then
        isMultiVersionWb = True
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

Public Function collectionJoin(coll As Collection, Optional delimiter As String = ",") As String
    On Error GoTo ErrorHandler
    collectionJoin = ""
    Dim del As String
    del = ""
    
    Dim it As Variant
    For Each it In coll
        collectionJoin = collectionJoin & del & CStr(it)
        del = delimiter
    Next
    Exit Function
ErrorHandler:
    Debug.Print "some exception in collectionJoin, " & Err.Description
    collectionJoin = ""
End Function

Public Function isAttrRow_IUB(sht As Worksheet, ByVal rowIdx As Integer) As Boolean
    On Error GoTo ErrorHandler
    isAttrRow_IUB = True
    If sht.Cells(rowIdx, 1) <> "" Then Exit Function
    
    isAttrRow_IUB = False
    
    Exit Function
ErrorHandler:
    Debug.Print "some exception in isAttrRow_IUB, " & Err.Description
    isAttrRow_IUB = False
End Function

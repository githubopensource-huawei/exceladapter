Attribute VB_Name = "Util"
Option Explicit

Global Const g_strInvalidChar4Sql As String = "'"
Global Const g_strInvalidChar4PureFileName  As String = "\/:*?<>|"""
Global Const g_strInvalidChar4Path          As String = "/*?<>|"""

Private Const SW_SHOWNORMAL = 1
Public HasHistoryData As Boolean

Public Const HyperLinkColorIndex = 6
Public Const BluePrintSheetColor = 5
Public Const MaxChosenSiteNum = 202

Public Const NeType_GSM = "GSM"
Public Const NeType_UMTS = "UMTS"
Public Const NeType_LTE = "LTE"
Public Const NeType_MRAT = "MRAT"
Public Const NeType_USU = "USU"
Public Const NeType_ICS = "ICS"
Public Const NeType_CBS = "CBS"
Public Const NeType_5G = "NR"

Public Const SheetType_List = "LIST"
Public Const SheetType_Pattern = "PATTERN"
Public Const SheetType_Main = "MAIN"
Public Const SheetType_Common = "COMMON"
Public Const SheetType_Board = "BOARD"
Public Const SheetType_Iub = "IUB"

Global Const StartRow_Name As String = "StartRow"
Global Const EndRow_Name As String = "EndRow"
Global Const BaseSheetName_Name As String = "Base Sheet Name"

Public bIsEng As Boolean  '用于控制设置中英文

Public FILE_TYPE As String '0- Summary 1-Bulk

Sub setFileType(fileType As String)
        FILE_TYPE = fileType
End Sub




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
        Case Else
            getNeType = ""
    End Select
End Function

Public Function GetCell(shtX As Worksheet, ByVal r, ByVal C)
    GetCell = shtX.Cells(r, C)
End Function

'写入：True；未写入：False
Public Function SetCell(shtX As Worksheet, ByVal r, ByVal C, ByVal strCellVal) As Boolean
    Dim strOld As String
    
    strOld = GetCell(shtX, r, C)
    If strOld <> CStr(strCellVal) Then
        shtX.Cells(r, C) = strCellVal
        SetCell = True
    End If
End Function

Public Function MakeRange(shtX As Worksheet, r0, c0, Optional r1, Optional c1) As range
    If IsMissing(r1) Then r1 = r0
    If IsMissing(c1) Then c1 = c0
    Set MakeRange = shtX.range(shtX.Cells(r0, CellCol2Str(c0)), shtX.Cells(r1, CellCol2Str(c1)))
End Function

Public Function max(ByVal a, ByVal b)
    max = IIf(a > b, a, b)
End Function

Public Function min(ByVal a, ByVal b)
    min = IIf(a < b, a, b)
End Function

Public Sub SetColWidth(shtX As Worksheet, ByVal col, ByVal lWidth)
    MakeRange(shtX, 1, col).EntireColumn.ColumnWidth = lWidth
End Sub

Public Sub AssertEx(Optional bCondition As Boolean = False)
    Debug.Assert (bCondition)
End Sub

'判断x是否介于[a, b]之间
Public Function Between(x, a, b) As Boolean
    Between = ((a <= x) And (x <= b))
End Function

'输入：1~256
Function CellCol2Str(ByVal C) As String
    Dim n0 As String
    Dim n1 As String
    
    If Not IsNumeric(C) Then
        CellCol2Str = UCase(C)
        Exit Function
    End If
    
    C = C - 1
    AssertEx Between(C, 0, 255)
    n0 = Chr((C Mod 26) + Asc("A"))
    C = C \ 26
    If C > 0 Then n1 = Chr(C + Asc("A") - 1)
    
    CellCol2Str = n1 & n0
End Function

Sub DisplayMessageOnStatusbar()
    Application.DisplayStatusBar = True '显示状态栏
    Application.StatusBar = "Running,please wait......" '状态栏显示信息
    Application.Cursor = xlWait
End Sub

Public Sub ReturnStatusbaring()
    Application.StatusBar = "Ready." '状态栏恢复正常
End Sub

'装载用于添加「Base Station Transport Data」页「*Site Template」列侯选值的窗体
Sub addTemplate()
    Load TemplateForm
    TemplateForm.Show
End Sub
Sub addIPRoute()
    Load IPRouteForm
    IPRouteForm.Show
End Sub

'获得垂直合并的组名代码
Public Function getVerticalGroupName(ByRef ws As Worksheet, ByRef rowNumber As Long, ByRef columnLetter As String, _
    ByRef groupStartRow As Long, ByRef groupEndRow As Long) As String
    Dim cellValue As String
    cellValue = ws.range(columnLetter & rowNumber).value
    Dim k As Long
    If cellValue = "" Then
        For k = rowNumber To 1 Step -1
            cellValue = ws.range(columnLetter & k).value
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
        If ws.range(columnLetter & k).value <> "" Then
            getEndRowNumer = k - 1
            Exit Function
        End If
    Next k
    getEndRowNumer = maxRowNumber
End Function

'从指定sheet页的指定行，查找指定列，返回列号,通过attrName和MocName获取到
Public Function getColNum(sheetName As String, recordRow As Long, attrName As String, mocName As String) As Long
    On Error GoTo ErrorHandler
    Dim colName As String
    Dim grpName As String
    
    Dim flag As Boolean
    Dim mappingDef As Worksheet
    Dim ws As Worksheet
    If innerPositionMgr Is Nothing Then loadInnerPositions

    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    flag = False
    getColNum = -1
    
    Dim firstAddr As String
    Dim targetRange As range
    With mappingDef
        Set targetRange = .columns(innerPositionMgr.mappingDef_attrNameColNo).Find(attrName, lookat:=xlWhole, LookIn:=xlValues)
        If Not targetRange Is Nothing Then
            firstAddr = targetRange.address
            Do
                If UCase(targetRange.Offset(0, innerPositionMgr.mappingDef_mocNameColNo - innerPositionMgr.mappingDef_attrNameColNo).value) = UCase(mocName) _
                    And targetRange.Offset(0, innerPositionMgr.mappingDef_shtNameColNo - innerPositionMgr.mappingDef_attrNameColNo).value = sheetName Then
                        colName = .Cells(targetRange.row, innerPositionMgr.mappingDef_colNameColNo).value
                        grpName = .Cells(targetRange.row, innerPositionMgr.mappingDef_grpNameColNo).value
                        flag = True
                        Exit Do
                End If
                Set targetRange = .columns(innerPositionMgr.mappingDef_attrNameColNo).FindNext(targetRange)
            Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
        End If
    End With
    
    Set targetRange = Nothing
    firstAddr = ""
    If flag = True Then
        Set ws = ThisWorkbook.Worksheets(sheetName)
        With ws.rows(recordRow)
            Set targetRange = .Find(colName, lookat:=xlWhole, LookIn:=xlValues)
            If Not targetRange Is Nothing Then
                firstAddr = targetRange.address
                Do
                    If get_GroupName(sheetName, targetRange.column) = grpName Then
                        getColNum = targetRange.column
                        Exit Do
                    End If
                    Set targetRange = .FindNext(targetRange)
                Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
            End If
        End With
    End If
    
    Exit Function
ErrorHandler:
    Debug.Print "some exception in getColNum, " & Err.Description
End Function

Public Function GetMainSheetName() As String
    On Error Resume Next
    If innerPositionMgr Is Nothing Then loadInnerPositions
    
    Dim targetRange As range
    Set targetRange = ThisWorkbook.Worksheets("SHEET DEF").columns(innerPositionMgr.sheetDef_shtTypeColNo).Find("MAIN", lookat:=xlWhole, LookIn:=xlValues)
    If Not targetRange Is Nothing Then GetMainSheetName = targetRange.Offset(0, innerPositionMgr.sheetDef_shtNameColNo - innerPositionMgr.sheetDef_shtTypeColNo).value
End Function

Function GetCommonSheetName() As String
         On Error Resume Next
        Dim name As String
        Dim rowNum As Long
        Dim sheetDef As Worksheet
        Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
        
        For rowNum = 1 To sheetDef.range("a65536").End(xlUp).row
            If sheetDef.Cells(rowNum, 2).value = "COMMON" Then
                name = sheetDef.Cells(rowNum, 1).value
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

Public Sub EndDisplayMessageOnStatusbar()
    Application.Cursor = xlDefault
    Application.StatusBar = "Finished."  '状态栏显示信息
End Sub

Public Function isNum(contedType As String) As Boolean
    If contedType = "Integer" Or contedType = "UInteger" Then
        isNum = True
    Else
        isNum = False
    End If
End Function

Public Sub clearXLGray()
    Dim index, cloumIndex, commIndex, commCloumIndex As Long
    Dim worksh, sheetDef As New Worksheet
    If innerPositionMgr Is Nothing Then loadInnerPositions
    
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
    For index = 2 To sheetDef.range("a65536").End(xlUp).row
        Set worksh = ThisWorkbook.Worksheets(sheetDef.Cells(index, innerPositionMgr.sheetDef_shtNameColNo).value)
        If sheetDef.Cells(index, innerPositionMgr.sheetDef_shtTypeColNo) = "COMMON" Then
            For commIndex = 1 To worksh.range("a65536").End(xlUp).row
                For commCloumIndex = 1 To worksh.range("IV" + CStr(commIndex)).End(xlToLeft).column
                    If worksh.Cells(commIndex, commCloumIndex).Interior.colorIndex = 16 And _
                         worksh.Cells(commIndex, commCloumIndex).Interior.Pattern = xlGray16 Then
                            worksh.Cells(commIndex, commCloumIndex).Interior.colorIndex = xlNone
                            worksh.Cells(commIndex, commCloumIndex).Interior.Pattern = xlNone
                            'worksh.Cells(commIndex, commCloumIndex).Validation.ShowInput = True
                    End If
                Next
            Next
        ElseIf "Pattern" = sheetDef.Cells(index, innerPositionMgr.sheetDef_shtTypeColNo) Then
                
        Else
            For cloumIndex = 1 To worksh.range("IV" + CStr(3)).End(xlToLeft).column
                If worksh.Cells(3, cloumIndex).Interior.colorIndex = 16 And _
                    worksh.Cells(3, cloumIndex).Interior.Pattern = xlGray16 Then
                    worksh.Cells(3, cloumIndex).Interior.colorIndex = xlNone
                    worksh.Cells(3, cloumIndex).Interior.Pattern = xlNone
                    'worksh.Cells(3, cloumIndex).Validation.ShowInput = True
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
    If innerPositionMgr Is Nothing Then loadInnerPositions
    
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
    For m_rowNum = 2 To sheetDef.range("a65536").End(xlUp).row
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

Sub clearStyles()
        Dim s As Style
        For Each s In ThisWorkbook.Styles
            If Not s.BuiltIn Then
                'Debug.Print s.Name
                Debug.Print s.name
                s.Delete '可以用来删除非内置样式
            End If
        Next
End Sub

'包含某个页签代码
Public Function containsASheet(ByRef wb As Workbook, ByRef sheetName As String) As Boolean
    On Error GoTo ErrorHandler
    containsASheet = True
    Dim sheet As Worksheet
    Set sheet = wb.Worksheets(sheetName)
    Exit Function
ErrorHandler:
    containsASheet = False
End Function

Public Function findCertainValRowNumberByTwoKeys(ByRef ws As Worksheet, ByVal columnLetter1 As String, ByRef cellVal1 As String, _
    ByVal columnLetter2 As String, ByRef cellVal2 As String, Optional ByVal startRow As Long = 1)
    
    Dim currentCellVal1 As String
    Dim currentCellVal2 As String
    Dim maxRowNumber As Long, k As Long
    maxRowNumber = ws.UsedRange.rows.count
    findCertainValRowNumberByTwoKeys = -1
    For k = startRow To maxRowNumber
        currentCellVal1 = ws.range(columnLetter1 & k).value
        currentCellVal2 = ws.range(columnLetter2 & k).value
        If currentCellVal1 = cellVal1 And currentCellVal2 = cellVal2 Then
            findCertainValRowNumberByTwoKeys = k
            Exit For
        End If
    Next
End Function

Public Sub setHyperlinkRangeFont(ByRef certainRange As range)
    With certainRange.Font
        .name = "Arial"
        .Size = 10
    End With
End Sub

'某行是否为空代码
Public Function rowIsBlank(ByRef ws As Worksheet, ByRef rowNumber As Long) As Boolean
    If Application.WorksheetFunction.CountBlank(ws.range("A" & rowNumber & ":IV" & rowNumber)) = 256 Then
        rowIsBlank = True
    Else
        rowIsBlank = False
    End If
End Function

Sub changeAlerts(ByRef flag As Boolean)
    Application.EnableEvents = flag
    Application.DisplayAlerts = flag
    Application.ScreenUpdating = flag
End Sub

Public Sub eraseLastChar(ByRef str As String)
    If str <> "" Then str = Left(str, Len(str) - 1)
End Sub

Public Function getNBIOTFlag() As Boolean
    On Error GoTo ErrorHandler
    Dim mocColNum As Long, headRange As range, prbRange As range
    Set headRange = Worksheets("MAPPING DEF").range("A1:X1").Find(what:="Attribute Name", lookat:=xlWhole)
    If Not headRange Is Nothing Then
         mocColNum = headRange.column
    End If

    Dim rowEnd As Long, colstr As String
    colstr = getColStr(mocColNum)
    Set prbRange = Worksheets("MAPPING DEF").range(colstr & ":" & colstr).Find(what:="NbCellFlag", lookat:=xlWhole)
    If Not prbRange Is Nothing Then
        getNBIOTFlag = True
        Exit Function
    End If

ErrorHandler:
getNBIOTFlag = False
End Function

Function is_Site(columnName As String) As Boolean
    is_Site = False
    If (columnName = getResByKey("*NODEB_NAME") Or columnName = getResByKey("*BTS_NAME") Or _
        columnName = getResByKey("*BASESTATION_NAME") Or columnName = getResByKey("*ENODEB_NAME") Or columnName = getResByKey("*USU_NAME") Or _
        columnName = getResByKey("*NBBSName") Or columnName = getResByKey("*ICSName")) Then
        is_Site = True
    End If
End Function

Public Function getSheetDefNameColNum(ByRef titleName As String) As Long
    Dim sheetDef As Worksheet
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
    Dim maxColNum As Long, index As Long, lastColNum As Long
    maxColNum = sheetDef.range("IV1").End(xlToLeft).column
    For index = 1 To maxColNum
         If UCase(sheetDef.Cells(1, index).value) = UCase(titleName) Then
             getSheetDefNameColNum = index
             Exit Function
         End If
    Next
    getSheetDefNameColNum = -1
    lastColNum = -1
    If titleName = BaseSheetName_Name And lastColNum = -1 Then
        getSheetDefNameColNum = 6
    End If
    Exit Function
ErrorHandler:
    getSheetDefNameColNum = -1
End Function

Public Function getUsedColumnCount(ByRef sheet As Worksheet, Optional ByRef attrNamerowIndex As Long = 2)
    getUsedColumnCount = sheet.range("IV" & attrNamerowIndex).End(xlToLeft).column
End Function

Public Function getUsedRowCount(ByRef sheet As Worksheet, Optional ByRef attrNamerowIndex As Long = 2)
    Dim colCount As Long
    colCount = getUsedColumnCount(sheet, attrNamerowIndex)
    getUsedRowCount = sheet.columns("A:" & getColStr(colCount)).Find(what:="*", searchorder:=xlByRows, searchdirection:=xlPrevious).row
End Function

Public Function findAttrName(ByRef attrName As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Len(Trim(attrName)) < 1 Then
        findAttrName = False
        Exit Function
    End If
    
    
    Dim mocColNum As Long, headRange As range, prbRange As range
    Set headRange = Worksheets("MAPPING DEF").range("A1:X1").Find(what:="Column Name", lookat:=xlWhole)
    If Not headRange Is Nothing Then
         mocColNum = headRange.column
    End If
    
    Dim rowEnd As Long, colstr As String
    colstr = getColStr(mocColNum)
    Set prbRange = Worksheets("MAPPING DEF").range(colstr & ":" & colstr).Find(what:=attrName, lookat:=xlWhole)
    If Not prbRange Is Nothing Then
        findAttrName = True
        Exit Function
    End If
ErrorHandler:
findAttrName = False
End Function

Public Function findGroupName(ByRef groupName As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Len(Trim(groupName)) < 1 Then
        findGroupName = False
        Exit Function
    End If
    
    Dim mocColNum As Long, headRange As range, prbRange As range
    Set headRange = Worksheets("MAPPING DEF").range("A1:X1").Find(what:="Group Name", lookat:=xlWhole)
    If Not headRange Is Nothing Then
         mocColNum = headRange.column
    End If
    
    Dim rowEnd As Long, colstr As String
    colstr = getColStr(mocColNum)
    Set prbRange = Worksheets("MAPPING DEF").range(colstr & ":" & colstr).Find(what:=groupName, lookat:=xlWhole)
    If Not prbRange Is Nothing Then
        findGroupName = True
        Exit Function
    End If
ErrorHandler:
findGroupName = False
End Function



Public Function replaceGenel(ByRef strName As String) As String
 On Error GoTo ErrorHandler
 If InStr(strName, "*") > 0 Then
  replaceGenel = Replace(strName, "*", "~*")
  Exit Function
 End If
ErrorHandler:
replaceGenel = strName
End Function


Public Function attrNameColNumInSpecialDef(ByRef sheet As Worksheet, ByRef attrName As String, Optional ByRef attrNamerowIndex As Long = 1) As Long
    On Error GoTo ErrorHandler
    attrNameColNumInSpecialDef = -1
    
    Dim targetRange As range
    Set targetRange = sheet.rows(attrNamerowIndex).Find(Trim(attrName), LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then attrNameColNumInSpecialDef = targetRange.column
    Exit Function
ErrorHandler:
    attrNameColNumInSpecialDef = -1
End Function

Public Function special() As Boolean
On Error GoTo ErrorHandler
special = False
Dim cover As String
Dim key As String
cover = getResByKey("Cover")
key = ThisWorkbook.Worksheets(cover).Cells(2, 2).value
If InStr(key, "CloudRANCU_ENB") > 0 Then
   special = True
End If

ErrorHandler:
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

Public Function IsSheetExist(sheetName As String) As Boolean
    Dim SheetNum, SheetCount As Long 'SheetCount每个原始数据文件的Sheet页总数
    SheetCount = ActiveWorkbook.Worksheets.count   '共有几个Sheet页
    For SheetNum = 1 To SheetCount
        If UCase(Worksheets(SheetNum).name) = UCase(sheetName) Then
            IsSheetExist = True
            Exit Function
        End If
    Next SheetNum
    IsSheetExist = False
End Function

Public Function GetBluePrintSheetName() As String '当前只支持一个
    GetBluePrintSheetName = ""
    
    Dim SheetNum, SheetCount As Long
    SheetCount = ActiveWorkbook.Worksheets.count
    For SheetNum = 1 To SheetCount
        If Worksheets(SheetNum).Tab.colorIndex = BluePrintSheetColor Then
            GetBluePrintSheetName = Worksheets(SheetNum).name
            Exit Function
        End If
    Next SheetNum
End Function

Function is_Controller(columnName As String) As Boolean
    is_Controller = False
    If (columnName = getResByKey("*RNCName") Or columnName = getResByKey("*BSCName")) Then
        is_Controller = True
    End If
End Function

Public Sub InitTemplateVersion()
    bIsEng = getResByKey("Cover") = "Cover"
End Sub

Public Function isMultiVersionWb() As Boolean
    isMultiVersionWb = False
    If existsASheet(getResByKey("ModelDiffSht")) Then
        isMultiVersionWb = True
    End If
End Function

Function isOperationWs(ByRef ws As Worksheet) As Boolean
    isOperationWs = False

    If operationColNum(ws) = -1 Then Exit Function
    
    isOperationWs = True
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

Public Function IsBluePrintSheetName(sheetName As String) As Boolean
    IsBluePrintSheetName = (Sheets(sheetName).Tab.colorIndex = BluePrintSheetColor)
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

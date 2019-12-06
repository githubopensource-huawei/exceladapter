Attribute VB_Name = "IubFormatReport"
Option Explicit

Public bIsEng As Boolean  '用于控制设置中英文
'==================================================================================================
'生成Iub形式代码
'==================================================================================================
Public Const HyperLinkColorIndex = 6
Public Const BluePrintSheetColor = 5
Public Const MaxChosenSiteNum = 2002

Public Const g_NodeBBeginRow = 3
Public Const g_TitleRowNo As Long = 2
Public NodeArray(MaxChosenSiteNum) As String
Public RateArray(MaxChosenSiteNum) As String
Public g_NodeIndex As Long

Public ChosenSiteArray(MaxChosenSiteNum) As String

Public FILE_TYPE As String '0- Summary 1-Bulk

Sub setFileType(fileType As String)
        FILE_TYPE = fileType
End Sub
Function is_SiteorController(columnName As String) As Boolean
    is_SiteorController = False
    If (columnName = "*NodeB Name" Or columnName = "*NodeB名称" Or _
        columnName = "*BTS Name" Or columnName = "*BTS名称" Or _
        columnName = "*Name" Or columnName = "*基站名称" Or _
        columnName = "*eNodeB Name" Or columnName = "*eNodeB名称" Or _
        columnName = "*USU Name" Or columnName = "*USU名称" Or _
        is_Controller(columnName)) Then
        is_SiteorController = True
    End If
End Function

Function is_Controller(columnName As String) As Boolean
    is_Controller = False
    If (columnName = "*RNC Name" Or columnName = "*RNC名称" Or _
        columnName = "*BSC Name" Or columnName = "*BSC名称") Then
        is_Controller = True
    End If
End Function

'[Summary转类似原IUB表格功能]将List Sheet数据转到NodeB Sheet页
Public Sub ConvertList(ByVal strSheetName As String, ByVal startRow As Long, ByVal endRow As Long, ByVal IsEmptyDataSheet As Boolean)
    Dim vListSht As Worksheet
    Dim vName As String
    Dim vMocEndRowNum As Long

    If (Not IsSheetExist(strSheetName)) Then
        Exit Sub
    End If
    
    Set vListSht = Sheets(strSheetName)
    vName = GetCell(vListSht, g_TitleRowNo, 1)
    '只处理NodeB&BTS对象
    If (Not is_SiteorController(vName)) Or (isTrasnPortSheet(strSheetName) And FILE_TYPE = "0") Then
        Exit Sub
    End If
    
    Call FormatFirst(vListSht, strSheetName, startRow, endRow)
    
    vMocEndRowNum = vListSht.Range("a65536").End(xlUp).row
    If (vMocEndRowNum <= g_TitleRowNo) Then
        vMocEndRowNum = g_TitleRowNo + 1
    End If

    If is_Controller(vName) Then
        Call copyControllerTitle(vListSht, strSheetName, startRow)
'        Call CopyControllerData(vListSht, strSheetName, vMocEndRowNum, startRow, IsEmptyDataSheet)
    Else
        Call CopySiteTitle(vListSht, strSheetName, startRow)
'        Call CopySiteData(vListSht, strSheetName, vMocEndRowNum, startRow, IsEmptyDataSheet)
    End If
    
    Call setCopiedRangeFont(vListSht, startRow, endRow)
End Sub

Private Sub setCopiedRangeFont(ByRef vListSht As Worksheet, ByRef startRow As Long, ByRef endRow As Long)
    Dim vSiteName As String
    vSiteName = ChosenSiteArray(0)
    If ("" = vSiteName) Then
        Exit Sub
    End If
    
    Dim vSiteSheet As Worksheet
    Dim vListColumnNum As Long
    Set vSiteSheet = ThisWorkbook.Worksheets(vSiteName)
    vListColumnNum = GetColNumByRowIndex(vListSht) - 1
    
    Dim myRange As Range
    Set myRange = vSiteSheet.Range("B" + CStr(startRow) + ":" + C(vListColumnNum) + CStr(endRow))
    Call setRangeFont(myRange, "Arial")
End Sub

Public Sub setRangeFont(ByRef certainRange As Range, ByRef fontName As String)
    certainRange.Font.name = fontName
End Sub

Public Sub setRangeWrap(ByRef certainRange As Range, ByRef flag As Boolean)
    certainRange.WrapText = flag
End Sub

Public Sub FormatFirst(ByVal vListSht As Worksheet, ByVal strSheetName As String, ByVal SiteStartRow As Long, ByVal siteEndRow As Long)
    Dim vSiteIndex As Long
    Dim vSiteName As String
    Dim vSiteSheet As Worksheet
    Dim vRowIndex As Long
    Dim vListColumnNum As Long
    Dim myRange As Range
    Dim vName As String
    
    'For vSiteIndex = 0 To UBound(ChosenSiteArray)
        vSiteName = ChosenSiteArray(0)
        If ("" = vSiteName) Then
            Exit Sub
        End If
        
        Set vSiteSheet = Sheets(vSiteName)
        vListColumnNum = GetColNumByRowIndex(vListSht) - 1
        vName = GetCell(vListSht, g_TitleRowNo, 1)
        '如果是控制器则少生成一列
        If is_Controller(vName) Then
            vListColumnNum = vListColumnNum - 1
        End If
        vSiteSheet.Activate
        Set myRange = vSiteSheet.Range("A" + CStr(SiteStartRow) + ":" + C(vListColumnNum) + CStr(siteEndRow))
        Call AddListRefSub.setRangeBoard(myRange)

        myRange.Select
        With Selection.Interior
            .colorIndex = xlColorIndexNone
        End With
        vSiteSheet.Range("A" + CStr(SiteStartRow) + ":" + "A" + CStr(siteEndRow)).Merge
    'Next vSiteIndex
End Sub

Public Sub copyControllerTitle(ByVal vListSht As Worksheet, ByVal strSheetName As String, ByVal SiteStartRow As Long)
    Dim vSiteIndex As Long
    Dim vListColumnNum As Long
    Dim vSiteName As String
    Dim vSiteSheet As Worksheet
    Dim myRange As Range
    
    vListColumnNum = GetColNumByRowIndex(vListSht) - 1
    vSiteName = ChosenSiteArray(0)
    If ("" = vSiteName) Then
        Exit Sub
    End If
    Call WriteDatabyRow(vListSht, vSiteName, g_TitleRowNo, vListColumnNum - 1, SiteStartRow)
    Set vSiteSheet = Sheets(vSiteName)
    Call WriteSheetName(vSiteSheet, SiteStartRow, strSheetName)
    '关闭备注大小调整，提示刷表格效率
    Set myRange = vSiteSheet.Range("B" + CStr(SiteStartRow) + ":" + C(vListColumnNum - 1) + CStr(SiteStartRow))
    Call refreshComment(myRange)
End Sub

Public Sub CopyControllerData(ByVal vListSht As Worksheet, ByVal strSheetName As String, ByVal vMocEndRowNum As Long, ByVal SiteStartRow As Long, ByVal IsEmptyDataSheet As Boolean)
    Dim vSiteIndex As Long
    Dim vSiteName As String
    Dim vSiteSheet As Worksheet
    Dim vSiteLastRow As Long
    Dim vListColumnNum As Long
    Dim vRowIndex As Long
    Dim vSiteList As String
    Dim myRange As Range
    
    '最后一列是Site List,不需要拷贝
    vListColumnNum = GetColNumByRowIndex(vListSht) - 1
    For vSiteIndex = 0 To UBound(ChosenSiteArray)
        vSiteName = ChosenSiteArray(vSiteIndex)
        If ("" = vSiteName) Then
            Exit Sub
        End If
        
        Set vSiteSheet = Sheets(vSiteName)
        vSiteLastRow = SiteStartRow
        For vRowIndex = g_TitleRowNo To vMocEndRowNum
            vSiteList = vListSht.Cells(vRowIndex, vListColumnNum)
            If (vRowIndex = g_TitleRowNo) Then
                'Call WriteDatabyRow(vListSht, vSiteName, vRowIndex, vListColumnNum - 1, vSiteLastRow)
                'Call WriteSheetName(vSiteSheet, vSiteLastRow, strSheetName)

                '关闭备注大小调整，提示刷表格效率
                'Set myRange = vSiteSheet.range("B" + CStr(vSiteLastRow) + ":" + C(vListColumnNum - 1) + CStr(vSiteLastRow))
                'Call refreshComment(myRange)
                vSiteLastRow = vSiteLastRow + 1
            ElseIf IsSharedbySite(vSiteName, vSiteList, strSheetName, vListColumnNum) Then
                Call WriteDatabyRow(vListSht, vSiteName, vRowIndex, vListColumnNum - 1, vSiteLastRow)
                vSiteLastRow = vSiteLastRow + 1
            End If
            If (IsEmptyDataSheet) Then
                vRowIndex = vMocEndRowNum + 1
            End If
        Next vRowIndex
        
    Next vSiteIndex
End Sub

Private Function IsSharedbySite(ByVal vSiteName As String, ByVal vSiteList As String, ByVal vMocSheetName As String, ByVal MocColmunNo As String)
    IsSharedbySite = False
    If ("" = vSiteList) Then
        IsSharedbySite = True
        Exit Function
    End If
    
    Dim vSiteArray() As String
    Dim vSiteIndex As Long
    Dim vTmpSite As String
    
    vSiteArray() = Split(vSiteList, ",")
    For vSiteIndex = 0 To UBound(vSiteArray)
        vTmpSite = Trim(vSiteArray(vSiteIndex))
        If ("" = vTmpSite) Then
            Exit Function
        End If
        
        If (vSiteName = GetSiteSheetName(vTmpSite, vMocSheetName, MocColmunNo)) Then
            IsSharedbySite = True
            Exit Function
        End If
    Next vSiteIndex
End Function


Public Sub CopySiteData(ByVal vListSht As Worksheet, ByVal strSheetName As String, ByVal vMocEndRowNum As Long, ByVal SiteStartRow As Long, ByVal IsEmptyDataSheet As Boolean)
    Dim vRowIndex As Long
    Dim vNodeBName As String
    Dim vNodeBSheet As Worksheet
    Dim vSiteSheetName As String
    Dim vSiteLastRow As Long
    Dim vLastNodeName As String
    Dim vListColumnNum As Long
    Dim siteLastRowMap As New CMap '以每个基站为Key值，以最后一行空行为索引
    'Copy title first
    'Call CopySiteTitle(vListSht, strSheetName, SiteStartRow)
    
    If (IsEmptyDataSheet) Then
        Exit Sub
    End If
    
    vLastNodeName = ""
    vListColumnNum = GetColNumByRowIndex(vListSht) - 1
    vSiteLastRow = SiteStartRow + 1
    For vRowIndex = g_TitleRowNo + 1 To vMocEndRowNum
        vNodeBName = Trim(GetCell(vListSht, vRowIndex, 1))
        If ("" <> vNodeBName) Then
            vSiteSheetName = GetSiteSheetName(vNodeBName, strSheetName, 1)
            If ("" <> vSiteSheetName And IsChosenSite(vSiteSheetName)) Then
                Set vNodeBSheet = Sheets(vSiteSheetName)
                
                If Not siteLastRowMap.hasKey(vSiteSheetName) Then
                    vSiteLastRow = SiteStartRow + 1
                    Call siteLastRowMap.SetAt(vSiteSheetName, CStr(vSiteLastRow))
                Else
                    vSiteLastRow = CLng(siteLastRowMap.GetAt(vSiteSheetName))
                End If
'                If (vLastNodeName <> vSiteSheetName) Then
'                    vSiteLastRow = SiteStartRow + 1
'                End If
                '拷贝数据
                Call WriteDatabyRow(vListSht, vSiteSheetName, vRowIndex, vListColumnNum, vSiteLastRow)
                vSiteLastRow = vSiteLastRow + 1
                
                Call siteLastRowMap.SetAt(vSiteSheetName, CStr(vSiteLastRow)) '更新最后一行空行的记录
                
'                vLastNodeName = vSiteSheetName
            End If
        End If
    Next vRowIndex
End Sub

Public Sub CopySiteTitle(ByVal vListSht As Worksheet, ByVal strSheetName As String, ByVal SiteStartRow As Long)
    Dim vSiteIndex As Long
    Dim vListColumnNum As Long
    Dim vSiteName As String
    Dim vSiteSheet As Worksheet
    Dim myRange As Range
    
    vListColumnNum = GetColNumByRowIndex(vListSht) - 1
    'For vSiteIndex = 0 To UBound(ChosenSiteArray)
        vSiteName = ChosenSiteArray(0)
        If ("" = vSiteName) Then
            Exit Sub
        End If

        Call WriteDatabyRow(vListSht, vSiteName, g_TitleRowNo, vListColumnNum, SiteStartRow)
        
        Set vSiteSheet = Sheets(vSiteName)
        Call WriteSheetName(vSiteSheet, SiteStartRow, strSheetName)
        
        '关闭备注大小调整，提示刷表格效率
        Set myRange = vSiteSheet.Range("B" + CStr(SiteStartRow) + ":" + C(vListColumnNum) + CStr(SiteStartRow))
        Call refreshComment(myRange)
    'Next vSiteIndex
End Sub

Private Sub refreshComment(ByVal myRange As Range)
    On Error Resume Next
    Dim cell As Range
    For Each cell In myRange
        cell.comment.Shape.TextFrame.AutoSize = False
    Next
End Sub

Public Function GetSiteSheetName(ByVal siteName As String, ByVal mocSheetName As String, ByVal MocColmunNo As Long) As String
    GetSiteSheetName = ""
    If (IsSheetExist(siteName)) Then
        GetSiteSheetName = siteName
        Exit Function
    End If
    
    GetSiteSheetName = GetNodeByRate(siteName)
    If ("" <> GetSiteSheetName) Then
        Exit Function
    End If
    
    Dim groupName As String
    Dim columnName As String
    Dim mocSheet As Worksheet
    Dim columnNum As Long
    Dim findRowNo As Long
    
    Set mocSheet = Sheets(mocSheetName)
    groupName = GetCell(mocSheet, 1, MocColmunNo)
    columnName = GetCell(mocSheet, 2, MocColmunNo)
    If (0 = InStr(columnName, "*")) Then
        columnName = "*" + columnName
    End If
    
    columnNum = Get_RefColbyColumnName(getResByKey("BaseTransPort"), 2, columnName)
    If columnNum <= 0 Then
        Exit Function
    End If
    findRowNo = GetRowNumbyValue(siteName, columnNum)
    If (findRowNo < 1) Then
        Exit Function
    End If
    GetSiteSheetName = Trim(GetCell(Sheets(getResByKey("BaseTransPort")), findRowNo, 1))
    NodeArray(g_NodeIndex) = Trim(GetSiteSheetName)
    RateArray(g_NodeIndex) = Trim(siteName)
    g_NodeIndex = g_NodeIndex + 1
End Function

Public Function Get_RefColbyColumnName(sheetName As String, recordRow As Long, ColValue As String) As Long
    Dim vSiteSheet As Worksheet
    Dim vColIndex As Long
    Dim vListColumnNum As Long
    Dim vCell As String
    
    Get_RefColbyColumnName = 0
    Set vSiteSheet = Sheets(sheetName)
    vListColumnNum = GetColNumByRowIndex(vSiteSheet)
    
    For vColIndex = 1 To vListColumnNum
        vCell = vSiteSheet.Cells(recordRow, vColIndex)
        If (vCell = ColValue) Then
            Get_RefColbyColumnName = vColIndex
            Exit Function
        End If
    Next vColIndex
End Function

Public Function GetNodeByRate(ByVal siteName As String) As String
    Dim SiteIndex As Long
    Dim TmpSiteName As String
    
    GetNodeByRate = ""
    For SiteIndex = 0 To UBound(ChosenSiteArray)
        TmpSiteName = RateArray(SiteIndex)
        If ("" = TmpSiteName) Then
            Exit Function
        End If
        
        If (siteName = TmpSiteName) Then
            GetNodeByRate = NodeArray(SiteIndex)
            Exit Function
        End If
    Next SiteIndex
End Function

Public Function GetRowNumbyValue(ByVal value As String, ByVal columnNum As Long) As Long
    Dim baseTransSheetName As String
    baseTransSheetName = getResByKey("BaseTransPort")
    
    Dim baseTransSheet As Worksheet
    Set baseTransSheet = Sheets(baseTransSheetName)
    
    Dim index As Long
    Dim endRowNum As Long
    endRowNum = baseTransSheet.UsedRange.rows.count
    For index = 3 To endRowNum
        If (value = GetCell(baseTransSheet, index, columnNum)) Then
            GetRowNumbyValue = index
            Exit Function
        End If
    Next index
    GetRowNumbyValue = 0
End Function

'将MOC(Sheet页名称)写入到NodeB表
Public Sub WriteSheetName(ByVal vNodeBSheet As Worksheet, ByVal vNodeBLastRow As Long, ByVal strSheetName As String)
     vNodeBSheet.Cells(vNodeBLastRow, 1) = strSheetName
     With vNodeBSheet.Cells(vNodeBLastRow, 1)
        .ColumnWidth = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.name = "Microsoft Sans Serif"
        .Font.Bold = True
        .Font.Size = 10
        .Font.Underline = xlUnderlineStyleNone
        .Font.colorIndex = 2
        .WrapText = True
        .Interior.colorIndex = 9
        .Interior.Pattern = xlGray8
        .Interior.PatternColorIndex = xlAutomatic
    End With
End Sub

'将数据按行拷贝到NodeB表
Public Sub WriteDatabyRow(ByVal vListSht As Worksheet, ByVal vSiteName As String, ByVal vRowIndex As Long, ByVal vListColumnNum As Long, ByVal vSiteWrtRow As Long)
    'vListSht.Activate
    'vListSht.range("B" + CStr(vRowIndex) + ":" + C(vListColumnNum) + CStr(vRowIndex)).Select
    'Selection.Copy

    Dim vSiteSheet As Worksheet
    Set vSiteSheet = Sheets(vSiteName)
    vSiteSheet.Activate
    vSiteSheet.Cells(vSiteWrtRow, 2).Select
    vListSht.Range("B" + CStr(vRowIndex) + ":" + C(vListColumnNum) + CStr(vRowIndex)).Copy Destination:=vSiteSheet.Cells(vSiteWrtRow, 2)
    'ActiveSheet.Paste
End Sub


'获取Sheet页某行的列表
Public Function GetColNumByRowIndex(ByVal vWorkSht As Worksheet) As Long
    Dim vColumnIndex As Long
    vColumnIndex = 1
    
    While ("" <> Trim(GetCell(vWorkSht, g_TitleRowNo, vColumnIndex)))
        vColumnIndex = vColumnIndex + 1
    Wend
    
    GetColNumByRowIndex = vColumnIndex
End Function
'chenjun end

'sonic begin
'**********************************************************
'从列数得到列名：1->A，27->AA
'**********************************************************
Public Function C(iColumn As Long) As String
  If iColumn >= 257 Or iColumn < 0 Then
    C = ""
    Return
  End If
  
  Dim Result As String
  Dim High, Low As Long
  
  High = Int((iColumn - 1) / 26)
  Low = iColumn Mod 26
  
  If High > 0 Then
    Result = Chr(High + 64)
  End If

  If Low = 0 Then
    Low = 26
  End If
  
  Result = Result & Chr(Low + 64)
  C = Result
End Function

'**********************************************************
'从列名得到列数：A->1，AA->27
'**********************************************************
Public Function D(ColumnStr As String) As Long
  If Len(ColumnStr) = 1 Then
    D = Int(ColumnStr) - 64
  ElseIf Len(ColumnStr) = 2 Then
    D = (Int(Left(ColumnStr, 1)) - 64) * 26 + (Int(Left(ColumnStr, 1)) - 64)
  End If
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

Public Sub CreateNewSiteSheet_i(NodeBName As String)
    'Application.StatusBar = "create sheet:[" + NodeBName + "] begin"
    'If IsSheetExist(NodeBName) Then
    '    Sheets(NodeBName).Delete
    'End If
    
    Sheets.Add after:=ThisWorkbook.ActiveSheet
    ThisWorkbook.ActiveSheet.name = NodeBName
    
    ThisWorkbook.ActiveSheet.Tab.colorIndex = BluePrintSheetColor
    
    With ThisWorkbook.ActiveSheet.Cells.Interior
        .colorIndex = 15
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
    End With
End Sub

Public Function CreateNewSiteSheet()
    Dim siteName As String
    
    Dim SiteIndex As Long
    SiteIndex = 0
    
    Dim Iter As Long
    Sheets(GetMainSheetName()).Activate
    'For Iter = 0 To MaxChosenSiteNum - 1
        siteName = ChosenSiteArray(0)
        If siteName = "" Then
            Exit Function
        End If
        
        Call CreateNewSiteSheet_i(siteName)
    'Next Iter
End Function

'按第一列算最大有值行
Public Function GetSheetUsedRows(sheet As Worksheet) As Long
    Dim maxRow As Long
    maxRow = 1
    
    Do While Trim(sheet.Cells(maxRow + 1, 1)) <> ""
        maxRow = maxRow + 1
    Loop
    
    GetSheetUsedRows = maxRow
End Function

'按第一行算最大有值列
Public Function GetSheetUsedColumns(sheet As Worksheet) As Long
  Dim MaxColumn As Long
  MaxColumn = 1
  
  Do While Trim(sheet.Cells(2, MaxColumn + 1)) <> ""
    MaxColumn = MaxColumn + 1
  Loop
  
  GetSheetUsedColumns = MaxColumn
End Function

Public Function GetLastRow(sheetName) As Long
        GetLastRow = Worksheets(sheetName).UsedRange.rows.row
End Function

'copy Pattern页后，定位到NodeB页，返回title行数
Public Function LocateLastRow(NodeBName) As Long
    Dim sheet As Worksheet
    Set sheet = Sheets(NodeBName)
    
    Dim NodeBBeginRow As Long
    Dim NodeBEndRow As Long
    NodeBBeginRow = 1
    NodeBEndRow = 1000
    
    Dim NodeBRow As Long
    NodeBRow = NodeBBeginRow
    
    Do While NodeBRow <= NodeBEndRow
        If Application.WorksheetFunction.CountA(sheet.rows(NodeBRow)) = 0 Then 'when the row is empty
            sheet.Activate
            sheet.Cells(NodeBRow, 2).Select
            LocateLastRow = NodeBRow
            Exit Function
        End If
        NodeBRow = NodeBRow + 1
    Loop
End Function

Public Function GetMocNameFromPatternName(PatternName As String)
    GetMocNameFromPatternName = Replace(PatternName, "Pattern", "")
End Function

Public Sub ConvertPattern_i(sheetName As String, NodeBName As String)
    Application.StatusBar = "ConvertPattern_i:" + sheetName + "NodeBName:" + NodeBName
    
    Dim PatternSheet As Worksheet
    Set PatternSheet = Sheets(sheetName)
    Dim PatternSheetBeginRow As Long
    Dim PatternSheetEndRow As Long
    PatternSheetBeginRow = 2
    PatternSheetEndRow = GetSheetUsedRows(PatternSheet)
    
    Dim PatternSheetEndColumn As Long
    PatternSheetEndColumn = GetSheetUsedColumns(PatternSheet)
    
    If PatternSheetEndRow > 2 Then
        PatternSheet.Activate
        PatternSheet.Range("A2:" + C(PatternSheetEndColumn) + CStr(PatternSheetEndRow)).Select
        Selection.Copy
        
        Dim TitleRow As Long
        TitleRow = LocateLastRow(NodeBName)
        ThisWorkbook.ActiveSheet.Paste
        ThisWorkbook.ActiveSheet.Cells(TitleRow, 1) = GetMocNameFromPatternName(sheetName)
        
        With ThisWorkbook.ActiveSheet.Cells(TitleRow, 1)
            .Merge
            .ColumnWidth = 20
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.name = "Microsoft Sans Serif"
            .Font.Bold = True
            .Font.Size = 10
            .Font.Underline = xlUnderlineStyleNone
            .Font.colorIndex = 2
            .WrapText = True
            .Interior.colorIndex = 9
            .Interior.Pattern = xlGray8
            .Interior.PatternColorIndex = xlAutomatic
        End With
    End If
End Sub

Public Sub ConvertPattern(PatternSheetName As String)
    Application.StatusBar = "ConvertPattern:" + PatternSheetName
    
    Dim siteName As String
    Dim Iter As Long
    For Iter = 0 To MaxChosenSiteNum - 1
        siteName = ChosenSiteArray(Iter)
        If siteName = "" Then
            Exit For
        End If
        Call ConvertPattern_i(PatternSheetName, siteName)
    Next Iter
End Sub

Public Function IsPatternLink(str As String) As Boolean
    IsPatternLink = False
    
    Dim ArrData() As String
    ArrData = Split(str, "\")
    If UBound(ArrData) >= 2 Then
        IsPatternLink = True
    End If
End Function

Public Function GetValueFromBaseSheet(NodeBName As String, RealColumn As Long, value As String) As String
    On Error GoTo ErrorHandler:
    Dim sheet As Worksheet
    Set sheet = Sheets(getResByKey("BaseTransPort"))
    
    Dim NodeBBeginRow As Long
    Dim NodeBEndRow As Long
    NodeBBeginRow = g_NodeBBeginRow
    NodeBEndRow = sheet.UsedRange.rows.count
        
    Dim NodeBNameColumn As Long
    NodeBNameColumn = 1
    
    Dim NodeBRow As Long
    NodeBRow = NodeBBeginRow
    
    Do While NodeBBeginRow <= NodeBEndRow And NodeBBeginRow < MaxChosenSiteNum
        If sheet.Cells(NodeBBeginRow, NodeBNameColumn) = NodeBName Then
            GetValueFromBaseSheet = sheet.Cells(NodeBBeginRow, RealColumn)
            Exit Function
        End If
        NodeBBeginRow = NodeBBeginRow + 1
    Loop
    GetValueFromBaseSheet = value
    Exit Function
ErrorHandler:
    GetValueFromBaseSheet = value
End Function

Public Function GetPatternRealValue(NodeBName As String, value As String) As String
    Dim Vec() As String
    Dim sheetName As String
    Dim groupName As String
    Dim ColName As String
    
    Vec = Split(value, "\")
    sheetName = Vec(0)
    groupName = Vec(1)
    ColName = Vec(2)

    Dim RealColumn As Long
    RealColumn = Get_RefCol(sheetName, 2, groupName, ColName)

    GetPatternRealValue = GetValueFromBaseSheet(NodeBName, RealColumn, value)
End Function

Public Sub RefreshPatternValue_i(NodeBName As String)
    Application.StatusBar = "refresh nodeb:" + NodeBName
    
    Dim NodeBSheet As Worksheet
    Set NodeBSheet = Sheets(NodeBName)
    
    Dim endRow As Long
    Dim EndColumn As Long
    endRow = NodeBSheet.UsedRange.rows.count
    EndColumn = NodeBSheet.UsedRange.columns.count
    
    Dim IterRow As Long
    Dim IterColumn As Long
    
    Dim value As String
    For IterRow = 1 To endRow
        For IterColumn = 1 To EndColumn
            value = NodeBSheet.Cells(IterRow, IterColumn)
            NodeBSheet.Cells(IterRow, IterColumn).Hyperlinks.Delete
            If value <> "" And IsPatternLink(value) Then
                With NodeBSheet.Cells(IterRow, IterColumn).Validation
                    .Delete
                    .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertInformation
                    .inputTitle = getResByKey("Reference Address")
                    .inputMessage = NodeBSheet.Cells(IterRow, IterColumn)
                    .ShowInput = True
                    .ShowError = False
                End With
                NodeBSheet.Cells(IterRow, IterColumn) = GetPatternRealValue(NodeBName, value)
                With NodeBSheet.Cells(IterRow, IterColumn).Interior
                    .colorIndex = HyperLinkColorIndex
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                End With
            End If
        Next IterColumn
    Next IterRow
End Sub

'针对每个NodeB Sheet页，进行PatternLink值的刷新
Public Sub RefreshPatternValue()
    Dim siteName As String
    Dim Iter As Long
    For Iter = 0 To MaxChosenSiteNum - 1
        siteName = ChosenSiteArray(Iter)
        If siteName = "" Then
            Exit For
        End If
    
        Call RefreshPatternValue_i(siteName)
    Next Iter
End Sub

Public Function JudgingEmptyDataSheet(ListSheetName As String) As Boolean
    Dim sheet As Worksheet
    Set sheet = Worksheets(ListSheetName)
    JudgingEmptyDataSheet = sheet.Cells(3, 1) = "" ' And Sheet.Cells(3, 2) = ""
End Function

Public Sub writeData()
    Dim startRow As Long
    Dim endRow As Long
    startRow = 1
    Dim IsEmptyDataSheet As Boolean
    Dim Iter As Long
    Dim ListSheetName As String
    Dim sheetDef As Worksheet
    Set sheetDef = Worksheets("SHEET DEF")
    For Iter = 2 To sheetDef.Range("a65536").End(xlUp).row
        If sheetDef.Cells(Iter, 4).value <> "" And sheetDef.Cells(Iter, 4).value <> "" Then
            startRow = CStr(sheetDef.Cells(Iter, 4).value)
            endRow = CStr(sheetDef.Cells(Iter, 5).value)
            ListSheetName = sheetDef.Cells(Iter, 1).value
            If IsSheetExist(ListSheetName) Then
                If FILE_TYPE = "1" Or Not isTrasnPortSheet(ListSheetName) Then
                    IsEmptyDataSheet = JudgingEmptyDataSheet(ListSheetName)
                    Call writeListData(ListSheetName, startRow, endRow, IsEmptyDataSheet)
                    Application.DisplayAlerts = False
                    Sheets(ListSheetName).Delete
                    Application.DisplayAlerts = True
                End If
            End If
        End If
    Next
End Sub

Public Sub writeListData(ByVal strSheetName As String, ByVal startRow As Long, ByVal endRow As Long, ByVal IsEmptyDataSheet As Boolean)
    Dim vListSht As Worksheet
    Dim vName As String
    Dim vMocEndRowNum As Long
    
    Set vListSht = Sheets(strSheetName)
    vName = GetCell(vListSht, g_TitleRowNo, 1)
    '只处理NodeB&BTS对象
    If (Not is_SiteorController(vName)) Or (isTrasnPortSheet(strSheetName) And FILE_TYPE = "0") Then
        Exit Sub
    End If
    
    vMocEndRowNum = vListSht.Range("a65536").End(xlUp).row
    If (vMocEndRowNum <= g_TitleRowNo) Then
        vMocEndRowNum = g_TitleRowNo + 1
    End If

    If is_Controller(vName) Then
'        Call copyControllerTitle(vListSht, strSheetName, startRow)
        Call CopyControllerData(vListSht, strSheetName, vMocEndRowNum, startRow, IsEmptyDataSheet)
    Else
'        Call CopySiteTitle(vListSht, strSheetName, startRow)
       Call CopySiteData(vListSht, strSheetName, vMocEndRowNum, startRow, IsEmptyDataSheet)
    End If
End Sub

Public Sub InvokeConversion()
    Dim startRow As Long
    Dim endRow As Long
    startRow = 1
    Dim IsEmptyDataSheet As Boolean
    Dim Iter As Long
    Dim ListSheetName As String
    Dim sheetDef As Worksheet
    Set sheetDef = Worksheets("SHEET DEF")
    
    Dim bscFlag As Boolean, rncFlag As Boolean, btsFlag As Boolean, alreadyAddEmptyRow As Boolean
    bscFlag = False
    rncFlag = False
    btsFlag = False
    alreadyAddEmptyRow = False
    
    For Iter = 2 To sheetDef.Range("a65536").End(xlUp).row
        If UCase(sheetDef.Cells(Iter, 2).value) = UCase("List") Then
            ListSheetName = sheetDef.Cells(Iter, 1).value
            If IsSheetExist(ListSheetName) Then
                If FILE_TYPE = "1" Or Not isTrasnPortSheet(ListSheetName) Then
                    IsEmptyDataSheet = JudgingEmptyDataSheet(ListSheetName)
                    If (IsEmptyDataSheet) Then
                        endRow = startRow + 1
                    Else
                        endRow = startRow + GetMaxCountPerSite(ListSheetName)
                    End If
                    
                    Call setBscRncBtsFlag(bscFlag, rncFlag, btsFlag, ListSheetName)
                    
                    '如果先找到了控制器再找到基站，则增加一行空行
                    If alreadyAddEmptyRow = False And (bscFlag = True Or rncFlag = True) And btsFlag = True Then
                        startRow = startRow + 1
                        endRow = endRow + 1
                        alreadyAddEmptyRow = True
                    End If
                    Call ConvertList(ListSheetName, startRow, endRow, IsEmptyDataSheet)
                    
                    'Application.DisplayAlerts = False
                    'Sheets(ListSheetName).Delete
                    'Application.DisplayAlerts = True
            
                    sheetDef.Cells(Iter, 4) = CStr(startRow)
                    sheetDef.Cells(Iter, 5) = CStr(endRow)
                    startRow = endRow + 1
                End If
            End If
        ElseIf UCase(sheetDef.Cells(Iter, 2).value) = UCase("Pattern") Then
            'Call ConvertPattern(sheet.name) 'ignore pattern process
        End If
    Next
    'Call RefreshPatternValue 'ignore pattern process
    Call WriteSheeDefTitle
End Sub

Private Sub setBscRncBtsFlag(ByRef bscFlag As Boolean, ByRef rncFlag As Boolean, ByRef btsFlag As Boolean, ByRef ListSheetName As String)
    If InStr(ListSheetName, "(BSC)") <> 0 Then
        bscFlag = True
    ElseIf InStr(ListSheetName, "(RNC)") <> 0 Then
        rncFlag = True
    Else
        btsFlag = True
    End If
End Sub

'Added by chenjun
Public Sub InitSiteNameArray()
    g_NodeIndex = 0
    Dim SiteIndex As Long
    For SiteIndex = 0 To UBound(ChosenSiteArray)
        NodeArray(SiteIndex) = ""
        RateArray(SiteIndex) = ""
    Next SiteIndex
End Sub

'Added by chenjun
Public Sub FormatSiteSheet()
    Dim SiteIndex As Long
    Dim SiteSht As Worksheet
    Dim siteName As String
    
    SiteIndex = 0
    siteName = ChosenSiteArray(SiteIndex)
    'While ("" <> SiteName And IsChosenSite(SiteName))
        Set SiteSht = Sheets(siteName)
        With SiteSht.Cells
            .ColumnWidth = 20
            .WrapText = False
        End With
        '冻结首列
        SiteSht.Select
        With ActiveWindow
            .SplitColumn = 1
            .SplitRow = 0
        End With
        ActiveWindow.FreezePanes = True
        
        'SiteIndex = SiteIndex + 1
        'SiteName = ChosenSiteArray(SiteIndex)
    'Wend
End Sub

Private Sub clearChosenSiteArray()
        Dim index As Long
        For index = LBound(ChosenSiteArray) To UBound(ChosenSiteArray)
                ChosenSiteArray(index) = ""
        Next
End Sub

Private Function siteNameExist(name As String) As Boolean
        Dim index As Long
        For index = LBound(ChosenSiteArray) To UBound(ChosenSiteArray)
                If LCase(ChosenSiteArray(index)) = LCase(name) Then
                    siteNameExist = True
                    Exit Function
                End If
        Next
        siteNameExist = False
End Function

Public Function InitChosenSiteArray() As Boolean
    InitChosenSiteArray = True
    'clear ChosenSiteArray
    Dim I As Long
    For I = 1 To MaxChosenSiteNum
        ChosenSiteArray(I) = ""
    Next I
    
    'fill in ChosenSiteArray
    Dim NodeBBaseSheet As Worksheet
    Set NodeBBaseSheet = Sheets(getResByKey("BaseTransPort"))
    
    Dim NodeBBeginRow As Long
    Dim NodeBEndRow As Long
    NodeBBeginRow = g_NodeBBeginRow
    NodeBEndRow = GetSheetUsedRows(Worksheets(getResByKey("BaseTransPort")))
    If NodeBEndRow > MaxChosenSiteNum Then
        NodeBEndRow = MaxChosenSiteNum
    ElseIf NodeBEndRow <= g_NodeBBeginRow - 1 Then 'if no sites has been found, abort with msg
        If bIsEng Then
            MsgBox "No site has been found.", vbOKOnly
        Else
            MsgBox "找不到基站.", vbOKOnly
        End If
        InitChosenSiteArray = False
        Exit Function
    End If

    Call clearChosenSiteArray
    
    Dim NodeBNameColumn As Long
    NodeBNameColumn = 1
    
    Dim NodeBRow As Long
    NodeBRow = NodeBBeginRow
    
    Dim siteName As String
    Dim IgnoredSiteNum As Long
    IgnoredSiteNum = 0
    Do While NodeBBeginRow <= NodeBEndRow
        siteName = NodeBBaseSheet.Cells(NodeBBeginRow, NodeBNameColumn)
        If siteName = "" Then
            Exit Do
        End If
        
        If IsValidSiteName(siteName) = False Then
            InitChosenSiteArray = False
            Exit Function
        ElseIf siteNameExist(siteName) Then
            If bIsEng Then
                MsgBox "The name of the [" + siteName + "] is same with another site,can't generate Moc-Integration view.", vbCritical
            Else
                MsgBox "存在与[" + siteName + "]同名的基站, 无法生成MOC整体视图.", vbCritical
            End If
            InitChosenSiteArray = False
            Exit Function
        ElseIf IsSheetExist(siteName) Then
            If bIsEng Then
                MsgBox "Sheet name[" + siteName + "] has already been occupied, and this site will be ignored in this operation.", vbOKOnly
            Else
                MsgBox "页签名称[" + siteName + "]已被占用,此基站在本次生成中将被忽略.", vbOKOnly
            End If
            NodeBBeginRow = NodeBBeginRow + 1
            IgnoredSiteNum = IgnoredSiteNum + 1
        Else
            ChosenSiteArray(NodeBBeginRow - g_NodeBBeginRow - IgnoredSiteNum) = Trim(siteName)
            NodeBBeginRow = NodeBBeginRow + 1
        End If
    Loop
End Function

Function IsValidSiteName(siteName As String) As Boolean
  IsValidSiteName = True
  
  If Len(siteName) > 31 _
     Or InStr(siteName, "\") > 0 _
     Or InStr(siteName, "/") > 0 _
     Or InStr(siteName, ":") > 0 _
     Or InStr(siteName, "*") > 0 _
     Or InStr(siteName, "?") > 0 _
     Or InStr(siteName, Chr(34)) > 0 _
     Or InStr(siteName, "<") > 0 _
     Or InStr(siteName, ">") > 0 _
     Or InStr(siteName, "|") > 0 _
     Or InStr(siteName, ",") > 0 _
     Or InStr(siteName, ";") > 0 _
     Or InStr(siteName, "=") > 0 _
     Or InStr(siteName, "!") > 0 _
     Or InStr(siteName, "@") > 0 _
     Or InStr(siteName, "#") > 0 _
     Or InStr(siteName, "&") > 0 _
     Or InStr(siteName, "%") > 0 _
     Or InStr(siteName, "^") > 0 _
     Or InStr(siteName, "[") > 0 _
     Or InStr(siteName, "]") > 0 _
     Or InStr(siteName, "  ") > 0 _
     Or InStr(siteName, "+++") > 0 _
     Or Trim(siteName) = "" Then
     MsgBox Replace(getResByKey("SITE_NAME_INVALID"), "(%0)", siteName) + vbCrLf
     IsValidSiteName = False
  End If

End Function


'Public Sub ShowChooseSiteForm()
'    '==========================load form and set it's language related
'    Load ChooseSiteForm
'
'    If bIsEng Then
'        ChooseSiteForm.Caption = "Please choose sites to generate Moc-Integration View"
'        ChooseSiteForm.OKButton.Caption = "OK"
'        ChooseSiteForm.CancelButton.Caption = "Cancel"
'    Else
'        ChooseSiteForm.Caption = "请选择需要生成整体视图的站"
'        ChooseSiteForm.OKButton.Caption = "确 定"
'        ChooseSiteForm.CancelButton.Caption = "取 消"
'    End If
'
'    '===========================collect NodeBs to form
'    Dim NodeBBaseSheet As Worksheet
'    Set NodeBBaseSheet = Sheets(getResByKey("BaseTransPort"))
'
'    Dim NodeBBeginRow as long
'    Dim NodeBEndRow as long
'    NodeBBeginRow = g_NodeBBeginRow
'    NodeBEndRow = GetSheetUsedRows(Worksheets(getResByKey("BaseTransPort")))
'    If NodeBEndRow > MaxChosenSiteNum Then
'        NodeBEndRow = MaxChosenSiteNum
'    ElseIf NodeBEndRow <= g_NodeBBeginRow - 1 Then 'if no sites has been found, abort with msg
'        If bIsEng Then
'            MsgBox "No site has been found.", vbOKOnly
'        Else
'            MsgBox "找不到基站.", vbOKOnly
'        End If
'        Exit Sub
'    End If
'
'    Dim NodeBNameColumn as long
'    NodeBNameColumn = 1
'
'    Dim NodeBRow as long
'    NodeBRow = NodeBBeginRow
'
'    Dim NodeBName As String
'    Do While NodeBBeginRow <= NodeBEndRow
'        NodeBName = NodeBBaseSheet.Cells(NodeBBeginRow, NodeBNameColumn)
'        If NodeBName = "" Then
'            Exit Do
'        End If
'
'        ChosenSiteArray(NodeBBeginRow - g_NodeBBeginRow) = NodeBName
'        NodeBBeginRow = NodeBBeginRow + 1
'    Loop
'
'    ReDim ShrinkChosen(0, NodeBBeginRow - g_NodeBBeginRow - 1) As String
'    Dim I as long
'    For I = 0 To NodeBBeginRow - g_NodeBBeginRow - 1
'        ShrinkChosen(0, I) = ChosenSiteArray(I)
'    Next I
'    ChooseSiteForm.SiteListBox.column() = ShrinkChosen
'    ChooseSiteForm.Show vbModeless
'End Sub

Sub GenIubFormatReport()
    Call InitTemplateVersion
    If InitChosenSiteArray Then
        'Call SortListLine
        Call GenIubFormatReport_i
    End If
End Sub

Private Sub createIubStyleWorkSheet(wb As Workbook, iubStyleSheet As Worksheet)
    On Error GoTo ErrorHandler
    Dim index As Long
    Dim count As Long
    Dim arrayIndex As Long
    
    Dim sheet As Worksheet
    
    Dim tempSheetName As String
    tempSheetName = "My Sheet"
    iubStyleSheet.name = tempSheetName
    
    Call setDataRowsWrap(tempSheetName)   '无数据行的自动换行原来没有，新增为其设置上
    
    For index = LBound(ChosenSiteArray) + 1 To UBound(ChosenSiteArray)
           If ChosenSiteArray(index) = "" Then
                Exit For
           End If
           iubStyleSheet.Copy after:=iubStyleSheet
    Next
    count = wb.Worksheets.count
    arrayIndex = 0
    For index = 1 To count
        If wb.Worksheets(index).Tab.colorIndex = BluePrintSheetColor And arrayIndex <= UBound(ChosenSiteArray) Then
            If ChosenSiteArray(arrayIndex) = "" Then
                Exit For
            End If
            wb.Worksheets(index).name = ChosenSiteArray(arrayIndex)
            arrayIndex = arrayIndex + 1
            If Not sheet Is Nothing Then
                wb.Worksheets(index).Move after:=sheet
            End If
            Set sheet = wb.Worksheets(index)
        End If
    Next
    Exit Sub
ErrorHandler:
End Sub

Private Sub setDataRowsWrap(ByRef sheetName As String)
    Dim sheetDefSheet As Worksheet, iubSheet As Worksheet
    Set sheetDefSheet = ThisWorkbook.Worksheets("SHEET DEF")
    Set iubSheet = ThisWorkbook.Worksheets(sheetName)
    
    Dim rowIndex As Long, maxColumnNumber As Long
    Dim startRow As String, endRow As String
    Dim dataRange As Range
    '遍历Sheet Def页，把每个IUB页签中的空行都设置上自动换行
    For rowIndex = 2 To sheetDefSheet.Range("A65535").End(xlUp).row
        startRow = sheetDefSheet.Range("D" & rowIndex).value
        endRow = sheetDefSheet.Range("E" & rowIndex).value
        If startRow <> "" And endRow <> "" Then
            maxColumnNumber = iubSheet.Range("IV" & startRow).End(xlToLeft).column
            Set dataRange = iubSheet.Range("B" & (startRow + 1) & ":" & C(maxColumnNumber) & endRow)
            Call setRangeWrap(dataRange, True)
        End If
    Next rowIndex
End Sub


Sub GenIubFormatReport_i()
    Call refreshStart
    Call InitSiteNameArray
    Call InitTemplateVersion
    
    '第一步，生成NodeB Sheet页
    Call CreateNewSiteSheet

    '第二步，递归所有数据页，调用ConvertList或ConvertPattern
    Call InvokeConversion

    Call FormatSiteSheet
    Call createIubStyleWorkSheet(ThisWorkbook, ThisWorkbook.Worksheets(ChosenSiteArray(0)))
    Call writeData
    Call ContructPositionInfo
    Call refreshEnd
    If bIsEng Then
        MsgBox "Finish generating view.", vbOKOnly
    Else
        MsgBox "生成视图完成.", vbOKOnly
    End If
End Sub

Public Sub InitTemplateVersion()
    bIsEng = getResByKey("Cover") = "Cover"
End Sub

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

Public Function IsBluePrintSheetName(sheetName As String) As Boolean
    IsBluePrintSheetName = (Sheets(sheetName).Tab.colorIndex = BluePrintSheetColor)
End Function

Public Function HasNoBluePrintSheet() As Boolean
    HasNoBluePrintSheet = (GetBluePrintSheetName = "")
End Function

Public Function IsChosenSite(ByVal chkSiteName As String) As Boolean
    IsChosenSite = False
    Dim siteName As String
    Dim Iter As Long
    For Iter = 0 To MaxChosenSiteNum - 1
        siteName = ChosenSiteArray(Iter)
        If siteName = "" Then
            Exit For
        End If
        
        If UCase(siteName) = UCase(chkSiteName) Then
            IsChosenSite = True
            Exit For
        End If
    
    Next Iter
End Function

Public Sub test()
    Dim I As Long
    I = GetMaxCountPerSite(ThisWorkbook.ActiveSheet.name)
    MsgBox CStr(I), vbOKOnly
End Sub

Public Function IndexInArray(ByRef SiteNames() As String, siteName As String) As Long
    IndexInArray = -1
    Dim Iter As Long
    For Iter = 0 To UBound(SiteNames)
        If SiteNames(Iter) = siteName Then
            IndexInArray = Iter
            Exit Function
        End If
    Next Iter
End Function

Public Function GetMaxCountPerSite_i(ByRef ListSheet As Worksheet) As Long
    'fill in SiteNames and MocCount
    Dim SiteNames(MaxChosenSiteNum) As String
    Dim MocCounts(MaxChosenSiteNum) As Long
    
    Dim SiteNameColumn As Long
    SiteNameColumn = 1
    
    Dim SiteBeginRow As Long
    Dim siteEndRow As Long
    SiteBeginRow = 3
    siteEndRow = GetSheetUsedRows(ListSheet)
    
    Dim index As Long
    Dim siteName As String
    Dim SiteCount As Long
    SiteCount = 0
    Do While SiteBeginRow <= siteEndRow
        siteName = ListSheet.Cells(SiteBeginRow, SiteNameColumn)
        
        index = IndexInArray(SiteNames, siteName)
        If index = -1 Then
            SiteNames(SiteCount) = siteName
            MocCounts(SiteCount) = 1
            SiteCount = SiteCount + 1
        Else
            MocCounts(index) = MocCounts(index) + 1
        End If
        SiteBeginRow = SiteBeginRow + 1
    Loop
    
    'get the max count
    Dim MaxCount As Long
    MaxCount = 0
    
    Dim Iter As Long
    For Iter = 0 To UBound(MocCounts)
        If MaxCount < MocCounts(Iter) Then
            MaxCount = MocCounts(Iter)
        End If
    Next Iter
    
    GetMaxCountPerSite_i = MaxCount + 1
End Function

Public Function GetMaxCountPerSite(ListName As String) As Long
    If (Not IsSheetExist(ListName)) Then
        GetMaxCountPerSite = 0
        Exit Function
    End If
    
    Dim ListSheet As Worksheet
    Set ListSheet = Sheets(ListName)
    
    If is_DelColumnName(GetCell(ListSheet, g_TitleRowNo, 1)) Then 'Site related List
        GetMaxCountPerSite = GetMaxCountPerSite_i(ListSheet)
    Else                                                          'Controller related list
        Dim row As Long
        Dim count As Long
        count = ListSheet.columns.count
        For row = 1 To ListSheet.UsedRange.rows.count
            If Application.WorksheetFunction.CountBlank(ListSheet.rows(row)) = count Then
                Exit For
            End If
        Next
        GetMaxCountPerSite = row
    End If
End Function

Public Function GetSheetUsedColumnsForRow(sheet As Worksheet, row As Long) As Long
  Dim MaxColumn As Long
  MaxColumn = 1
  
  Do While Trim(sheet.Cells(row, MaxColumn + 1)) <> ""
    MaxColumn = MaxColumn + 1
  Loop
  
  GetSheetUsedColumnsForRow = MaxColumn
End Function

Public Sub WriteSheeDefTitle()
    Dim sheetDef As Worksheet
    Set sheetDef = Sheets("SHEET DEF")
    
    Dim startColumn As Long
    startColumn = GetSheetUsedColumnsForRow(sheetDef, 1)
    sheetDef.Cells(1, startColumn + 1) = "StartRow"
    sheetDef.Cells(1, startColumn + 2) = "EndRow"
End Sub

Public Sub ContructPositionInfo()
    Dim sheet As Worksheet
    Set sheet = Sheets(getResByKey("BaseTransPort"))
    sheet.Activate
    
    Dim SiteMergeCount As Long
    With sheet
        rows("3:3").Select
        Selection.Insert Shift:=xlDown
        rows("3:3").Hidden = True
        Range("B2").Select
        Selection.EntireColumn.Insert
        
        If bIsEng Then
            Cells(2, 2) = "Referenced Site"
        Else
            Cells(2, 2) = "引用的站"
        End If
        
        Range("A1").Select
        
        If Selection.count = 1 Then
            Range("A1:B1").Select
            Selection.Merge
        End If
    End With
End Sub

'Public Sub SortListLine()
'    Dim sheetDef As Worksheet
'    Set sheetDef = Worksheets("SHEET DEF")
'
'    'find the start row of site level and controller level, and the last row
'    Dim SiteStartRow as long
'    Dim ControllerStartRow as long
'    Dim LastRow as long
'    SiteStartRow = 0
'    ControllerStartRow = 0
'    LastRow = sheetDef.range("a65536").End(xlUp).row
'
'    Dim Iter as long
'    Dim ListSheetName As String
'    For Iter = 2 To sheetDef.range("a65536").End(xlUp).row
'        If UCase(sheetDef.Cells(Iter, 2).value) = UCase("List") Then
'            ListSheetName = sheetDef.Cells(Iter, 1).value
'            If Not isTrasnPortSheet(ListSheetName) Then
'                If sheetDef.Cells(Iter, 3).value = "" Then 'Controller level
'                    If ControllerStartRow = 0 Then
'                        ControllerStartRow = Iter
'                    End If
'                Else 'site level
'                    If SiteStartRow = 0 Then
'                        SiteStartRow = Iter
'                    End If
'                End If
'            End If
'        End If
'    Next Iter
'
'    If SiteStartRow > 0 And ControllerStartRow > 0 And SiteStartRow < ControllerStartRow Then 'site level mocs are beyond controller level mocs
'        sheetDef.Visible = xlSheetVisible
'        sheetDef.Select
'        sheetDef.rows(CStr(ControllerStartRow) + ":" + CStr(LastRow)).Select
'        Selection.Cut
'        rows(CStr(SiteStartRow) + ":" + CStr(SiteStartRow)).Select
'        Selection.Insert Shift:=xlDown
'        range("A1").Select
'        sheetDef.Visible = xlSheetHidden
'    End If
'End Sub










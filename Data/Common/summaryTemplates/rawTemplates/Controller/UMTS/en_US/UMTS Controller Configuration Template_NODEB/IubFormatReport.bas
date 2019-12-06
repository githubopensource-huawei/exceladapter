Attribute VB_Name = "IubFormatReport"
Option Explicit

Public bIsEng As Boolean  '���ڿ���������Ӣ��
'==================================================================================================
'����Iub��ʽ����
'==================================================================================================
Public Const HyperLinkColorIndex = 6
Public Const BluePrintSheetColor = 5
Public Const MaxChosenSiteNum = 202

Public Const g_NodeBBeginRow = 3
Public Const g_TitleRowNo As Long = 2
Public NodeArray(MaxChosenSiteNum) As String
Public RateArray(MaxChosenSiteNum) As String
Public g_NodeIndex As Long

Public ChosenSiteArray(MaxChosenSiteNum) As String
Public ChosenSiteNames As New Collection

Public FILE_TYPE As String '0- Summary 1-Bulk

Sub setFileType(fileType As String)
        FILE_TYPE = fileType
End Sub
Function is_SiteorController(ByVal columnName As String) As Boolean
    is_SiteorController = False
    If (columnName = "*NodeB Name" Or columnName = getResByKey("A236") Or _
        columnName = "*BTS Name" Or columnName = getResByKey("A238") Or _
        columnName = "*eNodeB Name" Or columnName = getResByKey("A239") Or _
        columnName = "*USU Name" Or columnName = getResByKey("A240") Or _
        columnName = getResByKey("*NBBSName") Or _
        is_Controller(columnName)) Then
        is_SiteorController = True
    End If
End Function

Function is_Controller(columnName As String) As Boolean
    is_Controller = False
    If (columnName = "*RNC Name" Or columnName = getResByKey("A241") Or _
        columnName = "*BSC Name" Or columnName = getResByKey("A242")) Then
        is_Controller = True
    End If
End Function

'[Summaryת����ԭIUB�����]��List Sheet����ת��NodeB Sheetҳ
Public Sub ConvertList(ByVal strSheetName As String, ByVal startRow As Long, ByVal endRow As Long, ByVal IsEmptyDataSheet As Boolean, ByVal siteNameColNum As Long)
    Dim vListSht As Worksheet
    Dim vName As String
    'Dim vNameSite As String
    Dim vMocEndRowNum As Long
    Dim vNameEnd As String
    Dim vListColumnNum As Long

    If (Not IsSheetExist(strSheetName)) Then
        Exit Sub
    End If
    
    Set vListSht = Sheets(strSheetName)
    vListColumnNum = GetColNumByRowIndex(vListSht) - 1
    vName = GetCell(vListSht, g_TitleRowNo, siteNameColNum)
    'vNameSite = GetCell(vListSht, g_TitleRowNo, 2)
    vNameEnd = GetCell(vListSht, g_TitleRowNo, vListColumnNum)
    
    'ֻ����NodeB&BTS����
    If (Not is_SiteorController(vName)) Or (isTrasnPortSheet(strSheetName) And FILE_TYPE = "0") Then
        Exit Sub
    End If
    
    Call FormatFirst(vListSht, strSheetName, startRow, endRow, siteNameColNum)
    
    vMocEndRowNum = vListSht.range("a1048576").End(xlUp).row
    If (vMocEndRowNum <= g_TitleRowNo) Then
        vMocEndRowNum = g_TitleRowNo + 1
    End If
'
'    If is_Controller(vName) Then
        '������һ���ǹ������ƣ����ٿ���һ��'���һ�в��ǹ������ƣ�����������������
'        If isSharedControllerName(vNameEnd) Then
'            'Call copyControllerTitle_1(vListSht, strSheetName, startRow)
'        Else
'            Call copyControllerTitle(vListSht, strSheetName, startRow)
'        End If
'    Else
    Call CopySiteTitle(vListSht, strSheetName, startRow, siteNameColNum)
'        Call CopySiteData(vListSht, strSheetName, vMocEndRowNum, startRow, IsEmptyDataSheet)
'    End If
    
    Call setCopiedRangeFont(vListSht, startRow, endRow)
End Sub

Private Function isSharedControllerName(ByRef vName As String) As Boolean
    If (vName = "*NodeB Name" Or vName = getResByKey("A236")) _
        Or (vName = "*BTS Name" Or vName = getResByKey("A238")) Then
        isSharedControllerName = True
    Else
        isSharedControllerName = False
    End If
End Function

Private Function isLogicRncIdName(ByRef vName As String) As Boolean
    If (vName = "*Logical RNC ID" Or vName = getResByKey("LOGIC_RNC_ID")) Then
        isLogicRncIdName = True
    Else
        isLogicRncIdName = False
    End If
End Function

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
    
    Dim myRange As range
    Set myRange = vSiteSheet.range("B" + CStr(startRow) + ":" + C(vListColumnNum) + CStr(endRow))
    Call setRangeFont(myRange, "Arial")
    Call setRangeFontSize(myRange, 10)
End Sub

Public Sub setRangeFont(ByRef certainRange As range, ByRef fontName As String)
    certainRange.Font.name = fontName
End Sub

Public Sub setRangeWrap(ByRef certainRange As range, ByRef flag As Boolean)
    certainRange.WrapText = flag
End Sub

Public Sub setRangeFontSize(ByRef certainRange As range, ByRef fontSize As Long)
    certainRange.Font.Size = fontSize
End Sub

Public Sub FormatFirst(ByVal vListSht As Worksheet, ByVal strSheetName As String, ByVal SiteStartRow As Long, ByVal siteEndRow As Long, ByVal siteNameColNum As Long)
    Dim vSiteIndex As Long
    Dim vSiteName As String
    Dim vSiteSheet As Worksheet
    Dim vRowIndex As Long
    Dim vListColumnNum As Long
    Dim myRange As range
    Dim vName As String
    Dim vNameSite As String
    Dim vNameEnd As String
    Dim vNameRncId As String
    
    'For vSiteIndex = 0 To UBound(ChosenSiteArray)
        vSiteName = ChosenSiteArray(0)
        If ("" = vSiteName) Then
            Exit Sub
        End If
        
        Set vSiteSheet = Sheets(vSiteName)
        vListColumnNum = GetColNumByRowIndex(vListSht) - 1
        vName = GetCell(vListSht, g_TitleRowNo, 1)
        vNameEnd = GetCell(vListSht, g_TitleRowNo, vListColumnNum)
        '����ǿ�������������һ��
        If is_Controller(vName) Then
            '������һ���ǹ������ƣ����ٿ���һ��'���һ�в��ǹ������ƣ���������������
            If isSharedControllerName(vNameEnd) Then
                vListColumnNum = vListColumnNum - 1
            Else
                vListColumnNum = vListColumnNum
            End If
        End If
        '���Ӷ���UMTS 6910�汾�߼�RNCID�е��ж�
        vNameRncId = GetCell(vListSht, g_TitleRowNo, 3)
        If isLogicRncIdName(vNameRncId) Then
            vListColumnNum = vListColumnNum - 1
        Else
            vListColumnNum = vListColumnNum
        End If
        '���Ӷ��ڻ�վ�����е��жϣ������BTSNAME����NODEBNAME����������һ��
        vNameSite = GetCell(vListSht, g_TitleRowNo, siteNameColNum)
        If isSharedControllerName(vNameSite) Then
            vListColumnNum = vListColumnNum - 1
        Else
            vListColumnNum = vListColumnNum
        End If
        vSiteSheet.Activate
        Set myRange = vSiteSheet.range("A" + CStr(SiteStartRow) + ":" + C(vListColumnNum) + CStr(siteEndRow))
        Call AddListRefSub.setRangeBoard(myRange)

        myRange.Select
        With Selection.Interior
            .colorIndex = xlColorIndexNone
        End With
        vSiteSheet.range("A" + CStr(SiteStartRow) + ":" + "A" + CStr(siteEndRow)).Merge
    'Next vSiteIndex
End Sub

'Public Sub copyControllerTitle_1(ByVal vListSht As Worksheet, ByVal strSheetName As String, ByVal SiteStartRow As Long)
'    Dim vSiteIndex As Long
'    Dim vListColumnNum As Long
'    Dim vSiteName As String
'    Dim vSiteSheet As Worksheet
'    Dim myRange As range
'
'    vListColumnNum = GetColNumByRowIndex(vListSht) - 1
'    vSiteName = ChosenSiteArray(0)
'    If ("" = vSiteName) Then
'        Exit Sub
'    End If
'    Call WriteDatabyRow(vListSht, vSiteName, g_TitleRowNo, vListColumnNum - 1, SiteStartRow)
'    Set vSiteSheet = Sheets(vSiteName)
'    Call WriteSheetName(vSiteSheet, SiteStartRow, strSheetName)
'    '�رձ�ע��С��������ʾˢ���Ч��
'    Set myRange = vSiteSheet.range("B" + CStr(SiteStartRow) + ":" + C(vListColumnNum - 1) + CStr(SiteStartRow))
'    Call refreshComment(myRange)
'End Sub

'Public Sub copyControllerTitle(ByVal vListSht As Worksheet, ByVal strSheetName As String, ByVal SiteStartRow As Long)
'    Dim vSiteIndex As Long
'    Dim vListColumnNum As Long
'    Dim vSiteName As String
'    Dim vSiteSheet As Worksheet
'    Dim myRange As range
'
'    vListColumnNum = GetColNumByRowIndex(vListSht) - 1
'    vSiteName = ChosenSiteArray(0)
'    If ("" = vSiteName) Then
'        Exit Sub
'    End If
'    Call WriteDatabyRow(vListSht, vSiteName, g_TitleRowNo, vListColumnNum, SiteStartRow)
'    Set vSiteSheet = Sheets(vSiteName)
'    Call WriteSheetName(vSiteSheet, SiteStartRow, strSheetName)
'    '�رձ�ע��С��������ʾˢ���Ч��
'    Set myRange = vSiteSheet.range("B" + CStr(SiteStartRow) + ":" + C(vListColumnNum) + CStr(SiteStartRow))
'    Call refreshComment(myRange)
'End Sub

'Public Sub CopyControllerData(ByVal vListSht As Worksheet, ByVal strSheetName As String, ByVal vMocEndRowNum As Long, ByVal SiteStartRow As Long, ByVal IsEmptyDataSheet As Boolean)
'    Dim vSiteIndex As Long
'    Dim vSiteName As String
'    Dim vSiteSheet As Worksheet
'    Dim vSiteLastRow As Long
'    Dim vListColumnNum As Long
'    Dim vRowIndex As Long
'    Dim vSiteList As String
'    Dim myRange As range
'
'    '���һ����Site List,����Ҫ����
'    vListColumnNum = GetColNumByRowIndex(vListSht) - 1
'    For vSiteIndex = 0 To UBound(ChosenSiteArray)
'        vSiteName = ChosenSiteArray(vSiteIndex)
'        If ("" = vSiteName) Then
'            Exit Sub
'        End If
'
'        Set vSiteSheet = Sheets(vSiteName)
'        vSiteLastRow = SiteStartRow
'        For vRowIndex = g_TitleRowNo To vMocEndRowNum
'            vSiteList = vListSht.Cells(vRowIndex, vListColumnNum)
'            If (vRowIndex = g_TitleRowNo) Then
'                'Call WriteDatabyRow(vListSht, vSiteName, vRowIndex, vListColumnNum - 1, vSiteLastRow)
'                'Call WriteSheetName(vSiteSheet, vSiteLastRow, strSheetName)
'
'                '�رձ�ע��С��������ʾˢ���Ч��
'                'Set myRange = vSiteSheet.range("B" + CStr(vSiteLastRow) + ":" + C(vListColumnNum - 1) + CStr(vSiteLastRow))
'                'Call refreshComment(myRange)
'                vSiteLastRow = vSiteLastRow + 1
'            ElseIf IsSharedbySite(vSiteName, vSiteList, strSheetName, vListColumnNum) Or InStr(vSiteName, "~") <> 0 Then
'                Call WriteDatabyRow(vListSht, vSiteName, vRowIndex, vListColumnNum - 1, vSiteLastRow)
'                vSiteLastRow = vSiteLastRow + 1
'            End If
'            If (IsEmptyDataSheet) Then
'                vRowIndex = vMocEndRowNum + 1
'            End If
'        Next vRowIndex
'
'    Next vSiteIndex
'End Sub

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


Public Sub CopySiteData(ByVal vListSht As Worksheet, ByVal strSheetName As String, ByVal vMocEndRowNum As Long, ByVal SiteStartRow As Long, ByVal IsEmptyDataSheet As Boolean, ByVal siteNameColNum As Long)
    Dim vRowIndex As Long
    Dim vNodeBName As String
    Dim vNodeBSheet As Worksheet
    Dim vSiteSheetName As String
    Dim vSiteLastRow As Long
    Dim vLastNodeName As String
    Dim vListColumnNum As Long
    Dim siteLastRowMap As New CMap '��ÿ����վΪKeyֵ�������һ�п���Ϊ����
    'Copy title first
    'Call CopySiteTitle(vListSht, strSheetName, SiteStartRow)
    
    If (IsEmptyDataSheet) Then
        Exit Sub
    End If
    
    vLastNodeName = ""
    vListColumnNum = GetColNumByRowIndex(vListSht) - 1
    vSiteLastRow = SiteStartRow + 1
    For vRowIndex = g_TitleRowNo + 1 To vMocEndRowNum
        vNodeBName = Trim(GetCell(vListSht, vRowIndex, siteNameColNum))
        If ("" <> vNodeBName) Then
            vSiteSheetName = GetSiteSheetName(vNodeBName, strSheetName, siteNameColNum)
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
                '��������
                If isOperationColumn(vListSht) Then
                    Call WriteDatabyRowWithOper(vListSht, vSiteSheetName, vRowIndex, vListColumnNum, vSiteLastRow, siteNameColNum)
                Else
                    Call WriteDatabyRow(vListSht, vSiteSheetName, vRowIndex, vListColumnNum, vSiteLastRow, siteNameColNum)
                End If
                'Call WriteDatabyRow(vListSht, vSiteSheetName, vRowIndex, vListColumnNum, vSiteLastRow, siteNameColNum)
                vSiteLastRow = vSiteLastRow + 1
                
                Call siteLastRowMap.SetAt(vSiteSheetName, CStr(vSiteLastRow)) '�������һ�п��еļ�¼
                
'                vLastNodeName = vSiteSheetName
            End If
        End If
    Next vRowIndex
End Sub

Private Function isOperationColumn(ByVal vListSht As Worksheet)
    If vListSht.Cells(2, 1) = "Operation" Or vListSht.Cells(2, 1) = getResByKey("OPERATION") Then
        isOperationColumn = True
    Else
        isOperationColumn = False
    End If
End Function

Public Sub CopySiteTitle(ByVal vListSht As Worksheet, ByVal strSheetName As String, ByVal SiteStartRow As Long, ByVal siteNameColNum As Long)
    Dim vSiteIndex As Long
    Dim vListColumnNum As Long
    Dim vSiteName As String
    Dim vSiteSheet As Worksheet
    Dim myRange As range
    
    vListColumnNum = GetColNumByRowIndex(vListSht) - 1
    vSiteName = ChosenSiteArray(0)
    If ("" = vSiteName) Then
        Exit Sub
    End If
    '�ж������Ƿ�Ϊ��������
    If isOperationColumn(vListSht) Then
        Call WriteDatabyRowWithOper(vListSht, vSiteName, g_TitleRowNo, vListColumnNum, SiteStartRow, siteNameColNum)
    Else
        Call WriteDatabyRow(vListSht, vSiteName, g_TitleRowNo, vListColumnNum, SiteStartRow, siteNameColNum)
    End If
        Set vSiteSheet = Sheets(vSiteName)
        Call WriteSheetName(vSiteSheet, SiteStartRow, strSheetName)
        
        '�رձ�ע��С��������ʾˢ���Ч��
        Set myRange = vSiteSheet.range("B" + CStr(SiteStartRow) + ":" + C(vListColumnNum) + CStr(SiteStartRow))
        Call refreshComment(myRange)
    'Next vSiteIndex
End Sub

Private Sub refreshComment(ByVal myRange As range)
    On Error Resume Next
    Dim cell As range
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
    
    Dim GroupName As String
    Dim columnName As String
    Dim mocSheet As Worksheet
    Dim columnNum As Long
    Dim findRowNo As Long
    
    Set mocSheet = Sheets(mocSheetName)
    GroupName = GetCell(mocSheet, 1, MocColmunNo)
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
    GetSiteSheetName = Trim(GetCell(Sheets(getResByKey("BaseTransPort")), findRowNo, 2))
    If Len(GetSiteSheetName) > 31 Then
        If Trim(GetCell(Sheets(getResByKey("BaseTransPort")), 2, 2)) = "Sheet Name for Site" Or Trim(GetCell(Sheets(getResByKey("BaseTransPort")), 2, 2)) = getResByKey("A245") Then
           GetSiteSheetName = Trim(GetCell(Sheets(getResByKey("BaseTransPort")), findRowNo, 2))
        End If
    End If
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

'��MOC(Sheetҳ����)д�뵽NodeB��
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

Public Sub WriteDatabyRowWithOper(ByVal vListSht As Worksheet, ByVal vSiteName As String, ByVal vRowIndex As Long, ByVal vListColumnNum As Long, ByVal vSiteWrtRow As Long, ByVal siteNameColNum As Long)
    Dim vSiteSheet As Worksheet
    Set vSiteSheet = Sheets(vSiteName)
    vSiteSheet.Activate
    vSiteSheet.Cells(vSiteWrtRow, 2).Select
    '�ȸ��Ʋ�������
    vListSht.range(C(1) + CStr(vRowIndex) + ":" + C(1) + CStr(vRowIndex)).Copy Destination:=vSiteSheet.Cells(vSiteWrtRow, 2)
    '�ٸ���������
    vListSht.range(C(siteNameColNum + 1) + CStr(vRowIndex) + ":" + C(vListColumnNum) + CStr(vRowIndex)).Copy Destination:=vSiteSheet.Cells(vSiteWrtRow, 3)
End Sub

'�����ݰ��п�����NodeB��
Public Sub WriteDatabyRow(ByVal vListSht As Worksheet, ByVal vSiteName As String, ByVal vRowIndex As Long, ByVal vListColumnNum As Long, ByVal vSiteWrtRow As Long, ByVal siteNameColNum As Long)
    'vListSht.Activate
    'vListSht.range("B" + CStr(vRowIndex) + ":" + C(vListColumnNum) + CStr(vRowIndex)).Select
    'Selection.Copy

    Dim vSiteSheet As Worksheet
    Set vSiteSheet = Sheets(vSiteName)
    vSiteSheet.Activate
    vSiteSheet.Cells(vSiteWrtRow, 2).Select
    vListSht.range(C(siteNameColNum + 1) + CStr(vRowIndex) + ":" + C(vListColumnNum) + CStr(vRowIndex)).Copy Destination:=vSiteSheet.Cells(vSiteWrtRow, 2)
    'ActiveSheet.Paste
End Sub


'��ȡSheetҳĳ�е��б�
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
'�������õ�������1->A��27->AA
'**********************************************************
Public Function C(iColumn As Long) As String
  If iColumn >= 257 Or iColumn < 0 Then
    C = ""
    Return
  End If
  
  Dim result As String
  Dim High, Low As Long
  
  High = Int((iColumn - 1) / 26)
  Low = iColumn Mod 26
  
  If High > 0 Then
    result = Chr(High + 64)
  End If

  If Low = 0 Then
    Low = 26
  End If
  
  result = result & Chr(Low + 64)
  C = result
End Function

'**********************************************************
'�������õ�������A->1��AA->27
'**********************************************************
Public Function D(ColumnStr As String) As Long
  If Len(ColumnStr) = 1 Then
    D = Int(ColumnStr) - 64
  ElseIf Len(ColumnStr) = 2 Then
    D = (Int(Left(ColumnStr, 1)) - 64) * 26 + (Int(Left(ColumnStr, 1)) - 64)
  End If
End Function

Public Function IsSheetExist(ByVal sheetName As String) As Boolean
    Dim SheetNum, SheetCount As Long 'SheetCountÿ��ԭʼ�����ļ���Sheetҳ����
    SheetCount = ActiveWorkbook.Worksheets.count   '���м���Sheetҳ
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

'����һ���������ֵ��
Public Function GetSheetUsedRows(sheet As Worksheet) As Long
    Dim maxRow As Long
    maxRow = 1
    
    Do While Trim(sheet.Cells(maxRow + 1, 1)) <> ""
        maxRow = maxRow + 1
    Loop
    
    GetSheetUsedRows = maxRow
End Function

'����һ���������ֵ��
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

'copy Patternҳ�󣬶�λ��NodeBҳ������title����
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
        PatternSheet.range("A2:" + C(PatternSheetEndColumn) + CStr(PatternSheetEndRow)).Select
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
    Dim GroupName As String
    Dim colName As String
    
    Vec = Split(value, "\")
    sheetName = Vec(0)
    GroupName = Vec(1)
    colName = Vec(2)

    Dim RealColumn As Long
    RealColumn = Get_RefCol(sheetName, 2, GroupName, colName)

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

'���ÿ��NodeB Sheetҳ������PatternLinkֵ��ˢ��
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

Public Function JudgingEmptyDataSheet(ByVal ListSheetName As String) As Boolean
    Dim sheet As Worksheet
    Set sheet = Worksheets(ListSheetName)
    JudgingEmptyDataSheet = sheet.Cells(3, 2) = "" ' And Sheet.Cells(3, 2) = ""
End Function

Public Sub writeData()
    Dim startRow As Long
    Dim endRow As Long
    startRow = 1
    Dim IsEmptyDataSheet As Boolean
    Dim Iter As Long
    Dim ListSheetName As String
    Dim sheetDef As Worksheet
    Dim startRowCol As Long
    Dim endRowCol As Long
    
    startRowCol = startRowColNumInShtDef
    endRowCol = endRowColNumInShtDef
    If startRowCol = -1 Or endRowCol = -1 Then Exit Sub
    
    Set sheetDef = Worksheets("SHEET DEF")
    For Iter = 2 To sheetDef.range("a1048576").End(xlUp).row
        If sheetDef.Cells(Iter, startRowCol).value <> "" And sheetDef.Cells(Iter, endRowCol).value <> "" Then
            startRow = CStr(sheetDef.Cells(Iter, startRowCol).value)
            endRow = CStr(sheetDef.Cells(Iter, endRowCol).value)
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
    '���жϵ�ǰҳǩ���ڵ���NODEBNAME����BTSNAME
    Dim isBtsName As Boolean
    Dim isNodeBName As Boolean
    Dim value As Integer
    isBtsName = False
    isNodeBName = False
    For value = 2 To vListSht.Cells(g_TitleRowNo, columns.count).End(xlToLeft).column
        If vListSht.Cells(g_TitleRowNo, value) = "*BTS Name" Or _
        vListSht.Cells(g_TitleRowNo, value) = getResByKey("A238") Then
            isBtsName = True
        ElseIf vListSht.Cells(g_TitleRowNo, value) = "*NodeB Name" Or _
        vListSht.Cells(g_TitleRowNo, value) = getResByKey("A236") Then
            isNodeBName = True
        End If
    Next
    '����sheetҳ���ƻ�ȡMoc����
    Dim mocNameByShtName As String
    mocNameByShtName = getMocNameByShtName(strSheetName)
    '��ȡBTSNAME����NODEBNAME�����к�
    Dim siteNameColNum As Long
    If isBtsName Then
        siteNameColNum = getColNum(strSheetName, 2, "BTSNAME", mocNameByShtName)
    ElseIf isNodeBName Then
        siteNameColNum = getColNum(strSheetName, 2, "NODEBNAME", mocNameByShtName)
    End If
    vName = GetCell(vListSht, g_TitleRowNo, siteNameColNum)
    'ֻ����NodeB&BTS����
    If (Not is_SiteorController(vName)) Or (isTrasnPortSheet(strSheetName) And FILE_TYPE = "0") Then
        Exit Sub
    End If
    
    vMocEndRowNum = vListSht.range("b1048576").End(xlUp).row
    If (vMocEndRowNum <= g_TitleRowNo) Then
        vMocEndRowNum = g_TitleRowNo + 1
    End If

'    If is_Controller(vName) Then
'        Call copyControllerTitle(vListSht, strSheetName, startRow)
'        Call CopyControllerData(vListSht, strSheetName, vMocEndRowNum, startRow, IsEmptyDataSheet)
'    Else
'        Call CopySiteTitle(vListSht, strSheetName, startRow)
    Call CopySiteData(vListSht, strSheetName, vMocEndRowNum, startRow, IsEmptyDataSheet, siteNameColNum)
'    End If
End Sub

Public Function getMocNameByShtName(ByVal sheetName As String)
    Dim mappingDef As Worksheet
    Dim m_rowNum As Long
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For m_rowNum = 2 To mappingDef.range("a1048576").End(xlUp).row
        If UCase(sheetName) = UCase(mappingDef.Cells(m_rowNum, 1).value) Then
            getMocNameByShtName = mappingDef.Cells(m_rowNum + 1, 4).value
            Exit For
        End If
    Next
End Function
Public Function getShtNameBtsOrNodeBList(ByRef shtNameList As String)
    Dim mappingDef As Worksheet
    Dim m_rowNum As Long
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For m_rowNum = 4 To mappingDef.range("a1048576").End(xlUp).row
        If "BTSNAME" = UCase(mappingDef.Cells(m_rowNum, 5).value) Or "NODEBNAME" = UCase(mappingDef.Cells(m_rowNum, 5).value) Then
            '�ų���������NODEBNAME��С������
            If mappingDef.Cells(m_rowNum, 4).value <> "LOCELL" And mappingDef.Cells(m_rowNum, 4).value <> "CELL" Then
                If shtNameList = "" Then
                    shtNameList = mappingDef.Cells(m_rowNum, 1).value
                Else
                    shtNameList = shtNameList + "," + mappingDef.Cells(m_rowNum, 1).value
                End If
            End If
        End If
    Next
End Function


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
    Dim shtNameList As String
    Dim shtNameArr() As String
    Call getShtNameBtsOrNodeBList(shtNameList)
    shtNameArr = Split(shtNameList, ",")
    
    For Iter = 0 To UBound(shtNameArr)
            'ListSheetName = sheetDef.Cells(Iter, 1).value
            If IsSheetExist(shtNameArr(Iter)) And (Not isBSTransPortSht(shtNameArr(Iter))) Then
                If FILE_TYPE = "1" Or Not isTrasnPortSheet(shtNameArr(Iter)) Then
                    Dim tmpSht As Worksheet
                    Set tmpSht = Worksheets(shtNameArr(Iter))
                    
                    '���жϵ�ǰҳǩ���ڵ���NODEBNAME����BTSNAME
                    Dim isBtsName As Boolean
                    Dim isNodeBName As Boolean
                    Dim value As Integer
                    isBtsName = False
                    isNodeBName = False
                    For value = 2 To tmpSht.Cells(g_TitleRowNo, columns.count).End(xlToLeft).column
                        If tmpSht.Cells(g_TitleRowNo, value) = "*BTS Name" Or _
                            tmpSht.Cells(g_TitleRowNo, value) = getResByKey("A238") Then
                            isBtsName = True
                        ElseIf tmpSht.Cells(g_TitleRowNo, value) = "*NodeB Name" Or _
                            tmpSht.Cells(g_TitleRowNo, value) = getResByKey("A236") Then
                            isNodeBName = True
                        End If
                    Next
                    '����sheet���ƻ�ȡMoc����
                    Dim mocName As String
                    mocName = getMocNameByShtName(shtNameArr(Iter))
                    '��ȡBTSNAME����NODEBNAME�����к�
                    Dim siteNameColNum As Long
                    If isBtsName Then
                        siteNameColNum = getColNum(shtNameArr(Iter), 2, "BTSNAME", mocName)
                    ElseIf isNodeBName Then
                        siteNameColNum = getColNum(shtNameArr(Iter), 2, "NODEBNAME", mocName)
                    End If
                    
                    If is_SiteorController(GetCell(tmpSht, g_TitleRowNo, siteNameColNum)) Then
                        IsEmptyDataSheet = JudgingEmptyDataSheet(shtNameArr(Iter))
                        If (IsEmptyDataSheet) Then
                            endRow = startRow + 1
                        Else
                            endRow = startRow + GetMaxCountPerSite(shtNameArr(Iter))
                        End If
                        
                        Call setBscRncBtsFlag(bscFlag, rncFlag, btsFlag, shtNameArr(Iter))
                        
                        '������ҵ��˿��������ҵ���վ��������һ�п���
                        If alreadyAddEmptyRow = False And (bscFlag = True Or rncFlag = True) And btsFlag = True Then
                            startRow = startRow + 1
                            endRow = endRow + 1
                            alreadyAddEmptyRow = True
                        End If
                        Call ConvertList(shtNameArr(Iter), startRow, endRow, IsEmptyDataSheet, siteNameColNum)
                        
                        'Application.DisplayAlerts = False
                        'Sheets(ListSheetName).Delete
                        'Application.DisplayAlerts = True
                        Dim sheetDefRow As Long
                        Call getShtDefRow(shtNameArr(Iter), sheetDefRow)
                        Dim startColumn As Long
                        startColumn = GetSheetUsedColumnsForRow(sheetDef, 1)
                        sheetDef.Cells(sheetDefRow, startColumn + 1) = CStr(startRow)
                        sheetDef.Cells(sheetDefRow, startColumn + 2) = CStr(endRow)
                        startRow = endRow + 1
                    End If
                End If
            End If
    Next
    'Call RefreshPatternValue 'ignore pattern process
    Call WriteSheeDefTitle
End Sub

Public Function getShtDefRow(ByVal shtName As String, ByRef sheetDefRow As Long)
    Dim sheetDef As Worksheet
    Dim m_rowNum As Long
    'Dim getShtDefRow As Long
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
    For m_rowNum = 2 To sheetDef.range("a1048576").End(xlUp).row
        If UCase(shtName) = UCase(sheetDef.Cells(m_rowNum, 1).value) Then
            sheetDefRow = m_rowNum
            Exit For
        Else
            sheetDefRow = -1
        End If
    Next
End Function

Private Sub setBscRncBtsFlag(ByRef bscFlag As Boolean, ByRef rncFlag As Boolean, ByRef btsFlag As Boolean, ByVal ListSheetName As String)
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
            .ColumnWidth = 30
            .WrapText = False
        End With
        '��������
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

Private Sub makeSheetNameCol(ByRef sheetNameCol As Collection)
    Dim ws As Worksheet
    Dim sheetName As String
    For Each ws In ThisWorkbook.Worksheets
        sheetName = UCase(ws.name)
        If Not Contains(sheetNameCol, sheetName) Then
            sheetNameCol.Add Item:=sheetName, key:=sheetName
        End If
    Next
End Sub

Private Function checkExistingSheetName() As Boolean
    checkExistingSheetName = True
    Call InitTemplateVersion
    
    '�õ�����ҳǩ����������key��item����ҳǩ���ƵĴ�д
    Dim sheetNameCol As New Collection
    Call makeSheetNameCol(sheetNameCol)
    
    Dim NodeBBeginRow As Long, NodeBEndRow As Long, rowNumber As Long
    NodeBBeginRow = g_NodeBBeginRow
    
    Dim baseStationSheet As Worksheet
    Set baseStationSheet = ThisWorkbook.Worksheets(getResByKey("BaseTransPort"))
    NodeBEndRow = baseStationSheet.range("A1048576").End(xlUp).row
    
    Dim siteName As String, ucaseSiteName As String
    
    Dim existingNeNameCol As New Collection
    
    Dim sitesNameCol As New Collection
    For rowNumber = NodeBBeginRow To NodeBEndRow
        siteName = Trim(baseStationSheet.range("A" & rowNumber).value)
        ucaseSiteName = UCase(siteName)
        If siteName = "" Then GoTo NextLoop
        If Contains(sheetNameCol, ucaseSiteName) Then
            existingNeNameCol.Add Item:=siteName, key:=siteName
        End If
NextLoop:
    Next rowNumber

    '���count>0��˵���л�վ����������ҳǩ��������������������ת��
    Dim errorMsg As String
    If existingNeNameCol.count <> 0 Then
        checkExistingSheetName = False
        errorMsg = getErrorMsg(existingNeNameCol)
    End If
End Function

Private Function getErrorMsg(ByRef existingNeNameCol As Collection) As String
    Dim maxNumber As Long, index As Long
    Dim errorMsg As String
    If existingNeNameCol.count > 10 Then
        maxNumber = 10
    Else
        maxNumber = existingNeNameCol.count
    End If
    
    For index = 1 To maxNumber
        errorMsg = connectIndividualString(errorMsg, existingNeNameCol(index), vbCrLf)
    Next
    
    If existingNeNameCol.count > 10 Then
        errorMsg = errorMsg & vbCrLf & getResByKey("MSG_TOO_LONG")
    End If
    
    errorMsg = getResByKey("SheetNameExist") & vbCrLf & errorMsg
    
    Call MsgBox(errorMsg, vbCritical, getResByKey("ErrorInfo"))
End Function

'ƴ���ַ�������
Private Function connectIndividualString(ByVal wholeString As String, ByVal eachString As String, Optional ByVal delimeter As String = ",") As String
    If wholeString = "" Then
        connectIndividualString = eachString
    Else
        connectIndividualString = wholeString & delimeter & eachString
    End If
End Function

Public Function InitChosenSiteArray() As Boolean
    InitChosenSiteArray = True
    'clear ChosenSiteArray
    Dim i As Long
    For i = 1 To MaxChosenSiteNum
        ChosenSiteArray(i) = ""
    Next i
    
    'fill in ChosenSiteArray
    Dim NodeBBaseSheet As Worksheet
    Set NodeBBaseSheet = Sheets(getResByKey("BaseTransPort"))
    
    Dim NodeBBeginRow As Long
    Dim NodeBEndRow As Long
    NodeBBeginRow = g_NodeBBeginRow
    NodeBEndRow = GetSheetUsedRows(Worksheets(getResByKey("BaseTransPort")))
    If NodeBEndRow > MaxChosenSiteNum Then
        '�����������վ����������ʾ�޷�ת�����˳�
        Call MsgBox(getResByKey("SitesExceedsMaxNumber"), vbOKOnly + vbExclamation, getResByKey("ErrorInfo"))
        InitChosenSiteArray = False
        Exit Function
        'NodeBEndRow = MaxChosenSiteNum
    ElseIf NodeBEndRow <= g_NodeBBeginRow - 1 Then 'if no sites has been found, abort with msg
        If bIsEng Then
            MsgBox "No site has been found.", vbOKOnly
        Else
            MsgBox getResByKey("A246"), vbOKOnly
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
                MsgBox "The name of [" + siteName + "] is the same as another NE name, please modify the name.", vbCritical
            Else
                MsgBox getResByKey("A247") + siteName + getResByKey("A248"), vbCritical
            End If
            InitChosenSiteArray = False
            Exit Function
'        ElseIf IsSheetExist(siteName) Then
'            If bIsEng Then
'                MsgBox "Sheet name[" + siteName + "] has already been occupied, and this NE will be ignored in this operation.", vbOKOnly
'            Else
'                MsgBox "ҳǩ����[" + siteName + "]�ѱ�ռ��,�˻�վ�ڱ��������н�������.", vbOKOnly
'            End If
'            NodeBBeginRow = NodeBBeginRow + 1
'            IgnoredSiteNum = IgnoredSiteNum + 1
        Else
            If Len(siteName) > 31 Then
              siteName = ForNewSiteName(siteName, NodeBBeginRow)
            End If
            ChosenSiteArray(NodeBBeginRow - g_NodeBBeginRow - IgnoredSiteNum) = Trim(siteName)
            If Not Contains(ChosenSiteNames, siteName) Then
                ChosenSiteNames.Add Item:=(Trim(siteName)), key:=(Trim(siteName))
            End If
            NodeBBeginRow = NodeBBeginRow + 1
        End If
    Loop
    
    '����Ƿ��վ����������Listҳǩ���������򲻽���ת��
    If checkExistingSheetName = False Then
        InitChosenSiteArray = False
        Exit Function
    End If
End Function

Function IsValidSiteName(siteName As String) As Boolean
  IsValidSiteName = True
  
  If Len(siteName) > 64 _
     Or InStr(siteName, "\") > 0 _
     Or InStr(siteName, "/") > 0 _
     Or InStr(siteName, ":") > 0 _
     Or InStr(siteName, "*") > 0 _
     Or InStr(siteName, "?") > 0 _
     Or InStr(siteName, Chr(34)) > 0 _
     Or InStr(siteName, "[") > 0 _
     Or InStr(siteName, "]") > 0 _
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
'        ChooseSiteForm.Caption = "��ѡ����Ҫ����������ͼ��վ"
'        ChooseSiteForm.OKButton.Caption = "ȷ ��"
'        ChooseSiteForm.CancelButton.Caption = "ȡ ��"
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
'            MsgBox "�Ҳ�����վ.", vbOKOnly
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
        Dim response
        response = MsgBox(getResByKey("IUBTips"), vbExclamation + vbOKCancel)
        If response = vbCancel Then
            Exit Sub
        End If
    
        Application.ScreenUpdating = False
        
        'Call SortListLine
        Call GenIubFormatReport_i
        
        Application.ScreenUpdating = True
    End If
End Sub

Sub GenIubFormatReportMute()
    Call InitTemplateVersion
    If InitChosenSiteArray Then
 
        Application.ScreenUpdating = False

        Call GenIubFormatReport_i(True)
        
        Application.ScreenUpdating = True
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
    
    Call setDataRowsWrap(tempSheetName)   '�������е��Զ�����ԭ��û�У�����Ϊ��������
    
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
    Dim dataRange As range
    '����Sheet Defҳ����ÿ��IUBҳǩ�еĿ��ж��������Զ�����
    For rowIndex = 2 To sheetDefSheet.range("A1048576").End(xlUp).row
        startRow = sheetDefSheet.range("D" & rowIndex).value
        endRow = sheetDefSheet.range("E" & rowIndex).value
        If startRow <> "" And endRow <> "" Then
            maxColumnNumber = iubSheet.range("XFD" & startRow).End(xlToLeft).column
            Set dataRange = iubSheet.range("B" & (startRow + 1) & ":" & C(maxColumnNumber) & endRow)
            Call setRangeWrap(dataRange, True)
        End If
    Next rowIndex
End Sub


Sub GenIubFormatReport_i(Optional mute As Boolean = False)
    '�����һ����վ��Ϊ�գ���ֱ���˳���������
    If ChosenSiteArray(0) = "" Then Exit Sub
    Call refreshStart
    Call InitSiteNameArray
    Call InitTemplateVersion
    
    '��һ��������NodeB Sheetҳ
    Call CreateNewSiteSheet

    '�ڶ������ݹ���������ҳ������ConvertList��ConvertPattern
    Call InvokeConversion

    Call FormatSiteSheet
    Call createIubStyleWorkSheet(ThisWorkbook, ThisWorkbook.Worksheets(ChosenSiteArray(0)))
    Call writeData
    Call ContructPositionInfo
    Call refreshEnd
    
    If mute = False Then
        If bIsEng Then
            MsgBox "Finish generating view.", vbOKOnly
        Else
            MsgBox getResByKey("A249"), vbOKOnly
        End If
    End If
End Sub

Public Sub InitTemplateVersion()
    bIsEng = getResByKey("Cover") = "Cover"
End Sub

Public Function GetBluePrintSheetName() As String '��ǰֻ֧��һ��
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
    Dim i As Long
    i = GetMaxCountPerSite(ThisWorkbook.ActiveSheet.name)
    MsgBox CStr(i), vbOKOnly
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
        
        'If Not Contains(ChosenSiteNames, siteName) Then
        If Not Contains(ChosenSiteNames, siteName) And Not Contains(ChosenSiteNames, GetSiteSheetName(siteName, ListSheet.name, 1)) Then
            GoTo NextLoop
        End If
        
        index = IndexInArray(SiteNames, siteName)
        If index = -1 Then
            SiteNames(SiteCount) = siteName
            MocCounts(SiteCount) = 1
            SiteCount = SiteCount + 1
        Else
            MocCounts(index) = MocCounts(index) + 1
        End If
NextLoop:
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

Public Function GetMaxCountPerSite(ByVal ListName As String) As Long
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
        If GetCell(sheet, 2, 2) <> "Sheet Name for Site" And GetCell(sheet, 2, 2) <> getResByKey("A250") Then
        range("B2").Select
        Selection.EntireColumn.Insert
        
        If bIsEng Then
            Cells(2, 2) = "Referenced Site"
        Else
            Cells(2, 2) = getResByKey("A251")
        End If
        Else
            range("C2").Select
            Selection.EntireColumn.Insert
        
            If bIsEng Then
                Cells(2, 3) = "Referenced Site"
            Else
                Cells(2, 3) = getResByKey("A252")
            End If
        End If
               
        range("A1").Select
        
        If Selection.count = 1 Then
            range("A1:B1").Select
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
'    LastRow = sheetDef.range("a1048576").End(xlUp).row
'
'    Dim Iter as long
'    Dim ListSheetName As String
'    For Iter = 2 To sheetDef.range("a1048576").End(xlUp).row
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


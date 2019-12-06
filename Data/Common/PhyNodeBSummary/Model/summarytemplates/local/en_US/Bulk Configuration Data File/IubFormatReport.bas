Attribute VB_Name = "IubFormatReport"
Option Explicit

Private Const maxSiteCount = 200
Private Const listShtGrpRow = 1
Private Const listShtAttrRow = 2

Private targetMocs As Collection
Private targetSiteNames As Collection

Public nodeNameColName As String
Public isGenMocView As Boolean

'function entrance
Public Sub GenIubFormatReport()
    Dim response
    response = MsgBox(getResByKey("IUBTips"), vbExclamation + vbOKCancel)
    If response = vbCancel Then
        Exit Sub
    End If
    
    Call initResources
    
    Dim mainSht As Worksheet

    Set mainSht = Sheets(GetMainSheetName)

    Call classifyMocs
    Call parseSiteNames(mainSht)
    
    Dim mocViewMgr As CMocViewManager
    Set mocViewMgr = New CMocViewManager
    
    If Not targetMocs Is Nothing And targetSiteNames.count <> 0 And targetMocs.count <> 0 Then
        Call mocViewMgr.genMocView(mainSht, targetSiteNames, targetMocs, getNeType, "", "")
    End If
    
    Call releaseResources
    MsgBox getResByKey("FinishGenMocView")
End Sub

Private Function parseSiteNames(mainSht As Worksheet) As Boolean
    parseSiteNames = True
    
    Dim shtMaxRow As Long
    shtMaxRow = getSheetUsedRows(mainSht)
    
    '如果超出最大基站个数，则提示无法转换，退出
    If shtMaxRow > maxSiteCount + 2 Then
        Call MsgBox(getResByKey("SitesExceedsMaxNumber"), vbOKOnly + vbExclamation, getResByKey("ErrorInfo"))
        parseSiteNames = False
        Exit Function
    ElseIf shtMaxRow = listShtAttrRow Then 'if no sites has been found, abort with msg
        MsgBox getResByKey("NoSiteFound"), vbOKOnly 'todo
        parseSiteNames = False
        Exit Function
    End If
    
    Dim NameCol As Long
    NameCol = siteNameColNum(mainSht)
    
    If targetSiteNames Is Nothing Then Set targetSiteNames = New Collection
    
    Dim rowIdx As Integer
    For rowIdx = listShtAttrRow + 1 To getSheetUsedRows(mainSht)
        Dim name As String
        name = Trim(mainSht.Cells(rowIdx, NameCol))
        If name <> "" Then targetSiteNames.Add Item:=name, key:=name
    Next
End Function

Private Sub classifyMocs()
    Dim shtDef As Worksheet
    Set shtDef = Worksheets("SHEET DEF")
    
    Dim shtNameColIdx As Integer
    Dim shtTypeColIdx As Integer
    shtNameColIdx = shtNameColNumInShtDef
    shtTypeColIdx = shtTypeColNumInShtDef
    
    If targetMocs Is Nothing Then Set targetMocs = New Collection
    
    Dim mocName As String
    Dim rowIdx As Integer
    For rowIdx = 2 To shtDef.Range("a65535").End(xlUp).row
        If UCase(shtDef.Cells(rowIdx, shtTypeColIdx)) = "LIST" Then
            mocName = shtDef.Cells(rowIdx, shtNameColIdx)
            targetMocs.Add Item:=mocName, key:=mocName
        End If
    Next
    
End Sub

'按第基站名列算最大有值行
Public Function getSheetUsedRows(sheet As Worksheet) As Long
    getSheetUsedRows = 0
    
    Dim NameCol As Long
    NameCol = siteNameColNum(sheet)
    If NameCol = -1 Then
        NameCol = controllerNameColNum(sheet)
    End If
    
    If NameCol = -1 Then NameCol = 1
    
    getSheetUsedRows = sheet.Range(C(NameCol) & "65536").End(xlUp).row
End Function

'按第一行算最大有值列
Public Function getSheetUsedColumnsByRow(sheet As Worksheet, ByVal rowIdx As Long) As Long
  getSheetUsedColumnsByRow = sheet.Range("IV" & rowIdx).End(xlToLeft).column
End Function

Public Function getSheetUsedColumns(sheet As Worksheet) As Long
  Dim MaxColumn As Long
  MaxColumn = 1
  
  Do While Trim(sheet.Cells(2, MaxColumn + 1)) <> ""
    MaxColumn = MaxColumn + 1
  Loop
  
  getSheetUsedColumns = MaxColumn
End Function

Private Sub initResources()
    Call releaseResources
    
    isGenMocView = True
    
    Set targetMocs = New Collection

    Set targetSiteNames = New Collection

    Select Case getNeType
        Case NeType_MRAT
            nodeNameColName = getResByKey("*Name")
        Case NeType_CBS
            nodeNameColName = getResByKey("*NEName")
        Case NeType_LTE
            nodeNameColName = getResByKey("*eNodeBName")
        Case NeType_UMTS
            nodeNameColName = getResByKey("*NodeBName")
        Case NeType_GSM
            nodeNameColName = getResByKey("*BTSName")
        Case NeType_USU
            nodeNameColName = getResByKey("*USUName")
        Case NeType_ICS
            nodeNameColName = getResByKey("*ICSName")
        Case NeType_5G
            nodeNameColName = getResByKey("*gNodeBName")
        Case Else
            nodeNameColName = getResByKey("*Name")
    End Select
End Sub

Private Sub releaseResources()
    Set targetMocs = Nothing
    Set targetSiteNames = Nothing
    isGenMocView = False
End Sub




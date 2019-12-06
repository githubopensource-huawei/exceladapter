VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBaseStationData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private neBoardStyleMap_ As CMap 'key: nename: boardsylename
Private boardStyleNeMap_ As CMap
Private FunctionNeMap_ As CMap

Private baseScenarioDataColl_ As Collection
Private Const customDelimiter_ As String = "\"  '用于间隔custom字段的GroupName和ColumnName，自定义工具时这是非法字段

Public Sub init()
    Call initNeBoardStyleMap
    Call initBaseCustomInfoMap
End Sub
Public Function neBoardStyleMap() As CMap
    Set neBoardStyleMap = neBoardStyleMap_
End Function
Public Function boardStyleNeMap() As CMap
    Set boardStyleNeMap = boardStyleNeMap_
End Function

Public Function functionNeMap() As CMap
    Set functionNeMap = FunctionNeMap_
End Function

Private Sub initNeBoardStyleMap()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(getResByKey("BaseTransPort"))
    
    Dim rowNumber As Long
    Dim neName As String
    Dim boardstyleColName As String
    Dim boardstyleColLeter As String
    Dim boardstyleName As String
    
    Dim neNameColLeter As String
    Dim neNameColName As String

    Dim gbtsfunctionNameStr As String
    Dim neNameStr As String
    Dim gbtsfunctionName As String
    Dim gbtsfunctionNameColLeter As String
    Dim srcTarNeMap As CMapValueObject
    
    Set srcTarNeMap = initSrcTargetNeMapInfo
    
    Call getFunctionNameColum(gbtsfunctionName)
    
    boardstyleColName = getResByKey("Board Style Name")
    boardstyleColLeter = getColnumNameLeter(ws, boardstyleColName)
    
    neNameColName = getNeNameColum
    neNameColLeter = getColnumNameLeter(ws, neNameColName)
    
    If boardstyleColLeter = "" Or neNameColLeter = "" Then Exit Sub
    
    gbtsfunctionNameColLeter = getColnumNameLeter(ws, gbtsfunctionName)
    
    For rowNumber = 3 To ws.range("A1048576").End(xlUp).row
        neName = ws.range("B" & rowNumber).value
        boardstyleName = ws.range(boardstyleColLeter & rowNumber).value
        If neName <> "" And boardstyleName <> "" Then
            Call neBoardStyleMap_.SetAt(neName, boardstyleName)
            Call boardStyleNeMap_.SetAt(boardstyleName, neName)
        End If
        
        If neNameStr = "" Then
            neNameStr = ws.range(neNameColLeter & rowNumber).value
        Else
            neNameStr = neNameStr + "," + ws.range(neNameColLeter & rowNumber).value
        End If
        
        If gbtsfunctionNameColLeter <> "" Then
            If gbtsfunctionNameStr = "" Then
                gbtsfunctionNameStr = ws.range(gbtsfunctionNameColLeter & rowNumber).value
            Else
                gbtsfunctionNameStr = gbtsfunctionNameStr + "," + ws.range(gbtsfunctionNameColLeter & rowNumber).value
            End If
        End If
    Next rowNumber
    
     If neNameStr <> "" Then
        Call FunctionNeMap_.SetAt("BASESTATION", neNameStr)
    End If
    
    If gbtsfunctionNameStr <> "" Then
        Call FunctionNeMap_.SetAt("GBTSFUNCTION", gbtsfunctionNameStr)
    End If
    
End Sub

Private Function getColnumNameLeter(ByRef ws As Worksheet, ByRef colnumName As String) As String
    Dim m_colNum As Long
    For m_colNum = 1 To ws.range("XFD2").End(xlToLeft).column
        If ws.Cells(2, m_colNum).value = colnumName Then
            getColnumNameLeter = getColStr(m_colNum)
            Exit Function
        End If
    Next
    getColnumNameLeter = ""
End Function

Private Sub Class_Initialize()
    Set neBoardStyleMap_ = New CMap
    Set boardStyleNeMap_ = New CMap
    Set FunctionNeMap_ = New CMap
    Set baseScenarioDataColl_ = New Collection
End Sub
Private Sub Class_Terminate()
    Set neBoardStyleMap_ = Nothing
    Set boardStyleNeMap_ = Nothing
    Set FunctionNeMap_ = Nothing
    Set baseScenarioDataColl_ = Nothing
End Sub

Private Sub getFunctionNameColum(ByRef gbtsfunctionNameCol As String)
    Dim sheetDef As Worksheet
    Dim index As Long
    Dim defCollection As New Collection
    Dim mappingdefgroupName As String
    Dim mappingdefcolumnName As String
    Dim mocName As String
    Dim attributeName As String

    Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For index = 2 To sheetDef.range("a1048576").End(xlUp).row
        mappingdefgroupName = sheetDef.Cells(index, 2)
        mappingdefcolumnName = sheetDef.Cells(index, 3)
        mocName = sheetDef.Cells(index, 4)
        attributeName = sheetDef.Cells(index, 5)
        If mocName = "BTS" And attributeName = "BTSNAME" Then
            gbtsfunctionNameCol = mappingdefcolumnName
        End If
    Next
End Sub

Private Function getNeNameColum() As String
    Dim sheetDef As Worksheet
    Dim index As Long
    Dim defCollection As New Collection
    Dim mappingdefgroupName As String
    Dim mappingdefcolumnName As String
    Dim mocName As String
    Dim attributeName As String
    
    getNeNameColum = ""
    Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For index = 2 To sheetDef.range("a1048576").End(xlUp).row
        mappingdefgroupName = sheetDef.Cells(index, 2)
        mappingdefcolumnName = sheetDef.Cells(index, 3)
        mocName = sheetDef.Cells(index, 4)
        attributeName = sheetDef.Cells(index, 5)
        If mocName = "BTS" And attributeName = "BTSNAME" Then
            getNeNameColum = mappingdefcolumnName
            Exit For
        End If
    Next
End Function

Public Function baseScenarioDataColl() As Collection
    Set baseScenarioDataColl = baseScenarioDataColl_
End Function

Private Sub initBaseCustomInfoMap()
    Dim sheet As Worksheet
    Dim maxRow As Long
    Set sheet = ThisWorkbook.Worksheets(getResByKey("BaseTransPort"))
    maxRow = sheet.range("a1048576").End(xlUp).row
    
    '获取基站传输页签基站名称列
    Dim btsColNum As Long
    btsColNum = getTransBTSNameCol(getResByKey("BaseTransPort"))
    '获取单板样式名称列
    Dim brdColNum As Long
    brdColNum = getColNum(sheet.name, 2, "BRDSTYLE", "BTS")
    
    '获取场景列
    Dim scenarioColNumArr() As String
    Dim scenarioColNumStr As String
    scenarioColNumStr = getScenarioColNum(sheet.name, 2, CUSTOM_SCENARIO_ATTR_NAME, CUSTOM_SCENARIO_MOC_NAME)
    If scenarioColNumStr = "" Then Exit Sub
    scenarioColNumArr = Split(scenarioColNumStr, ",")
    
    Dim rowIndex As Long
    Dim sceIndex As Long
    Dim temSceIndex As Long
    Dim temBtsName As String
    Dim temBrdShtName As String
    Dim temGroupName As String
    Dim temColName As String
    Dim temScenarioStr As String
    Dim temScenarioArr() As String
    Dim temSecnario As String
    '遍历基站传输页，初始化场景化数据
    For rowIndex = 3 To maxRow
        temBtsName = sheet.Cells(rowIndex, btsColNum)
        temBrdShtName = sheet.Cells(rowIndex, brdColNum)
        For sceIndex = LBound(scenarioColNumArr) To UBound(scenarioColNumArr)
            temGroupName = sheet.Cells(1, val(scenarioColNumArr(sceIndex)))
            temColName = sheet.Cells(2, val(scenarioColNumArr(sceIndex)))
            temScenarioStr = sheet.Cells(rowIndex, val(scenarioColNumArr(sceIndex)))
            temScenarioArr = Split(temScenarioStr, ",")
            For temSceIndex = LBound(temScenarioArr) To UBound(temScenarioArr)
                temSecnario = temScenarioArr(temSceIndex)
                '初始化场景数据
                Call prepareScenarioData(temBtsName, temBrdShtName, temGroupName, temColName, temSecnario)
            Next
        Next
    Next
    
End Sub
'数据结构格式为：collection(基站名称#BoardStyle，collection(Group\column，collection(场景号)))
Private Sub prepareScenarioData(ByRef btsName As String, ByRef BrdShtName As String, ByRef groupName As String, ByRef ColName As String, ByRef Secnario As String)
    If btsName = "" Or BrdShtName = "" Or groupName = "" Or ColName = "" Or Secnario = "" Then
        Exit Sub
    End If
    '将（基站名称和单板样式页签名称）组合起来当做主键
    Dim btsBrdKeyValue As String
    Dim grpColKeyValue As String
    btsBrdKeyValue = btsName + "#" + BrdShtName
    grpColKeyValue = groupName + customDelimiter_ + ColName
    
    If Contains(baseScenarioDataColl_, btsBrdKeyValue) Then
        Dim groupColColl As Collection
        Set groupColColl = baseScenarioDataColl_(btsBrdKeyValue)
        If Contains(groupColColl, grpColKeyValue) Then
            Dim secnarioColl As Collection
            Set secnarioColl = groupColColl(grpColKeyValue)
            If Not Contains(secnarioColl, Secnario) Then
                secnarioColl.Add Item:=Secnario
            End If
        Else
            Dim temSecnarioColl As Collection
            Set temSecnarioColl = New Collection
            temSecnarioColl.Add Item:=Secnario
            
            groupColColl.Add Item:=temSecnarioColl, key:=grpColKeyValue
        End If
    Else
        Dim temScenarioColl_1 As Collection
        Set temScenarioColl_1 = New Collection
        temScenarioColl_1.Add Item:=Secnario
        
        Dim temGrpColColl As Collection
        Set temGrpColColl = New Collection
        temGrpColColl.Add Item:=temScenarioColl_1, key:=grpColKeyValue
        
        baseScenarioDataColl_.Add Item:=temGrpColColl, key:=btsBrdKeyValue
    End If
End Sub

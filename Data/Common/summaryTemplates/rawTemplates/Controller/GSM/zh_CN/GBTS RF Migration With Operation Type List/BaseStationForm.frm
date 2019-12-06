VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BaseStationForm 
   Caption         =   "Select Base Station"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6525
   OleObjectBlob   =   "BaseStationForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "BaseStationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub OKButton_Click()
    Dim baseStationName As String
    Dim CellSheetName As String
    baseStationName = Me.BaseStationList.value
    CellSheetName = ActiveSheet.name
    If IsGBTSTemplate() Then
        Call AddTrxBinds(baseStationName, CellSheetName)
    Else
        Call AddSectorEqm(baseStationName)
    End If
    
    Unload Me
End Sub

'从「当前激活页获得基站名称」
Private Sub Set_BaseStation_Related()
    Dim rowNum As Long
    Dim maxRow As String
    Dim nowSelection As Range
    Dim index As Long
    Dim baseStationCollection As Collection
    Dim baseStationName As Variant
    Dim selectBtsName As String
    Set baseStationCollection = New Collection
    BaseStationList.Clear
    
    Set nowSelection = Selection
    selectBtsName = ActiveSheet.Cells(nowSelection.row, 1).value
    If IsGBTSTemplate() Then
        selectBtsName = ActiveSheet.Cells(nowSelection.row, getGTRXBTSNameCol(ActiveSheet.name)).value
    End If
    
    If CELL_TYPE = 0 Or CELL_TYPE = 4 Then
        baseStationCollection.Add (selectBtsName)
    Else
        maxRow = ActiveSheet.Range("b1048576").End(xlUp).row
        For rowNum = 3 To maxRow
            baseStationName = ActiveSheet.Cells(rowNum, 1).value
            '@gbts
            If IsGBTSTemplate() Then
                baseStationName = ActiveSheet.Cells(rowNum, getGTRXBTSNameCol(ActiveSheet.name)).value
            End If
    
            If existInCollection(baseStationName, baseStationCollection) = False And Trim(baseStationName) <> "" Then
                baseStationCollection.Add (baseStationName)
            End If
        Next
    End If
    
    For Each baseStationName In baseStationCollection
        If Trim(baseStationName) <> "" Then
            BaseStationList.AddItem (baseStationName)
        End If
            
    Next
    
    If baseStationCollection.count <> 0 Then
        Me.BaseStationList.ListIndex = getIndexInCollection(selectBtsName, baseStationCollection)
    End If
End Sub


Private Function getIndexInCollection(strValue As Variant, strCollection As Collection) As Long
    Dim sItem As Variant
    Dim index As Long
    index = 0
    For Each sItem In strCollection
        If sItem = strValue Then
            getIndexInCollection = index
            Exit Function
        End If
        index = index + 1
    Next
    getIndexInCollection = 0
End Function


Private Sub initBaseStationScenario()
    Dim sheet As Worksheet
    Dim maxRow As Long
    Set sheet = ThisWorkbook.Worksheets(getResByKey("BaseTransPort"))
    maxRow = sheet.Range("a1048576").End(xlUp).row
    '初始化场景集合
    Set scenaioCollData = New Collection
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
    '遍历基站传输页，初始化场景化数据，数据结构格式为collection(单板样式页签，collection(组，collection(列，collection(场景号))))
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

'数据结构格式为collection(基站名称#单板样式页签，collection(组，collection(列，collection(场景号))))
Private Sub prepareScenarioData(ByRef btsName As String, ByRef BrdShtName As String, ByRef groupName As String, ByRef ColName As String, ByRef Secnario As String)
    If btsName = "" Or BrdShtName = "" Or groupName = "" Or ColName = "" Or Secnario = "" Then
        Exit Sub
    End If
    '将（基站名称和单板样式页签名称）组合起来当做主键
    Dim keyValue As String
    keyValue = btsName + "#" + BrdShtName
    If Contains(scenaioCollData, keyValue) Then
        Dim groupColl As Collection
        Set groupColl = scenaioCollData(keyValue)
        If Contains(groupColl, groupName) Then
            Dim columnColl As Collection
            Set columnColl = groupColl(groupName)
            If Contains(columnColl, ColName) Then
                Dim secnarioColl As Collection
                Set secnarioColl = columnColl(ColName)
                If Not Contains(secnarioColl, Secnario) Then
                    secnarioColl.Add Item:=Secnario
                End If
            Else
                Dim temSecnarioColl As Collection
                Set temSecnarioColl = New Collection
                
                temSecnarioColl.Add Item:=Secnario
                columnColl.Add Item:=temSecnarioColl, key:=ColName
            End If
        Else
            Dim temSecnarioColl_1 As Collection
            Set temSecnarioColl_1 = New Collection
            temSecnarioColl_1.Add Item:=Secnario
            
            Dim temColumnColl As Collection
            Set temColumnColl = New Collection
            temColumnColl.Add Item:=temSecnarioColl_1, key:=ColName
            
            groupColl.Add Item:=temColumnColl, key:=groupName
        End If
    Else
        Dim temSecnarioColl_2 As Collection
        Set temSecnarioColl_2 = New Collection
        temSecnarioColl_2.Add Item:=Secnario
        
        Dim temColumnColl_1 As Collection
        Set temColumnColl_1 = New Collection
        temColumnColl_1.Add Item:=temSecnarioColl_2, key:=ColName
        
        Dim temGroupColl As Collection
        Set temGroupColl = New Collection
        temGroupColl.Add Item:=temColumnColl_1, key:=groupName
        
        scenaioCollData.Add Item:=temGroupColl, key:=keyValue
    End If
End Sub

Private Sub UserForm_Initialize()
    Call Upt_Desc
    Call Set_BaseStation_Related
    '初始化基站页签场景
    Call initBaseStationScenario
End Sub

Private Sub Upt_Desc()
    BaseStationForm.Caption = getResByKey("BaseStationForm.Caption")
    BaseStationNameBox.Caption = getResByKey("BaseStationNameBox.Caption")
    OKButton.Caption = getResByKey("OKButton.Caption")
    CancelButton.Caption = getResByKey("CancelButton.Caption")
End Sub


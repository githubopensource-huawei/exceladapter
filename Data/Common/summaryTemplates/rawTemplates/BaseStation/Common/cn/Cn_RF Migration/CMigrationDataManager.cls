VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMigrationDataManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private targetSourceNeMap_ As CMapValueObject
Private baseStationNeMap_ As CMapValueObject
Private cellNeNameColumMap_ As CMap

Public Sub init()
    Call initCellsheetNeNameColMap
    Call initBaseStationNeMap
    Call inittargetSourceNeMap
End Sub
Public Function targetSourceNeMap() As CMapValueObject
    Set targetSourceNeMap = targetSourceNeMap_
End Function
Public Function baseStationNeMap() As CMapValueObject
    Set baseStationNeMap = baseStationNeMap_
End Function
Public Function cellNeNameColumMap() As CMap
    Set cellNeNameColumMap = cellNeNameColumMap_
End Function

Private Sub initBaseStationNeMap()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(getResByKey("BaseTransPort"))
    
    Dim functionNeMap As CMap
    Dim rowNumber As Long
    Dim neName As String
    Dim neNameColLeter As String
    Dim neNameColName As String
    Dim enodebfuntionName As String
    Dim nodebfunctionName As String
    Dim gbtsfunctionName As String
    Dim enodebfuntionNameColLeter As String
    Dim nodebfunctionNameColLeter As String
    Dim gbtsfunctionNameColLeter As String
    
    Call getFunctionNameColum(enodebfuntionName, nodebfunctionName, gbtsfunctionName)

    neNameColName = getNeNameColum
    neNameColLeter = getColnumNameLeter(ws, neNameColName)
    
    If neNameColLeter = "" Then Exit Sub
    
    enodebfuntionNameColLeter = getColnumNameLeter(ws, enodebfuntionName)
    nodebfunctionNameColLeter = getColnumNameLeter(ws, nodebfunctionName)
    gbtsfunctionNameColLeter = getColnumNameLeter(ws, gbtsfunctionName)
    
    For rowNumber = 3 To ws.Range("A1048576").End(xlUp).row
        neName = ws.Range(neNameColLeter & rowNumber).value
        Set functionNeMap = New CMap
        If enodebfuntionNameColLeter <> "" Then Call functionNeMap.SetAt("eNodeB", ws.Range(enodebfuntionNameColLeter & rowNumber).value)
        If nodebfunctionNameColLeter <> "" Then Call functionNeMap.SetAt("NodeB", ws.Range(nodebfunctionNameColLeter & rowNumber).value)
        If gbtsfunctionNameColLeter <> "" Then Call functionNeMap.SetAt("BTS", ws.Range(gbtsfunctionNameColLeter & rowNumber).value)
        If enodebfuntionNameColLeter = "" And nodebfunctionNameColLeter = "" And gbtsfunctionNameColLeter = "" Then Call functionNeMap.SetAt("BASESTATION", neName)
        If neName <> "" Then Call baseStationNeMap_.SetAt(neName, functionNeMap)
    Next rowNumber
      
End Sub

Private Sub inittargetSourceNeMap()
    Dim targetColNum As Long
    Dim targetNeName As String
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(getResByKey("RF Migration NEs Relationship"))

    Dim index As Long
    Dim onerowSrcNeMap As CMap
    Dim migrationNemap As CMigrationNeMap
    Set migrationNemap = New CMigrationNeMap
    
    targetColNum = getTargetNeClounmNum(ws)
    For index = 4 To ws.Range(getColStr(targetColNum) & "1048576").End(xlUp).row
        targetNeName = ws.Cells(index, targetColNum).value
        Set onerowSrcNeMap = migrationNemap.RelationNeColumnValueMap(index)
        If Not targetSourceNeMap_.hasKey(targetNeName) Then Call targetSourceNeMap_.SetAt(targetNeName, onerowSrcNeMap)
    Next
End Sub

Private Sub initCellsheetNeNameColMap()
    Dim sheetDef As Worksheet
    Dim index As Long
    Dim mappingdefgroupName As String
    Dim mappingdefcolumnName As String
    Dim mocName As String
    Dim attributeName As String
    Dim sheetName As String
    
    Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For index = 2 To sheetDef.Range("a1048576").End(xlUp).row
        sheetName = sheetDef.Cells(index, 1)
        mappingdefgroupName = sheetDef.Cells(index, 2)
        mappingdefcolumnName = sheetDef.Cells(index, 3)
        mocName = sheetDef.Cells(index, 4)
        attributeName = sheetDef.Cells(index, 5)
        
        If sheetName = "GSM Cell" Or sheetName = getResByKey("A177") Then
            If (mocName = "GBTSFUNCTION" And attributeName = "GBTSFUNCTIONNAME") Or (mocName = "NE" And attributeName = "NENAME") Then
                Call cellNeNameColumMap_.SetAt("BTS", mappingdefgroupName + "," + mappingdefcolumnName)
            End If
        End If
        
        If sheetName = "UMTS Cell" Or sheetName = getResByKey("A178") Then
            If (mocName = "NODEBFUNCTION" And attributeName = "NODEBFUNCTIONNAME") Or (mocName = "NE" And attributeName = "NENAME") Then
                Call cellNeNameColumMap_.SetAt("NodeB", mappingdefgroupName + "," + mappingdefcolumnName)
            End If
        End If
        
        If sheetName = "LTE Cell" Or sheetName = getResByKey("A179") Then
            If (mocName = "eNodeBFunction" And attributeName = "eNodeBFunctionName") Or (mocName = "NE" And attributeName = "NENAME") Then
                Call cellNeNameColumMap_.SetAt("eNodeB", mappingdefgroupName + "," + mappingdefcolumnName)
            End If
        End If
    Next
End Sub

Private Function getColnumNameLeter(ByRef ws As Worksheet, ByRef colnumName As String) As String
    Dim m_colNum As Long
    For m_colNum = 1 To ws.Range("XFD2").End(xlToLeft).column
        If ws.Cells(2, m_colNum).value = colnumName Then
            getColnumNameLeter = getColStr(m_colNum)
            Exit Function
        End If
    Next
    getColnumNameLeter = ""
End Function

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
    For index = 2 To sheetDef.Range("a1048576").End(xlUp).row
        mappingdefgroupName = sheetDef.Cells(index, 2)
        mappingdefcolumnName = sheetDef.Cells(index, 3)
        mocName = sheetDef.Cells(index, 4)
        attributeName = sheetDef.Cells(index, 5)
        If mocName = "NE" And attributeName = "NENAME" Then
            getNeNameColum = mappingdefcolumnName
            Exit For
        End If
    Next
End Function

Private Sub getFunctionNameColum(ByRef enodebfunctionNameCol As String, ByRef nodebfunctionNameCol As String, ByRef gbtsfunctionNameCol As String)
    Dim sheetDef As Worksheet
    Dim index As Long
    Dim defCollection As New Collection
    Dim mappingdefgroupName As String
    Dim mappingdefcolumnName As String
    Dim mocName As String
    Dim attributeName As String
    Dim sheetName As String
    

    Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For index = 2 To sheetDef.Range("a1048576").End(xlUp).row
        sheetName = sheetDef.Cells(index, 1)
        mappingdefgroupName = sheetDef.Cells(index, 2)
        mappingdefcolumnName = sheetDef.Cells(index, 3)
        mocName = sheetDef.Cells(index, 4)
        attributeName = sheetDef.Cells(index, 5)
        If sheetName = getResByKey("BaseTransPort") Then
            If mocName = "eNodeBFunction" And attributeName = "eNodeBFunctionName" Then enodebfunctionNameCol = mappingdefcolumnName
            If mocName = "NODEBFUNCTION" And attributeName = "NODEBFUNCTIONNAME" Then nodebfunctionNameCol = mappingdefcolumnName
            If mocName = "GBTSFUNCTION" And attributeName = "GBTSFUNCTIONNAME" Then gbtsfunctionNameCol = mappingdefcolumnName
        End If
    Next
End Sub


Private Sub Class_Initialize()
    Set targetSourceNeMap_ = New CMapValueObject
    Set baseStationNeMap_ = New CMapValueObject
    Set cellNeNameColumMap_ = New CMap
End Sub






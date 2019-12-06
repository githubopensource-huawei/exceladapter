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
Private baseCustomNameLetterMap_ As CMap 'key: CustomizeGroupName\CustomizeColumnName, value:  CustomizeColumnLetter
Private baseNeCustomInfoMap_ As CMapValueObject  'key：neName，value：CMap(key: CustomizeGroupName\CustomizeColumnName，customColValue)

Private Const customDelimiter_ As String = "\"  '用于间隔custom字段的GroupName和ColumnName，自定义工具时这是非法字段


Public Sub init()
    Call initBaseCustomInfoMap
    Call initNeBoardStyleMap
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

Public Function getNeNamebyFunctionName(ByVal functionName) As String
    Dim neNameVar As Variant
    Dim functionNames As String
    Dim functionNameArry() As String
    Dim index As Long
    
    getNeNamebyFunctionName = ""
    
    For Each neNameVar In FunctionNeMap_.KeyCollection
    'functionNames:NeName,FuntioncName1,FunctionName2,FunctionName3(至少有NeName)
    functionNames = FunctionNeMap_.GetAt(neNameVar)

        If InStr(functionNames, ",") <> 0 Then
            functionNameArry = Split(functionNames, ",")
            For index = LBound(functionNameArry) To UBound(functionNameArry)
                If functionNameArry(index) = functionName Then
                    getNeNamebyFunctionName = CStr(neNameVar)
                    Exit Function
                End If
            Next
        Else
            If functionName = CStr(neNameVar) Then
                getNeNamebyFunctionName = functionName
                Exit Function
            End If
        End If
    Next
End Function

Public Function baseNeCustomInfoMap() As CMapValueObject
    Set baseNeCustomInfoMap = baseNeCustomInfoMap_
End Function

'根据NeName找定制场景信息
Public Function getCustomInfoMapByNeName(ByVal neName As String, ByRef customInfoMap As CMap) As Boolean
    getCustomInfoMapByNeName = False
    
    If baseNeCustomInfoMap_.hasKey(neName) Then
        Set customInfoMap = baseNeCustomInfoMap_.GetAt(neName)
        If Not customInfoMap.isEmpty Then
            getCustomInfoMapByNeName = True
        End If
    End If
End Function

Private Sub initNeBoardStyleMap()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(getResByKey("BaseTransPort"))

    Dim rowNumber As Long
    Dim neName As String
    Dim boardstyleColName As String
    Dim boardstyleColLeter As String
    Dim boardStyleName As String
    Dim neNameColLeter As String
    Dim neNameColName As String
    Dim neNameStr As String
    

    boardstyleColName = getResByKey("Board Style Name")
    boardstyleColLeter = getColnumNameLeter(ws, boardstyleColName)
    neNameColName = getNeNameColum
    neNameColLeter = getColnumNameLeter(ws, neNameColName)
    
    If boardstyleColLeter = "" Or neNameColLeter = "" Then Exit Sub
    
    '找到基站传输页上functionName所在的列的列号（Letter形式）
    Dim enodebfuntionNameColLeter As String, nodebfunctionNameColLeter As String, gbtsfunctionNameColLeter As String
    Call getFunctionNameColumLetter(enodebfuntionNameColLeter, nodebfunctionNameColLeter, gbtsfunctionNameColLeter)

    
    For rowNumber = 3 To ws.range("A1048576").End(xlUp).row
        neName = ws.range(neNameColLeter & rowNumber).value
        boardStyleName = ws.range(boardstyleColLeter & rowNumber).value
        If neName <> "" And boardStyleName <> "" Then
            Call neBoardStyleMap_.SetAt(neName, boardStyleName)
            Call boardStyleNeMap_.SetAt(boardStyleName, neName)
        End If
        neNameStr = neName
        If enodebfuntionNameColLeter <> "" Then neNameStr = neNameStr + "," + ws.range(enodebfuntionNameColLeter & rowNumber).value
        If nodebfunctionNameColLeter <> "" Then neNameStr = neNameStr + "," + ws.range(nodebfunctionNameColLeter & rowNumber).value
        If gbtsfunctionNameColLeter <> "" Then neNameStr = neNameStr + "," + ws.range(gbtsfunctionNameColLeter & rowNumber).value
        
        Call FunctionNeMap_.SetAt(neName, neNameStr)
        
        If neName <> "" Then Call initBaseNeCustomInfoMap(ws, rowNumber, neName)
    Next rowNumber
End Sub

'找到基站传输页上FunctionName的列位置（Letter形式）
Private Function getFunctionNameColumLetter(ByRef enodebfunctionNameCol As String, ByRef nodebfunctionNameCol As String, ByRef gbtsfunctionNameCol As String)
    
    Dim sheetDef As Worksheet
    Dim index As Long
    Dim defCollection As New Collection
    Dim mappingdefgroupName As String
    Dim mappingdefcolumnName As String
    Dim mocName As String
    Dim attributeName As String
    Dim enodebfunctionNameColName As String, nodebfunctionNameColName As String, gbtsfunctionNameColName As String

    Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For index = 2 To sheetDef.range("a1048576").End(xlUp).row
        mappingdefgroupName = sheetDef.Cells(index, 2)
        mappingdefcolumnName = sheetDef.Cells(index, 3)
        mocName = sheetDef.Cells(index, 4)
        attributeName = sheetDef.Cells(index, 5)
        
        If mocName = "eNodeBFunction" And attributeName = "eNodeBFunctionName" Then
            enodebfunctionNameColName = mappingdefcolumnName
        End If
        If mocName = "NODEBFUNCTION" And attributeName = "NODEBFUNCTIONNAME" Then
            nodebfunctionNameColName = mappingdefcolumnName
        End If
        If mocName = "GBTSFUNCTION" And attributeName = "GBTSFUNCTIONNAME" Then
            gbtsfunctionNameColName = mappingdefcolumnName
        End If
    Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(getResByKey("BaseTransPort"))
    
    enodebfunctionNameCol = findColLetterByRowAndValue(ws, 2, enodebfunctionNameColName)
    nodebfunctionNameCol = findColLetterByRowAndValue(ws, 2, nodebfunctionNameColName)
    gbtsfunctionNameCol = findColLetterByRowAndValue(ws, 2, gbtsfunctionNameColName)

End Function

'缓存每个NE对应的场景定制信息
Private Sub initBaseNeCustomInfoMap(ByRef ws As Worksheet, ByRef rowNumber As Long, ByRef neName As String)
    Dim customKeyName As Variant
    Dim customColLetter As String
    Dim customColValue As String
    Dim customNameValue As CMap
    Set customNameValue = New CMap

    For Each customKeyName In baseCustomNameLetterMap_.KeyCollection
        customColLetter = baseCustomNameLetterMap_.GetAt(customKeyName)
        customColValue = Trim(ws.range(customColLetter & rowNumber).value)
        If customColValue <> "" Then Call customNameValue.SetAt(CStr(customKeyName), customColValue)
    Next
    
    Call baseNeCustomInfoMap_.SetAt(neName, customNameValue)
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

Private Function getNeNameColum() As String
    Dim sheetDef As Worksheet
    Dim index As Long
    Dim defCollection As New Collection
    Dim mappingdefgroupName As String
    Dim mappingdefcolumnName As String
    Dim mocName As String
    Dim attributeName As String
    Dim neType As String
    
    getNeNameColum = ""
    Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For index = 2 To sheetDef.range("a1048576").End(xlUp).row
        mappingdefgroupName = sheetDef.Cells(index, 2)
        mappingdefcolumnName = sheetDef.Cells(index, 3)
        mocName = sheetDef.Cells(index, 4)
        attributeName = sheetDef.Cells(index, 5)
        neType = sheetDef.Cells(index, 12)
        
        If mocName = "NE" And attributeName = "NENAME" And neType = "Base Station" Then
            getNeNameColum = mappingdefcolumnName
            Exit For
        End If
    Next
End Function

Private Sub initBaseCustomInfoMap()
    Dim sheetDef As Worksheet
    Dim index As Long
    Dim defCollection As New Collection
    Dim mappingdefgroupName As String, mappingdefcolumnName As String, mappingdefsheetName As String
    Dim mocName As String, attributeName As String
    Dim columnLetter As String
    Dim baseStaionWs As Worksheet
    Dim baseCustomKeyName As String '标识基站传输页上定制字段的主键，格式：GroupName,ColumnName
    
    Set baseStaionWs = ThisWorkbook.Worksheets(getResByKey("BaseTransPort"))
    
    Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
    For index = 2 To sheetDef.range("a1048576").End(xlUp).row
        mappingdefsheetName = sheetDef.Cells(index, 1)
        mappingdefgroupName = sheetDef.Cells(index, 2)
        mappingdefcolumnName = sheetDef.Cells(index, 3)
        mocName = sheetDef.Cells(index, 4)
        attributeName = sheetDef.Cells(index, 5)
        
        If mappingdefsheetName <> baseStaionWs.name Or mappingdefgroupName = "" Or mappingdefcolumnName = "" _
         Or InStr(mocName, "Customization") = 0 Then GoTo NextLoop
        
        baseCustomKeyName = mappingdefgroupName & customDelimiter_ & mappingdefcolumnName
        
        If Not baseCustomNameLetterMap_.hasKey(baseCustomKeyName) Then
            columnLetter = findColLetterByGrpAndColName(baseStaionWs, mappingdefgroupName, mappingdefcolumnName, 2)
            If columnLetter <> "" Then Call baseCustomNameLetterMap_.SetAt(baseCustomKeyName, columnLetter)
        End If
NextLoop:
    Next
End Sub

Private Function getBaseCustomGroupLetter(ByRef customColName As String) As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(getResByKey("BaseTransPort"))
    Dim columnNumber As Long
    Dim columnName As String
    
    getBaseCustomGroupLetter = ""
    
    For columnNumber = 1 To ws.range("XFD" & 2).End(xlToLeft).column
        columnName = ws.Cells(2, columnNumber).value
        If customColName = columnName Then
            getBaseCustomGroupLetter = getColStr(columnNumber)
            Exit Function
        End If
    Next columnNumber
End Function

Private Sub getFunctionNameColum(ByRef enodebfunctionNameCol As String, ByRef nodebfunctionNameCol As String, ByRef gbtsfunctionNameCol As String)
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
        If mocName = "eNodeBFunction" And attributeName = "eNodeBFunctionName" Then
            enodebfunctionNameCol = mappingdefcolumnName
        End If
        If mocName = "NODEBFUNCTION" And attributeName = "NODEBFUNCTIONNAME" Then
            nodebfunctionNameCol = mappingdefcolumnName
        End If
        If mocName = "GBTSFUNCTION" And attributeName = "GBTSFUNCTIONNAME" Then
            gbtsfunctionNameCol = mappingdefcolumnName
        End If
    Next
End Sub

Private Sub Class_Initialize()
    Set neBoardStyleMap_ = New CMap
    Set boardStyleNeMap_ = New CMap
    Set FunctionNeMap_ = New CMap
    Set baseCustomNameLetterMap_ = New CMap
    Set baseNeCustomInfoMap_ = New CMapValueObject
End Sub



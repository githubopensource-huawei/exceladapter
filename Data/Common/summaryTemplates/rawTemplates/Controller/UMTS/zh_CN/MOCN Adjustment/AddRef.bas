Attribute VB_Name = "AddRef"
Option Explicit
Public hisHyperlinkMap  As Collection
Private controlDefineMap As Collection
Public mappingDefineMap As Collection


Private Sub setRefValidation(ByVal contedType As String, ByVal boundValue As String, sheet As Worksheet, ByVal RowNum As Long, ByVal colNum As Long)
    On Error Resume Next
    Dim inputTitle As String
    inputTitle = getResByKey("Range")
    If contedType = "Enum" Then
        If boundValue <> sheet.Cells(RowNum, colNum).Validation.Formula1 Or sheet.Cells(RowNum, colNum).Validation.inputMessage = "" Then
            With sheet.Cells(RowNum, colNum).Validation
                .Add Type:=xlValidateList, Formula1:=boundValue
                .inputTitle = getResByKey("Range")
                .inputMessage = "[" + boundValue + "]"
                .ShowInput = True
                .ShowError = False
            End With
            sheet.Cells(RowNum, colNum).Validation.Modify Type:=xlValidateList, Formula1:=boundValue
        End If
    End If
End Sub

Sub initControlDefineMap()
        Set controlDefineMap = getAllControlDefines()
End Sub
Sub initmappingDefineMap()
        Set mappingDefineMap = getAllMappingDefs()
End Sub
Function getControlDefine(sheetName As String, groupName As String, columnName As String) As CControlDef
        Dim key As String
        Dim def As CControlDef
        
        key = sheetName + "," + groupName + "," + columnName
        If controlDefineMap Is Nothing Then
            initControlDefineMap
        End If
        
        If Contains(controlDefineMap, key) Then
            Set def = controlDefineMap(key)
        End If
        Set getControlDefine = def
End Function


Function getMappingDefine(sheetName As String, groupName As String, columnName As String) As CMappingDef
        Dim key As String
        Dim def As CMappingDef
        
        key = sheetName + "," + groupName + "," + columnName
        If mappingDefineMap Is Nothing Then
            initmappingDefineMap
        End If
        
        If Contains(mappingDefineMap, key) Then
            Set def = mappingDefineMap(key)
        End If
        Set getMappingDefine = def
End Function

Function getAllMappingDefs() As Collection
        Dim mp As Collection
        Dim mpdef As CMappingDef
        Dim sheetDef As Worksheet
        Dim index As Long
        Dim defCollection As New Collection
        Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
        For index = 2 To sheetDef.range("a1048576").End(xlUp).row
                Set mpdef = New CMappingDef
                mpdef.sheetName = sheetDef.Cells(index, 1)
                mpdef.groupName = sheetDef.Cells(index, 2)
                mpdef.columnName = sheetDef.Cells(index, 3)
                mpdef.mocName = sheetDef.Cells(index, 4)
                mpdef.attributeName = sheetDef.Cells(index, 5)
                mpdef.neType = sheetDef.Cells(index, 12)
                mpdef.neVersion = sheetDef.Cells(index, 13)
                If Not Contains(defCollection, mpdef.getKey) Then
                    defCollection.Add Item:=mpdef, key:=mpdef.getKey
                End If
        Next
        Set getAllMappingDefs = defCollection
End Function

Function getAllControlDefines() As Collection
    On Error Resume Next
    Dim ctlDef As CControlDef
    Dim sheetDef As Worksheet
    Dim index As Long
    Dim defCollection As New Collection
    
    Set sheetDef = ThisWorkbook.Worksheets("CONTROL DEF")
    
    For index = 2 To sheetDef.range("a1048576").End(xlUp).row
            Set ctlDef = New CControlDef
            ctlDef.mocName = sheetDef.Cells(index, 1).value
            ctlDef.attributeName = sheetDef.Cells(index, 2).value
            ctlDef.dataType = sheetDef.Cells(index, 3).value
            ctlDef.bound = sheetDef.Cells(index, 4).value
            ctlDef.lstValue = sheetDef.Cells(index, 5).value
            ctlDef.controlInfo = sheetDef.Cells(index, 6).value
            ctlDef.sheetName = sheetDef.Cells(index, 7).value
            ctlDef.groupName = sheetDef.Cells(index, 8).value
            ctlDef.columnName = sheetDef.Cells(index, 9).value
            ctlDef.neType = sheetDef.Cells(index, 10).value
            If Not Contains(defCollection, ctlDef.getKey) Then
                defCollection.Add Item:=ctlDef, key:=ctlDef.getKey
            End If
    Next
    Set getAllControlDefines = defCollection
End Function









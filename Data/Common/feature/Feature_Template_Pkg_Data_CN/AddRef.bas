Attribute VB_Name = "AddRef"
Option Explicit
Private controlDefineMap As Collection
Private mappingDefineMap As Collection

Private Function Contains(coll As Collection, key As String) As Boolean
        On Error GoTo NotFound
        Call coll(key)
        Contains = True
        Exit Function
NotFound:
        Contains = False
End Function

Sub initmappingDefineMap()
        Set mappingDefineMap = getAllMappingDefs()
End Sub

Function getAllMappingDefs() As Collection
        Dim mp As Collection
        Dim mpdef As CMappingDef
        Dim sheetDef As Worksheet
        Dim index As Long
        Dim defCollection As New Collection
        Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
        For index = 2 To sheetDef.Range("a65536").End(xlUp).row
                Set mpdef = New CMappingDef
                mpdef.sheetName = sheetDef.Cells(index, 1)
                mpdef.GroupName = sheetDef.Cells(index, 2)
                mpdef.columnName = sheetDef.Cells(index, 3)
                mpdef.mocName = sheetDef.Cells(index, 4)
                mpdef.attributeName = sheetDef.Cells(index, 5)
                If sheetDef.Cells(index, 7) = "1" Or LCase(sheetDef.Cells(index, 7)) = "true" Then
                    mpdef.isKey = True
                Else
                    mpdef.isKey = False
                End If
                mpdef.neType = sheetDef.Cells(index, 12)
                mpdef.neVersion = sheetDef.Cells(index, 13)
                If Not Contains(defCollection, mpdef.getKey) Then
                    defCollection.Add Item:=mpdef, key:=mpdef.getKey
                End If
        Next
        Set getAllMappingDefs = defCollection
End Function

Function getMappingDefine(mocName As String, attrName As String) As CMappingDef
        Dim key As String
        Dim def As CMappingDef
        key = mocName + "," + attrName
        If mappingDefineMap Is Nothing Then
            initmappingDefineMap
        End If
        
        If Contains(mappingDefineMap, key) Then
            Set def = mappingDefineMap(key)
        End If
        Set getMappingDefine = def
End Function

Sub initControlDefineMap()
        Set controlDefineMap = getAllControlDefines()
End Sub

Function getControlDefine(sheetName As String, GroupName As String, columnName As String) As CControlDef
        Dim key As String
        Dim def As CControlDef
        
        key = sheetName + "," + GroupName + "," + columnName
        If controlDefineMap Is Nothing Then
            initControlDefineMap
        End If
        
        If Contains(controlDefineMap, key) Then
            Set def = controlDefineMap(key)
        End If
        Set getControlDefine = def
End Function

Function getAllControlDefines() As Collection
    On Error Resume Next
    Dim ctlDef As CControlDef
    Dim sheetDef As Worksheet
    Dim index As Integer
    Dim defCollection As New Collection
    Set sheetDef = ThisWorkbook.Worksheets(ControllSheetName)
    For index = HiddenSheetValidRowBegin To sheetDef.Range("a65536").End(xlUp).row
            Set ctlDef = New CControlDef
            ctlDef.mocName = sheetDef.Cells(index, 1).value
            ctlDef.attributeName = sheetDef.Cells(index, 2).value
            ctlDef.dataType = sheetDef.Cells(index, 3).value
            ctlDef.bound = sheetDef.Cells(index, 4).value
            ctlDef.lstValue = sheetDef.Cells(index, 5).value
            ctlDef.controlInfo = sheetDef.Cells(index, 6).value
            ctlDef.sheetName = sheetDef.Cells(index, 7).value
            ctlDef.GroupName = sheetDef.Cells(index, 8).value
            ctlDef.columnName = sheetDef.Cells(index, 9).value
            ctlDef.neType = sheetDef.Cells(index, 10).value
            If Not Contains(defCollection, ctlDef.getKey) Then
                defCollection.Add Item:=ctlDef, key:=ctlDef.getKey
            End If
    Next
    Set getAllControlDefines = defCollection
End Function




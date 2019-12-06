Attribute VB_Name = "AddRef"
Option Explicit
Public hisHyperlinkMap  As Collection
Private controlDefineMap As Collection
Public mappingDefineMap As Collection


Sub initControlDefineMap()
        Set controlDefineMap = getAllControlDefines()
End Sub

'Sub initmappingDefineMap()
'        Set mappingDefineMap = getAllMappingDefs()
'End Sub

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


'Function getMappingDefine(sheetName As String, groupName As String, columnName As String) As CMappingDef
'        Dim key As String
'        Dim def As CMappingDef
'
'        key = sheetName + "," + groupName + "," + columnName
'        If mappingDefineMap Is Nothing Then
'            initmappingDefineMap
'        End If
'
'        If Contains(mappingDefineMap, key) Then
'            Set def = mappingDefineMap(key)
'        End If
'        Set getMappingDefine = def
'End Function

'Function getAllMappingDefs() As Collection
'        Dim mp As Collection
'        Dim mpdef As CMappingDef
'        Dim sheetDef As Worksheet
'        Dim index As Long
'        Dim defCollection As New Collection
'        Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
'        For index = 2 To sheetDef.Range("a65536").End(xlUp).row
'                Set mpdef = New CMappingDef
'                mpdef.sheetName = sheetDef.Cells(index, 1)
'                mpdef.groupName = sheetDef.Cells(index, 2)
'                mpdef.columnName = sheetDef.Cells(index, 3)
'                mpdef.mocName = sheetDef.Cells(index, 4)
'                mpdef.attributeName = sheetDef.Cells(index, 5)
'                mpdef.neType = sheetDef.Cells(index, 12)
'                mpdef.neVersion = sheetDef.Cells(index, 13)
'                If Not Contains(defCollection, mpdef.getKey) Then
'                    defCollection.Add Item:=mpdef, key:=mpdef.getKey
'                End If
'        Next
'        Set getAllMappingDefs = defCollection
'End Function

Function getAllControlDefines() As Collection
    On Error Resume Next
    Dim ctlDef As CControlDef
    Dim ctrlDefSht As Worksheet
    Dim index As Long
    Dim defCollection As New Collection
    
    Set ctrlDefSht = ThisWorkbook.Worksheets("CONTROL DEF")

    Dim conInfor As String
    With ctrlDefSht
        For index = 2 To .range("a65536").End(xlUp).row
            Set ctlDef = New CControlDef
            Dim ctrlInfoItemsArray As Variant
            ctrlInfoItemsArray = .range("A" & index & ":J" & index).value
        
            ctlDef.mocName = Trim(CStr(ctrlInfoItemsArray(1, 1)))
            ctlDef.attributeName = Trim(CStr(ctrlInfoItemsArray(1, 2)))
            ctlDef.dataType = Trim(CStr(ctrlInfoItemsArray(1, 3)))
            ctlDef.bound = Trim(CStr(ctrlInfoItemsArray(1, 4)))
            ctlDef.lstValue = Trim(CStr(ctrlInfoItemsArray(1, 5)))
            conInfor = Trim(CStr(ctrlInfoItemsArray(1, 6)))
            If isControlInfoRef(conInfor) Then conInfor = getRealControlInfo(conInfor)
            ctlDef.controlInfo = conInfor
            ctlDef.sheetName = Trim(CStr(ctrlInfoItemsArray(1, 7)))
            ctlDef.groupName = Trim(CStr(ctrlInfoItemsArray(1, 8)))
            ctlDef.columnName = Trim(CStr(ctrlInfoItemsArray(1, 9)))
            ctlDef.neType = Trim(CStr(ctrlInfoItemsArray(1, 10)))
            
            If Not Contains(defCollection, ctlDef.getKey) Then
                defCollection.Add Item:=ctlDef, key:=ctlDef.getKey
            End If
        Next
    End With
    Set getAllControlDefines = defCollection
End Function











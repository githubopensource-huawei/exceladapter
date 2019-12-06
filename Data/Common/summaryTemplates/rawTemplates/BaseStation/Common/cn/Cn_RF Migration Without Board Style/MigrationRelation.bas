Attribute VB_Name = "MigrationRelation"
Public Const MigrationBarNameAddSourceNe As String = "MigrationAddSourceNe"

Public Sub AddDelMigrationSourceNe()
    Dim neType As String
    neType = getNeType()
    If neType = "MRAT" Then
        ConfigMRATSrcNEForm.Show
    Else
        ConfigULSrcNEForm.Show
    End If
    
End Sub


Public Function isMigrationRelationSheet(ByRef ws As Worksheet) As Boolean
    Dim sheetName As String, migrationSheetName As String
    isMigrationRelationSheet = False

    If ws.name = getResByKey("RF Migration NEs Relationship") Then
        isMigrationRelationSheet = True
    End If
End Function

Public Sub MigrationSelectionChange(ByRef ws As Worksheet, ByRef target As Range)
    On Error GoTo ErrorHandler
    Dim groupName As String, columnName As String
    Dim rowNumber As Long, columnNumber As Long
    
    rowNumber = target.row
    columnNumber = target.column
    
    If target.Rows.count <> 1 Or target.Columns.count <> 1 Then Exit Sub
    If isMigrationRelationSheet(ws) = False Then Exit Sub
    If target.row < 4 Then Exit Sub
    
    Dim migrationNemap As CMigrationNeMap
    Set migrationNemap = New CMigrationNeMap

    Dim srcNeMap As CMap
    Set srcNeMap = migrationNemap.RelationNeColumnNameMap(rowNumber)
    
    Dim referencedString As String
    groupName = ws.Cells(2, columnNumber).value
    columnName = ws.Cells(3, columnNumber).value
    If InStr(columnName, getResByKey("RADIO")) = 0 Then Exit Sub
    
    If columnName = getResByKey("Global Radio GBTS reference") And srcNeMap.hasKey("BTS") Then
        referencedString = srcNeMap.GetAt("BTS")
    End If
    
    If columnName = getResByKey("Global Radio NodeB reference") And srcNeMap.hasKey("NodeB") Then
        referencedString = srcNeMap.GetAt("NodeB")
    End If
    
    If columnName = getResByKey("Global Radio eNodeB reference") And srcNeMap.hasKey("eNodeB") Then
        referencedString = srcNeMap.GetAt("eNodeB")
    End If
    
    If target.Borders.LineStyle <> xlNone Then
        Call setBoardStyleListBoxRangeValidation(ws.name, groupName, columnName, referencedString, ws, target)
    Else
        target.Validation.Delete
    End If
    
ErrorHandler:
End Sub

Public Function isSrcNameCol(neType As String, cellValue As String) As Boolean
    Dim num As Long
    Dim tempValue As String
'    Dim tempCnValue As String
    isSrcNameCol = False
    
    
    
    For num = 1 To 50
'        tempEnValue = neType + CStr(num) + " NE Name"
'        tempCnValue = neType + CStr(num) + " 网元名称"
        tempValue = neType + CStr(num) + " " + getResByKey("NE_NAME")
        If tempValue = cellValue Then
            isSrcNameCol = True
            Exit Function
        End If
    Next
End Function

Public Function getTargetNeClounmNum(ByRef ws As Worksheet) As Long
    Dim m_colNum As Long
    Dim tempFunctionStr As String
'    Dim tempEnFunctionStr As String
    
'    tempCnFunctionStr = "目标网元"
'    tempEnFunctionStr = "Target NE"
    tempFunctionStr = getResByKey("TARGET_NE_NAME")
    
    getTargetNeClounmNum = -1
    
    For m_colNum = 1 To ws.Range("XFD2").End(xlToLeft).column
        If ws.Cells(2, m_colNum).value = tempFunctionStr Then
            getTargetNeClounmNum = m_colNum
            Exit Function
        End If
    Next
End Function
'Public Function isCnEnv(ByRef ws As Worksheet) As Boolean
'    Dim m_colNum As Long
'    Dim tempCnFunctionStr As String
'    Dim tempEnFunctionStr As String
'
'    tempCnFunctionStr = "目标网元"
'    tempEnFunctionStr = "Target NE"
'
'    isCnEnv = False
'
'    For m_colNum = 1 To ws.range("XFD2").End(xlToLeft).column
'        If ws.Cells(2, m_colNum).value = tempCnFunctionStr Then
'            isCnEnv = True
'            Exit Function
'        End If
'    Next
'End Function

'Creat Tool Bar
Public Sub initAddSourceNeToolBar(ByRef ws As Worksheet)
    Call deleteSourceNeToolBar
    
    If isMigrationRelationSheet(ws) Then
        Call insertSourceNeToolBar
    End If

End Sub
Public Sub deleteSourceNeToolBar()
    If containsAToolBar(MigrationBarNameAddSourceNe) Then
        Application.CommandBars(MigrationBarNameAddSourceNe).Delete
    End If

End Sub
Private Sub insertSourceNeToolBar()
    Dim AddSourceNeBar As CommandBar
    Dim AddSourceNeButton As CommandBarButton

    Set AddSourceNeBar = Application.CommandBars.Add(MigrationBarNameAddSourceNe, msoBarTop)
    With AddSourceNeBar
        .Protection = msoBarNoResize
        .Visible = True
        Set AddSourceNeButton = .Controls.Add(Type:=msoControlButton)
        With AddSourceNeButton
            .Style = msoButtonIconAndCaption
            .Caption = getResByKey("Configure Source NE Columns")
            .TooltipText = getResByKey("Configure Source NE Columns")
            .OnAction = "AddDelMigrationSourceNe"
            .FaceId = 186
            .Enabled = True
        End With
    End With

End Sub

Public Sub getSrcColNum(ByRef srcNeType As String, ByRef srcColNum As Long)
    If isExistRadioGroup("BTS", srcColNum) Then
        srcNeType = "GBTS"
    ElseIf isExistRadioGroup("NodeB", srcColNum) Then
        srcNeType = "NodeB"
    ElseIf isExistRadioGroup("eNodeB", srcColNum) Then
        srcNeType = "eNodeB"
    End If
End Sub
Public Function isExistGroup(ByRef neType As String) As Boolean
    Dim m_colNum As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    Dim tempFunctionStr As String
'    Dim tempEnFunctionStr As String
    Dim tempNeType As String
    Dim getTargetNeCol As Long
    tempNeType = neType
    If neType = "BTS" Then tempNeType = "GBTS"
    
    getTargetNeCol = getTargetNeClounmNum(ws)
    
    tempFunctionStr = tempNeType + "/ " + tempNeType + " " + getResByKey("FUNCTION")
'    tempEnFunctionStr = tempNeType + "/ " + tempNeType + " Function"
    
    isExistGroup = False
    
    For m_colNum = 1 To getTargetNeCol
        If ws.Cells(2, m_colNum).value = tempFunctionStr Then
            isExistGroup = True
            Exit Function
        End If
    Next
End Function
Public Function isExistRadioGroup(ByRef neType As String, ByRef srcColNum As Long) As Boolean
    Dim m_colNum As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    Dim tempCnFunctionStr As String
    Dim tempEnFunctionStr As String
    Dim tempNeType As String
    tempNeType = neType
    If neType = "BTS" Then tempNeType = "GBTS"
    
    tempEnFunctionStr = "Global Radio " + tempNeType + " " + getResByKey("RADIO_REFERENCE")
    tempCnFunctionStr = tempNeType + getResByKey("RADIO_REFERENCE")
    
    isExistRadioGroup = False
    
    For m_colNum = 1 To ws.Range("XFD2").End(xlToLeft).column
        If ws.Cells(3, m_colNum).value = tempCnFunctionStr Or ws.Cells(3, m_colNum).value = tempEnFunctionStr Then
            isExistRadioGroup = True
            srcColNum = m_colNum
            Exit Function
        End If
    Next
End Function

Public Sub setFunctionNameBoxLst(ByVal sh As Object, ByVal target As Range)
    Dim sheet As New Worksheet
    Dim cellRange As Range
    Dim tempcellRange As Range
    Dim groupName As String
    Dim columnName As String
    Dim rowNum As Long
    Dim colNum As Long
    Dim m_colNum As Long
    Dim stationNeType As String
    Set sheet = sh
    stationNeType = getNeType()
    If target.count Mod 256 = 0 Then Exit Sub

    
    Dim migrationDataManager As CMigrationDataManager
    Dim cellNeNameColumMap As CMap
    Dim baseStationNeMap As CMapValueObject
    Set migrationDataManager = New CMigrationDataManager
    Call migrationDataManager.init
    Set cellNeNameColumMap = migrationDataManager.cellNeNameColumMap
    Set baseStationNeMap = migrationDataManager.baseStationNeMap
    
    Dim groupColnameStr As String
    Dim neType As String
    Dim srcneName As String
    Dim neName As String
    
    Dim boxstr As String
    
    For Each cellRange In target
        If cellRange.Interior.colorIndex = 34 Or cellRange.Interior.colorIndex = 40 Or cellRange.Borders.LineStyle = xlLineStyleNone Then
            Exit Sub
        End If
        
        Call getGroupAndColumnName(sheet, cellRange, groupName, columnName)
        rowNum = cellRange.row
        colNum = cellRange.column
        groupColnameStr = groupName + "," + columnName
        
        neType = getFunctionNameColumNeType(cellNeNameColumMap, groupColnameStr)
        If (groupName = "GTRX" Or groupName = getResByKey("GTRX_ZH")) And (columnName = "*BTS Name" Or columnName = getResByKey("COLNAME_BTSNAME")) Then
            neType = "BTS"
        End If
        
        If neType = "" And isNeNameColumn(sheet, groupName, columnName) Then neType = "Node"
        If neType = "" Then GoTo NextLoop
        
        '默认源站名称在目标站名称的左边一个单元格
        If isSrcNeColumn(sheet, rowNum, colNum - 1) Then
            srcneName = sheet.Cells(rowNum, colNum - 1).value '需要转为物理网元
            
            boxstr = getBoxStrWhenHasSrcNeName(migrationDataManager, srcneName, stationNeType, neType)
            Call setFunctionNameListBoxRange(sheet, groupName, columnName, cellRange, boxstr)
            
        Else '没有源基站列时
            neName = sheet.Cells(rowNum, colNum).value
            
            '如果基站名称为空，没有ListBox；如果基站名称不为空，ListBox显示当前值，及其对应的目标站和源站的集合
            If neName <> "" Then
                boxstr = getBoxStrWhenOnlyDstNeName(migrationDataManager, neName, stationNeType, neType)
                Call setFunctionNameListBoxRange(sheet, groupName, columnName, cellRange, boxstr)
            Else
                target.Validation.Delete
            End If
        End If
NextLoop:
    Next
End Sub

Private Function getBoxStrWhenHasSrcNeName(ByRef migrationDataManager As CMigrationDataManager, ByRef srcneName As String, _
    ByRef stationNeType As String, ByRef neType As String) As String
    Dim pythsrcNeName As String
    Dim boxstr As String
    
    pythsrcNeName = srcneName
    If stationNeType = "MRAT" And neType <> "Node" Then pythsrcNeName = getpythNeNamebyFunctionName(srcneName, neType, migrationDataManager.baseStationNeMap)

    boxstr = getTargetNeBoxStr(migrationDataManager, neType, pythsrcNeName, stationNeType)
    If boxstr <> "" Then
        boxstr = boxstr + "," + srcneName
    Else
        boxstr = srcneName
    End If
    
    getBoxStrWhenHasSrcNeName = boxstr
End Function

Private Function getBoxStrWhenOnlyDstNeName(ByRef migrationDataManager As CMigrationDataManager, ByRef neName As String, _
    ByRef stationNeType As String, ByRef neType As String) As String
    Dim pythNeName As String
    Dim boxstr As String
    
    pythNeName = neName
    If stationNeType = "MRAT" And neType <> "Node" Then pythNeName = getpythNeNamebyFunctionName(neName, neType, migrationDataManager.baseStationNeMap)
                
    '先假设此为目标站，取其源站信息
    boxstr = getSourceNeBoxStr(migrationDataManager, neType, pythNeName, stationNeType)
    
    '如果源站信息为空,假设此为源站，取其目标站信息
    If boxstr = "" Then
        boxstr = getTargetNeBoxStr(migrationDataManager, neType, pythNeName, stationNeType)
    End If
    
    If boxstr <> "" Then
        boxstr = boxstr + "," + neName
    Else
        boxstr = neName
    End If

    getBoxStrWhenOnlyDstNeName = boxstr
End Function

'判断是否是源基站名称列
Private Function isSrcNeColumn(ByRef curSheet As Worksheet, rowNum As Long, colNum As Long) As Boolean
    isSrcNeColumn = False
    
    '默认源站名称在目标站名称的左边一个单元格
    If colNum < 1 Then
        Exit Function
    End If
    
    '取列名称
    Dim groupName As String
    Dim columnName As String
    Set cellRange = curSheet.Cells(rowNum, colNum)
    Call getGroupAndColumnName(curSheet, cellRange, groupName, columnName)
    
    '根据列名称找到属性名称，判断属性名称是否为SOURCENENAME 或 SOURCEBTSNAME 或 SOURCENODEBNAME
    Dim mapdef As CMappingDef
    Set mapdef = getMappingDefine(curSheet.name, groupName, columnName)
    If mapdef Is Nothing Then Exit Function
    If mapdef.neType = "" Then Exit Function
    
    If mapdef.attributeName = "SOURCENENAME" Or mapdef.attributeName = "SOURCEBTSNAME" Or mapdef.attributeName = "SOURCENODEBNAME" Then
        isSrcNeColumn = True
    End If
End Function

'判断是否目标基站名称列
Private Function isNeNameColumn(ByRef curSheet As Worksheet, groupName As String, columnName As String) As Boolean
    isNeNameColumn = False
    
    Dim mapdef As CMappingDef
    Set mapdef = getMappingDefine(curSheet.name, groupName, columnName)
    If mapdef Is Nothing Then Exit Function
    If mapdef.neType = "" Then Exit Function
        
    If mapdef.attributeName = "NENAME" Or mapdef.attributeName = "GBTSFUNCTIONNAME" _
        Or mapdef.attributeName = "NODEBFUNCTIONNAME" Or mapdef.attributeName = "eNodeBFunctionName" Then
        isNeNameColumn = True
    End If
    
End Function

Private Function setFunctionNameListBoxRange(ByRef sheet As Worksheet, ByRef groupName As String, ByRef columnName As String, _
    ByRef target As Range, ByRef referencedString As String)
    
    If referencedString <> "" Then
        Call setBoardStyleListBoxRangeValidation(sheet.name, groupName, columnName, referencedString, sheet, target)
    Else
        target.Validation.Delete
    End If
            
End Function


Private Function getpythNeNamebyFunctionName(ByRef functionName As String, ByRef neType As String, ByRef baseStationNeMap As CMapValueObject) As String
    Dim onerowRec As CMap
    Dim targetValue As Variant
    Dim tempfunctionValue As String
    getpythNeNamebyFunctionName = ""
    For Each targetValue In baseStationNeMap.KeyCollection
        Set onerowRec = baseStationNeMap.GetAt(targetValue)
        If onerowRec.hasKey(neType) Then tempfunctionValue = onerowRec.GetAt(neType)
        If functionName = tempfunctionValue Then
            getpythNeNamebyFunctionName = targetValue
            Exit Function
        End If
    Next
End Function


Private Function getTargetNeBoxStr(ByRef migrationDataManager As CMigrationDataManager, ByRef neType As String, ByRef srcneName As String, ByRef stationNeType As String) As String
    Dim targetSourceNeMap As CMapValueObject
    Dim baseStationNeMap As CMapValueObject
    Dim onerowNeMap As CMap
    Dim srcNeNameStr As String
    Dim tempNeType As String
    tempNeType = neType
    If stationNeType <> "MRAT" Then tempNeType = "BASESTATION"
    
    getTargetNeBoxStr = ""

    Set targetSourceNeMap = migrationDataManager.targetSourceNeMap
    Set baseStationNeMap = migrationDataManager.baseStationNeMap
    
    Dim targetNeName As Variant
    For Each targetNeName In targetSourceNeMap.KeyCollection
        Set onerowNeMap = targetSourceNeMap.GetAt(targetNeName)
        If neType = "Node" Then
            For Each neTypeIter In onerowNeMap.KeyCollection
                If isExitSrcNe(srcneName, onerowNeMap.GetAt(neTypeIter)) Then getTargetNeBoxStr = targetNeName
            Next
        End If
        
        If onerowNeMap.hasKey(neType) Then
            srcNeNameStr = onerowNeMap.GetAt(neType)
            If isExitSrcNe(srcneName, srcNeNameStr) Then getTargetNeBoxStr = getTargetFunctionNeName(baseStationNeMap, CStr(targetNeName), tempNeType) '单站需要处理这个neType
        End If
    Next
    
End Function

Private Function getSourceNeBoxStr(ByRef migrationDataManager As CMigrationDataManager, ByRef neType As String, ByRef dstPhyNeName As String, ByRef stationNeType As String) As String
    Dim targetSourceNeMap As CMapValueObject
    Dim baseStationNeMap As CMapValueObject
    Dim onerowNeMap As CMap
    Dim phySrcNeNameStr As String
    Dim phySrcNeNameArray() As String
    
    Dim tempNeType As String
    tempNeType = neType
    If stationNeType <> "MRAT" Then tempNeType = "BASESTATION"
    
    getSourceNeBoxStr = ""

    Set targetSourceNeMap = migrationDataManager.targetSourceNeMap
    Set baseStationNeMap = migrationDataManager.baseStationNeMap
    
    If targetSourceNeMap.hasKey(dstPhyNeName) Then
        Set onerowNeMap = targetSourceNeMap.GetAt(dstPhyNeName)
        If onerowNeMap.hasKey(neType) Then
            phySrcNeNameStr = onerowNeMap.GetAt(neType)  '得到的是Node名称，需要再转换成RAT名称
        End If
    End If
    
    If phySrcNeNameStr = "" Then
        Exit Function
    End If
        
    phySrcNeNameArray = Split(phySrcNeNameStr, ",")
    
    Dim delimiter As String
    delimiter = ""
    
    Dim index As Long
    For index = LBound(phySrcNeNameArray) To UBound(phySrcNeNameArray)
        getSourceNeBoxStr = getSourceNeBoxStr + delimiter + getTargetFunctionNeName(baseStationNeMap, CStr(phySrcNeNameArray(index)), tempNeType)
        delimiter = ","
    Next
    
End Function

Private Function isExitSrcNe(ByRef srcneName As String, ByRef srcNeNameStr As String) As Boolean
    isExitSrcNe = False
    If srcNeNameStr = "" Then Exit Function
    
    Dim srcNeNameArry() As String
    srcNeNameArry = Split(srcNeNameStr, ",")
    
    Dim index As Long
    For index = LBound(srcNeNameArry) To UBound(srcNeNameArry)
        If srcNeNameArry(index) = srcneName Then
            isExitSrcNe = True
            Exit Function
        End If
    Next
End Function

Private Function getTargetFunctionNeName(ByRef baseStationNeMap As CMapValueObject, ByRef targetNeName As String, ByRef neType As String) As String
    Dim onerowRec As CMap
    getTargetFunctionNeName = ""
    If baseStationNeMap.hasKey(targetNeName) Then Set onerowRec = baseStationNeMap.GetAt(targetNeName)
    If onerowRec.hasKey(neType) Then getTargetFunctionNeName = onerowRec.GetAt(neType)
End Function

Private Function getFunctionNameColumNeType(ByRef cellNeNameColumMap As CMap, ByRef groupColnameStr As String) As String
    Dim keyValueStr As Variant
    Dim valueStr As String
    getFunctionNameColumNeType = ""
    For Each keyValueStr In cellNeNameColumMap.KeyCollection
        valueStr = cellNeNameColumMap.GetAt(keyValueStr)
        If valueStr = groupColnameStr Then
            getFunctionNameColumNeType = keyValueStr
            Exit Function
        End If
    Next
End Function

Public Sub refreshMigrationSourceData(ByRef ws As Worksheet)
    Dim sheetName As String
    Dim groupName As String
    Dim keyValue As Variant
    Dim keyValueStr As String
    Dim sheetgroupNameArry() As String
    Dim sourceRowNumber As Long
    Dim maxColnum As Long
    
    If sourceBoardStyleDataMap Is Nothing Then Exit Sub
    
    If allBoardStyleData Is Nothing Then Call initAllBoardStyleDataPublic
    Call allBoardStyleData.initBoardStyleDataMap
    
    For Each keyValue In sourceBoardStyleDataMap.KeyCollection
        keyValueStr = keyValue
        maxColnum = sourceBoardStyleDataMap.GetAt(keyValueStr)
        sheetgroupNameArry = Split(keyValueStr, ",")
        sheetName = sheetgroupNameArry(0)
        groupName = sheetgroupNameArry(1)
        sourceRowNumber = sheetgroupNameArry(2)
        
        If ws.name = sheetName Then Call setMigrationRecbackColor(ws, sourceRowNumber, maxColnum)
    Next
    
    sourceBoardStyleDataMap.Clean
End Sub



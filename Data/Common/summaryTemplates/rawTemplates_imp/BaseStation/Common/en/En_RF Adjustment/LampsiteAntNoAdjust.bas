Attribute VB_Name = "LampsiteAntNoAdjust"
Option Explicit

Private Const ANT_ADJUST_BAR_NAME = "AntAdjustBar"
Private SEC_GRP_SHEET_NAME As String
Private Const FINISH_BAR_NAME = "AdjustFinishBar"
Private Const CANCEL_BAR_NAME = "AdjustCancelBar"
'用于记录是否在调整天线端口的标识
Public InAdjustLampsiteAntnPort As Boolean
'1表示为UMTS小区，2表示为LTE小区
Public ANT_CELL_TYPE As Long
'小区扇区对应关系集合，声明为public是为了用于时间响应使用
Private cellId_secEqmGrpIdMap As Collection

Private ROW_COUNT As Long


Private Cell_Row_Map As CMap
Private Cell_CellType_Map As CMap

Private SITE_NAME As String
Private WRITESUCCESS As Boolean
Private Cell_TrxListInfo As CMap
Private valueMap As Collection
Private Cell_Index As Long
Private cellValueStr As String
Private Sector_Index As Long
Private Board_Index As Long
Private boardValueStr As String
Private Ante_Index As Long
Private Model_Index As Long
Private antnValueStr As String
Private modelValueStr As String
Private BaseEqm_Index As Long
Private SectorEqmGrpId_Index As Long
Private sectoreqmValueStr As String
Private VMprfAutoAfg_Index As Long

Private errCollect As Collection

Private dupCollect As Collection

Private BTSName_Index As Long
Private selectBtsNameCol As Collection '用户选择基站名称列表
Private btsNameRowCountMap As CMap 'key:BTSNAME#CELLNAME,value:RowCount，基站名称和临时页签行数的映射
Private tempShtStartRow As Long '临时页签boardNo起始行记录
Private btsNameBrdNoMap As CMap 'key:BTSNAME,value:boardNo.，基站名称和单板列表的映射


Private Sub adjustAntPortMain()
    Dim activeShtName As String
    activeShtName = ThisWorkbook.ActiveSheet.name
    SEC_GRP_SHEET_NAME = activeShtName
    '小区类型判断，按照当前工作表中页签判断，
    ANT_CELL_TYPE = getCellShtType(SEC_GRP_SHEET_NAME)
    
    If ANT_CELL_TYPE = -1 Then Exit Sub
    '调用生成窗体
    LampsiteMuliBtsFilterForm.Show
End Sub

Public Function isSectorEqmGroupSht(sheetName As String) As Boolean
    If sheetName = "EUSECTOREQMGROUP" Or sheetName = getResByKey("EUSECTOREQMGROUP") _
    Or sheetName = "ULOCELLSECEQMGRP" Or sheetName = getResByKey("ULOCELLSECEQMGRP") Then
        isSectorEqmGroupSht = True
        Exit Function
    End If
    isSectorEqmGroupSht = False
End Function

Public Sub createAntAdjustBar()
    Dim baseStationChooseBar As CommandBar
    Dim baseStationStyle As CommandBarButton
    
    Call deleteAntAdjustBar
    Set baseStationChooseBar = Application.CommandBars.Add(ANT_ADJUST_BAR_NAME, msoBarTop)
    With baseStationChooseBar
        .Protection = msoBarNoResize
        .Visible = True
        Set baseStationStyle = .Controls.Add(Type:=msoControlButton)
        With baseStationStyle
            .Style = msoButtonIconAndCaption
            .Caption = getResByKey("AdjustSecGrpAntPort")
            .TooltipText = getResByKey("AdjustSecGrpAntPort")
            .OnAction = "adjustAntPortMain"
            .FaceId = 186
            .Enabled = True
        End With
      End With
      
End Sub

Public Sub deleteAntAdjustBar()
    If existToolBar(ANT_ADJUST_BAR_NAME) Then
        Application.CommandBars(ANT_ADJUST_BAR_NAME).Delete
    End If
End Sub

Public Sub deleteAntTempBar()
    If existToolBar(FINISH_BAR_NAME) Then
        Application.CommandBars(FINISH_BAR_NAME).Delete
    End If
    If existToolBar(CANCEL_BAR_NAME) Then
        Application.CommandBars(CANCEL_BAR_NAME).Delete
    End If
    
End Sub

Private Function existToolBar(ByRef barName As String) As Boolean
    On Error GoTo ErrorHandler
    existToolBar = True
    Dim bar As CommandBar
    Set bar = CommandBars(barName)
    Exit Function
ErrorHandler:
    existToolBar = False
End Function

Sub createAntTempBar()
    Dim finishBar As CommandBar
    Dim finishStyle As CommandBarButton
    Dim cancelStyle As CommandBarButton
    If ThisWorkbook.ActiveSheet.name <> getResByKey("Temp_Adjust_Sheet") Then
        Exit Sub
    End If
    
    Set finishBar = Application.CommandBars.Add(FINISH_BAR_NAME, msoBarTop)
    With finishBar
        .Protection = msoBarNoResize
        .Visible = True
        Set finishStyle = .Controls.Add(Type:=msoControlButton)
        With finishStyle
            .Style = msoButtonIconAndCaption
            .Caption = getResByKey("Finish")
            .TooltipText = getResByKey("Finish")
            .OnAction = "writeBackDataToSecGrp"
            .FaceId = 186
            .Enabled = True
        End With
        Set cancelStyle = .Controls.Add(Type:=msoControlButton)
        With cancelStyle
            .Style = msoButtonIconAndCaption
            .Caption = getResByKey("Cancel")
            .TooltipText = getResByKey("Cancel")
            .OnAction = "deleteSecGrpTempSheet"
            .FaceId = 186
            .Enabled = True
        End With
      End With
      
End Sub

'删除临时页签
Private Sub deleteSecGrpTempSheet()
    Dim tmpsheet As Worksheet
    Dim cellsheet As Worksheet
    InAdjustLampsiteAntnPort = False
    If SEC_GRP_SHEET_NAME <> "" Then
        Set cellsheet = ThisWorkbook.Worksheets(SEC_GRP_SHEET_NAME)
        cellsheet.Activate
    End If
    Set tmpsheet = ThisWorkbook.Worksheets(getResByKey("Temp_Adjust_Sheet"))
    Call changeAlerts(False)
    tmpsheet.Delete
    Call changeAlerts(True)
End Sub

Private Sub changeAlerts(ByRef flag As Boolean)
    Application.EnableEvents = flag
    Application.DisplayAlerts = flag
    Application.ScreenUpdating = flag
End Sub

'根据当前工作薄中页签，判断是LTE小区还是UMTS小区
Public Function getCellShtType(activeShtName As String) As Long
    If activeShtName = "EUSECTOREQMGROUP" Or activeShtName = getResByKey("EUSECTOREQMGROUP") Or activeShtName = "EUPRBSECTOREQMGROUP" Or activeShtName = getResByKey("EUPRBSECTOREQMGROUP") Then
        getCellShtType = 2
    ElseIf activeShtName = "ULOCELLSECEQMGRP" Or activeShtName = getResByKey("ULOCELLSECEQMGRP") Then
        getCellShtType = 1
    Else
        getCellShtType = -1
    End If
End Function

Public Function getBaseStationMocNameAndAttrNameForLampsite(ByRef mocName As String, ByRef attrName As String)
    If getNeType() = "MRAT" Then
        If ANT_CELL_TYPE = 1 Then
            attrName = "NODEBFUNCTIONNAME"
            mocName = "NODEBFUNCTION"
        ElseIf ANT_CELL_TYPE = 2 Then
            attrName = "eNodeBFunctionName"
            mocName = "eNodeBFunction"
        End If
    Else
        attrName = "NENAME"
        mocName = "NE"
    End If
End Function

'生成临时页签，用于用户调整数据
Public Sub insertColAndWriteData(selectedMocCol As Collection, CellSheetName As String)
    On Error GoTo ErrorHandler
    Dim temBtsName As Variant
    Set valueMap = New Collection
    Set btsNameBrdNoMap = New CMap
    Set selectBtsNameCol = selectedMocCol
    
    'SITE_NAME = siteName
    WRITESUCCESS = False
    
    Call judgeGNormalCell
    
    Dim chkPassed As Boolean
    chkPassed = True

    ROW_COUNT = calculateRow(chkPassed)
    
    If Not chkPassed Then
        Exit Sub
    End If
    
    ThisWorkbook.Worksheets.Add after:=ThisWorkbook.ActiveSheet
    ThisWorkbook.ActiveSheet.name = getResByKey("Temp_Adjust_Sheet")
    
    Call createAntTempBar
    Call initMenuStatus(ThisWorkbook.ActiveSheet)
    InAdjustLampsiteAntnPort = True
    '初始化Map用于存放小区ID和扇区设备组ID对应关系
    'Dim cellId_secEqmGrpIdMap As Collection
    Call getCellIdAndSecEqmGrpIdMap(cellId_secEqmGrpIdMap)
    
    Call insertBtsNameColumn
    
    '插入小区ID列
    If Not insertCellIdColumn() Then
        Exit Sub
    End If
    '插入扇区设备组标识
    If Not insertSectorEqmGrpIdColumn() Then
        Exit Sub
    End If
    '插入扇区分裂ID
    If Not insertSectorIdColumn() Then
        Exit Sub
    End If
    
    tempShtStartRow = 2
    For Each temBtsName In selectBtsNameCol
        '插入单板编号，找任意基站只插入列头，成功则退出循环
        If insertRxuBoardColumn(CStr(temBtsName)) = True Then
            Exit For
        End If
    Next
    
    '插入天线端口列
    Call insertAntenneColumn
    '插入天线端口模式列
    Call insertSecGrpAnteModelColumn
    
    'If ANT_CELL_TYPE = 2 Then
        '插入基带设备编号列
        'If Not insertBaseEqmColumn() Then
        '    Exit Sub
        'End If
        '插入扇区设备组标识
'        If Not insertSectorEqmGrpIdColumn() Then
'            Exit Sub
'        End If
        'If Not insertCellBeamModeColumn() Then
        'End If
    'End If
    
    '插入子卡免配置列
    Call insertVMprfAutoAfgColumn
'    If CELL_TYPE = 1 Then
'        If Not insertSectoreqmColumn() Then
'        End If
'    End If
    '向临时页签写入数据
    Call writeDataToTemSheet(cellId_secEqmGrpIdMap)
    Call AdjustTemSheetStyle
    WRITESUCCESS = True
    Exit Sub
ErrorHandler:
    WRITESUCCESS = False
End Sub

Private Sub getCellIdAndSecEqmGrpIdMap(cellId_secEqmGrpIdMap As Collection)
    Set cellId_secEqmGrpIdMap = New Collection
    Dim cellsheet As Worksheet
    Dim constCellTempCol As Long
    Dim constSecEqmGrpIdcolIndex As Long
    Dim tempCellId As String
    Dim tempSecGrpId As String
    Dim mocName As String
    Dim attrName As String
    Dim temBtsName As Variant
    
    Call getCellMocNameAndAttrName(mocName, attrName)
    If SEC_GRP_SHEET_NAME = "EUPRBSECTOREQMGROUP" Or SEC_GRP_SHEET_NAME = getResByKey("EUPRBSECTOREQMGROUP") Then
        mocName = "EuPrbSectorEqmGroup"
        attrName = "LocalCellId"
    End If
    constCellTempCol = getColNum(SEC_GRP_SHEET_NAME, 2, attrName, mocName)
    
    Call getSecGrpSectoreqmGrpIdMocNameAndAttrName(mocName, attrName)
    If SEC_GRP_SHEET_NAME = "EUPRBSECTOREQMGROUP" Or SEC_GRP_SHEET_NAME = getResByKey("EUPRBSECTOREQMGROUP") Then
            mocName = "EuPrbSectorEqmGroup"
        End If
    constSecEqmGrpIdcolIndex = getColNum(SEC_GRP_SHEET_NAME, 2, attrName, mocName)
    
    Set cellsheet = ThisWorkbook.Worksheets(SEC_GRP_SHEET_NAME)
    
    Dim index As Long
    Dim temKeyVal As String
    For Each temBtsName In selectBtsNameCol
        For index = 2 To cellsheet.range("a1048576").End(xlUp).row
            If cellsheet.Cells(index, 1).value = temBtsName And cellsheet.Cells(index, 2).value <> "RMV" Then
                tempCellId = cellsheet.Cells(index, constCellTempCol).value
                tempSecGrpId = cellsheet.Cells(index, constSecEqmGrpIdcolIndex).value
                
                temKeyVal = temBtsName + "*" + tempCellId
                If Contains(cellId_secEqmGrpIdMap, temKeyVal) Then
                    Dim secEqmGrpIdColl As Collection
                    Set secEqmGrpIdColl = cellId_secEqmGrpIdMap(temKeyVal)
                    
                    secEqmGrpIdColl.Add Item:=tempSecGrpId
                Else
                    Dim temSecEqmGrpIdColl As Collection
                    Set temSecEqmGrpIdColl = New Collection
                    
                    temSecEqmGrpIdColl.Add Item:=tempSecGrpId
                    
                    cellId_secEqmGrpIdMap.Add Item:=temSecEqmGrpIdColl, key:=temKeyVal
                End If
                
            End If
        Next
    Next
End Sub
    

Private Sub judgeGNormalCell()
    On Error GoTo ErrorHandler
    Dim mocName As String
    Dim attrName As String
    Dim cellTypeIndex As Long
    Dim cellTypeVal As String
    
    Dim rowNum As Long
    Dim maxRow As Long
    Dim cellsheet As Worksheet
    Dim cellId As String
    Dim temBtsName As Variant
    Dim mapKey As String
    
    Set Cell_Row_Map = New CMap
    Set Cell_CellType_Map = New CMap
    Set cellsheet = ThisWorkbook.Worksheets(SEC_GRP_SHEET_NAME)
    Call getCellMocNameAndAttrName(mocName, attrName)
    If SEC_GRP_SHEET_NAME = "EUPRBSECTOREQMGROUP" Or SEC_GRP_SHEET_NAME = getResByKey("EUPRBSECTOREQMGROUP") Then
        mocName = "EuPrbSectorEqmGroup"
        attrName = "LocalCellId"
    End If
    cellTypeIndex = getColNum(SEC_GRP_SHEET_NAME, 2, attrName, mocName)
    
    For Each temBtsName In selectBtsNameCol
        maxRow = cellsheet.range("a1048576").End(xlUp).row
        For rowNum = 3 To maxRow
            If cellsheet.Cells(rowNum, 1).value = CStr(temBtsName) Then
                cellId = cellsheet.Cells(rowNum, cellTypeIndex).value
                mapKey = CStr(temBtsName) + "*" + cellId
                If Cell_Row_Map.hasKey(mapKey) = False Then
                    Call Cell_Row_Map.SetAt(mapKey, rowNum)
                End If
                
                If Cell_CellType_Map.hasKey(mapKey) = False Then
                    Call Cell_CellType_Map.SetAt(mapKey, ANT_CELL_TYPE)
                End If
            End If
        Next
    Next
ErrorHandler:
End Sub

Private Function calculateRow(ByRef chkPassed As Boolean) As Long
    Dim mocName As String
    Dim attrName As String
    Dim columnName As String
    Dim rowCount As Long
    Dim columnIndex As Long
    Dim sectorIndex As Long
    Dim baseEqmIndex As Long
    Dim sectorEqmIndex As Long
    Dim sectorEqmGrpIdIndex As Long
    Dim cellBeamModeIndex As Long
    Dim vmprfAutoCfgIndex As Long
    Dim secGrpSheet As Worksheet
    Dim index As Long
    Dim antenneIndex As Long
    Dim sectorArray As Validation
    
    Set Cell_TrxListInfo = New CMap
    Set secGrpSheet = ThisWorkbook.Worksheets(SEC_GRP_SHEET_NAME)
    Set btsNameRowCountMap = New CMap
    
    '基站名称默认在第一列
    Dim btsNameColIndex As Long
    Call getBaseStationMocNameAndAttrNameForLampsite(mocName, attrName)
    btsNameColIndex = getColNum(SEC_GRP_SHEET_NAME, 2, attrName, mocName)
    
    '找小区ID所在的列
    Dim constCellTempCol As Long
    Call getCellMocNameAndAttrName(mocName, attrName)
    If SEC_GRP_SHEET_NAME = "EUPRBSECTOREQMGROUP" Or SEC_GRP_SHEET_NAME = getResByKey("EUPRBSECTOREQMGROUP") Then
        mocName = "EuPrbSectorEqmGroup"
        attrName = "LocalCellId"
    End If
    constCellTempCol = getColNum(SEC_GRP_SHEET_NAME, 2, attrName, mocName)
    
    '获取天线端口所在列
    Call getAntenneMocNameAndAttr(mocName, attrName)
    If SEC_GRP_SHEET_NAME = "EUPRBSECTOREQMGROUP" Or SEC_GRP_SHEET_NAME = getResByKey("EUPRBSECTOREQMGROUP") Then
        mocName = "EuPrbSectorEqmGroup"
    End If
    'columnName = findColumnFromRelationDef(SEC_GRP_SHEET_NAME, mocName, attrName)
    'columnIndex = findColumnByName(secGrpSheet, columnName, 2)
    columnIndex = getColNum(SEC_GRP_SHEET_NAME, 2, attrName, mocName)
    '获取扇区所在列
    Call getSectorMocNameAndAttr(mocName, attrName)
    If SEC_GRP_SHEET_NAME = "EUPRBSECTOREQMGROUP" Or SEC_GRP_SHEET_NAME = getResByKey("EUPRBSECTOREQMGROUP") Then
        mocName = "EuPrbSectorEqmGroup"
    End If
    'columnName = findColumnFromRelationDef(SEC_GRP_SHEET_NAME, mocName, attrName)
    'sectorIndex = findColumnByName(secGrpSheet, columnName, 2)
    sectorIndex = getColNum(SEC_GRP_SHEET_NAME, 2, attrName, mocName)
    
    '获取子卡免配置列
    Call getVMprfAutoAfgMocNameAndAttrName(mocName, attrName)
    If SEC_GRP_SHEET_NAME = "EUPRBSECTOREQMGROUP" Or SEC_GRP_SHEET_NAME = getResByKey("EUPRBSECTOREQMGROUP") Then
        mocName = "EuPrbSectorEqmGroup"
    End If
    vmprfAutoCfgIndex = getColNum(SEC_GRP_SHEET_NAME, 2, attrName, mocName)
    
    'If ANT_CELL_TYPE = 2 Then
        'Call getBaseEqmMocNameAndAttrName(mocName, attrName)
        'columnName = findColumnFromRelationDef(SEC_GRP_SHEET_NAME, mocName, attrName)
        'baseEqmIndex = findColumnByName(secGrpSheet, columnName, 2)
        'baseEqmIndex = getColNum(SEC_GRP_SHEET_NAME, 2, attrName, mocName)
        
        Call getSecGrpSectoreqmGrpIdMocNameAndAttrName(mocName, attrName)
        If SEC_GRP_SHEET_NAME = "EUPRBSECTOREQMGROUP" Or SEC_GRP_SHEET_NAME = getResByKey("EUPRBSECTOREQMGROUP") Then
            mocName = "EuPrbSectorEqmGroup"
        End If
        'columnName = findColumnFromRelationDef(SEC_GRP_SHEET_NAME, mocName, attrName)
        sectorEqmGrpIdIndex = getColNum(SEC_GRP_SHEET_NAME, 2, attrName, mocName)
        
'        Call getCellbeamModeMocNameAndAttrName(mocName, attrName)
'        columnName = findColumnFromRelationDef(SEC_GRP_SHEET_NAME, mocName, attrName)
'        cellBeamModeIndex = findColumnByName(secGrpSheet, columnName, 2)
    'End If
    
'    If ANT_CELL_TYPE = 1 Then
'        Call getSectoreqmMocNameAndAttrName(mocName, attrName)
'        columnName = findColumnFromRelationDef(SEC_GRP_SHEET_NAME, mocName, attrName)
'        sectorEqmIndex = findColumnByName(secGrpSheet, columnName, 2)
'    End If
    
    rowCount = 0
    If columnIndex <= 0 Then
        calculateRow = rowCount
        Exit Function
    End If
    Dim antennes As CLampsiteAntennes
    Dim temBtsName As String
  
    For index = 3 To secGrpSheet.range("a1048576").End(xlUp).row
        temBtsName = secGrpSheet.Cells(index, btsNameColIndex).value
        If existInCollection(temBtsName, selectBtsNameCol) And secGrpSheet.Cells(index, constCellTempCol).value <> "" And secGrpSheet.Cells(index, 2).value <> "RMV" Then
            Set antennes = New CLampsiteAntennes
            Dim tem_key As String
            
            antennes.btsName = secGrpSheet.Cells(index, btsNameColIndex).value
            antennes.cellId = secGrpSheet.Cells(index, constCellTempCol).value
            
            'key采用BTSNAME#CELLNAME的形式
            Dim tempBtsName As String
            Dim tempCellId As String
            tempBtsName = secGrpSheet.Cells(index, btsNameColIndex).value
            tempCellId = secGrpSheet.Cells(index, constCellTempCol).value
            tem_key = tempBtsName + "*" + tempCellId
            
            antennes.sectorEqmGrpIds = secGrpSheet.Cells(index, sectorEqmGrpIdIndex).value
            antennes.sectorIds = secGrpSheet.Cells(index, sectorIndex).value
            antennes.antennes = secGrpSheet.Cells(index, columnIndex).value
            antennes.vmprfAutoCfg = secGrpSheet.Cells(index, vmprfAutoCfgIndex).value
            
            antennes.ranges = secGrpSheet.Cells(index, columnIndex).address(False, False)
            antennes.row = index
            
            '获取每个“基站#小区”对应的行数，用于确定单板编号下拉列表范围
            Dim antennesCol As Collection
            Dim sflag As Boolean
            Dim rowVal As Long
            
           Set antennesCol = antennes.getAntenneCollection(sflag, rowVal)
            
            '当sflag为false证明获取antennesCol失败，直接删除临时页签，退出
            If sflag = False Then
                Call deleteSecGrpTempSheet
                ThisWorkbook.Worksheets(SEC_GRP_SHEET_NAME).rows(rowVal).Select
                Exit Function
            End If
            
            If Not btsNameRowCountMap.hasKey(CStr(tem_key)) Then
                Call btsNameRowCountMap.SetAt(CStr(tem_key), antennesCol.count)
                rowCount = rowCount + antennesCol.count
            End If
            
            '维护数据集合，collection(基站名称#小区ID，collection(扇区设备组ID， collection(天线端口)))
            If Contains(valueMap, tem_key) Then
                Dim sectorEqmGrpMap As Collection
                Set sectorEqmGrpMap = valueMap(tem_key)
                If Contains(sectorEqmGrpMap, secGrpSheet.Cells(index, sectorEqmGrpIdIndex).value) Then
                    sectorEqmGrpMap.Add Item:=antennes
                Else
                    Dim antennesMap As Collection
                    Set antennesMap = New Collection
                    antennesMap.Add Item:=antennes
                    
                    sectorEqmGrpMap.Add Item:=antennesMap, key:=secGrpSheet.Cells(index, sectorEqmGrpIdIndex).value
                End If
            Else
                Dim antennesNewMap As Collection
                Set antennesNewMap = New Collection
                antennesNewMap.Add Item:=antennes
                
                Dim sectorEqmGrpNewMap As Collection
                Set sectorEqmGrpNewMap = New Collection
                sectorEqmGrpNewMap.Add Item:=antennesNewMap, key:=secGrpSheet.Cells(index, sectorEqmGrpIdIndex).value
                
                valueMap.Add Item:=sectorEqmGrpNewMap, key:=tem_key
                
            End If
        End If
    Next
    
    calculateRow = rowCount
End Function


'Private Function getBaseStationMocNameAndAttrName(ByRef mocName As String, ByRef attrName As String)
'    If getNeType() = "MRAT" Then
'        If CELL_TYPE = 0 Or CELL_TYPE = 4 Then
'            attrName = "GBTSFUNCTIONNAME"
'            mocName = "GBTSFUNCTION"
'        ElseIf CELL_TYPE = 1 Then
'            attrName = "NODEBFUNCTIONNAME"
'            mocName = "NODEBFUNCTION"
'        ElseIf CELL_TYPE = 2 Then
'            attrName = "eNodeBFunctionName"
'            mocName = "eNodeBFunction"
'        End If
'    Else
'        attrName = "NENAME"
'        mocName = "NE"
'    End If
'End Function

Private Function getCellMocNameAndAttrName(ByRef mocName As String, ByRef attrName As String)
    If ANT_CELL_TYPE = 1 Then
        attrName = "ULOCELLID"
        mocName = "ULOCELLSECEQMGRP"
    ElseIf ANT_CELL_TYPE = 2 Then
        attrName = "LocalCellId"
        mocName = "EuSectorEqmGroup"
    End If
End Function

Private Sub getSectorMocNameAndAttr(ByRef mocName As String, ByRef attrName As String)
    If ANT_CELL_TYPE = 1 Then
        mocName = "ULOCELLSECEQMGRP"
        attrName = "VSECTORID"
    ElseIf ANT_CELL_TYPE = 2 Then
        mocName = "EuSectorEqmGroup"
        attrName = "VSECTORID"
    End If
End Sub

Private Sub getAntenneMocNameAndAttr(ByRef mocName As String, ByRef attrName As String)
    If ANT_CELL_TYPE = 1 Then
        mocName = "ULOCELLSECEQMGRP"
        attrName = "VRXUANTNO"
    ElseIf ANT_CELL_TYPE = 2 Then
        mocName = "EuSectorEqmGroup"
        attrName = "VRXUANTNO"
    End If
End Sub

Public Function getSecGrpSectoreqmGrpIdMocNameAndAttrName(ByRef mocName As String, ByRef attrName As String)
    If ANT_CELL_TYPE = 2 Then
        attrName = "SectorEqmGroupId"
        mocName = "EuSectorEqmGroup"
    ElseIf ANT_CELL_TYPE = 1 Then
        attrName = "SECTOREQMGRPID"
        mocName = "ULOCELLSECEQMGRP"
    End If
End Function

'Private Function getBaseEqmMocNameAndAttrName(ByRef mocName As String, ByRef attrName As String)
'    If ANT_CELL_TYPE = 2 Then
'        attrName = "BaseBandEqmId"
'        mocName = "EuSectorEqmGroup"
'    End If
'End Function

Private Function getVMprfAutoAfgMocNameAndAttrName(ByRef mocName As String, ByRef attrName As String)
    If ANT_CELL_TYPE = 2 Then
        attrName = "VMPRFAUTOCFG"
        mocName = "EuSectorEqmGroup"
    ElseIf ANT_CELL_TYPE = 1 Then
        attrName = "VMPRFAUTOCFG"
        mocName = "ULOCELLSECEQMGRP"
    End If
End Function

Private Function insertCellIdColumn() As Boolean
        Dim myAttrName As String
        Dim myCellMocName As String
        Dim constCellTempCol As Long
        Dim mocName As String
        Dim attrName As String
        
        Call getCellMocNameAndAttrName(myCellMocName, myAttrName)
        If SEC_GRP_SHEET_NAME = "EUPRBSECTOREQMGROUP" Or SEC_GRP_SHEET_NAME = getResByKey("EUPRBSECTOREQMGROUP") Then
            myCellMocName = "EuPrbSectorEqmGroup"
            myAttrName = "LocalCellId"
        End If
        constCellTempCol = getColNum(SEC_GRP_SHEET_NAME, 2, myAttrName, myCellMocName)
        
        Dim cellsheet As Worksheet
        insertCellIdColumn = True
        Set cellsheet = ThisWorkbook.Worksheets(SEC_GRP_SHEET_NAME)
        Cell_Index = 2
        cellsheet.Cells(2, constCellTempCol).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, Cell_Index)
        Dim cellsStr As String
        cellsStr = ""
        
        Dim cellIdMap As CMap
        Set cellIdMap = New CMap
        
        Dim btsNameColIndex As Long
        Call getBaseStationMocNameAndAttrNameForLampsite(mocName, attrName)
        btsNameColIndex = getColNum(SEC_GRP_SHEET_NAME, 2, attrName, mocName)

        Dim index As Long
        Dim temCellIdStr As String
        Dim temBtsName As Variant
        Dim temKeyVal As String
        
        For Each temBtsName In selectBtsNameCol
            For index = 3 To cellsheet.range("a1048576").End(xlUp).row
                If temBtsName = cellsheet.Cells(index, btsNameColIndex).value And cellsheet.Cells(index, 2).value <> "RMV" Then
                    temCellIdStr = cellsheet.Cells(index, constCellTempCol).value
                    temKeyVal = temBtsName + "*" + temCellIdStr
                    
                    If cellIdMap.hasKey(temKeyVal) = False Then
                        Call cellIdMap.SetAt(temKeyVal, temCellIdStr)
                    End If
    
                End If
            Next
        Next

        Dim sVal As Variant
        For Each sVal In cellIdMap.ValueCollection
            If cellsStr <> "" Then
                cellsStr = cellsStr + "," + CStr(sVal)
            Else
                cellsStr = CStr(sVal)
            End If
        Next
        
'        Dim cellRang As range
'        Set cellRang = ThisWorkbook.ActiveSheet.range("A2:A" + CStr(2 + ROW_COUNT + 3))
        If cellsStr <> "" Then
                cellValueStr = cellsStr
'                With cellRang.Validation
'                   .delete
'                   .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=cellsStr
'                End With
'                ThisWorkbook.ActiveSheet.Columns(Cell_Index).EntireColumn.AutoFit
        End If
        insertCellIdColumn = True
End Function

Private Function insertSectorIdColumn() As Boolean
        Dim mocName As String
        Dim attrName As String
        Dim sectorColumnName As String
        Dim columnIndex As Long
        Dim cellsheet As Worksheet
        
        Set cellsheet = ThisWorkbook.Worksheets(SEC_GRP_SHEET_NAME)
        Call getSectorMocNameAndAttr(mocName, attrName)
        If SEC_GRP_SHEET_NAME = "EUPRBSECTOREQMGROUP" Or SEC_GRP_SHEET_NAME = getResByKey("EUPRBSECTOREQMGROUP") Then
            mocName = "EuPrbSectorEqmGroup"
        End If
        'sectorColumnName = findColumnFromRelationDef(SEC_GRP_SHEET_NAME, mocName, attrName)
        columnIndex = getColNum(SEC_GRP_SHEET_NAME, 2, attrName, mocName)
        If columnIndex < 0 Then
            insertSectorIdColumn = False
            Exit Function
        End If

        Sector_Index = 4

        cellsheet.Cells(2, columnIndex).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, Sector_Index)
         Call refreshCommentText(ThisWorkbook.ActiveSheet.Cells(1, Sector_Index), getResByKey("SECTOR_SECTORID"))
        ThisWorkbook.ActiveSheet.Columns(Sector_Index).EntireColumn.AutoFit
        insertSectorIdColumn = True
End Function

'插入单板编号列
Private Function insertRxuBoardColumn(ByVal temBtsName As String) As Boolean
    insertRxuBoardColumn = False
    Dim brdStyleSheetName As String
    Dim grpCollection As Collection
    Dim brdStr As String
    Dim brdGrp
    Dim startRow As Long
    Dim endRow As Long
    Dim index As Long
    Dim btsIndex As Long
    Dim charStr As String
    Dim mainSheetName As String
    Dim mainSheet As Worksheet
    Dim boardNoIndex As Long
    btsIndex = -1
    
    brdStyleSheetName = findBoardStyleSheetByBtsName(temBtsName)
    If brdStyleSheetName = "" Then
        Call MsgBox(getResByKey("NoBoradStyle"), vbInformation, getResByKey("Warning"))
        Call deleteSecGrpTempSheet
        Exit Function
    End If
    
    brdStr = ""
    Set grpCollection = findBrdGroups
    Dim boardStyleSheet As Worksheet
    
    Set boardStyleSheet = ThisWorkbook.Worksheets(brdStyleSheetName)
    
    Board_Index = 5
    charStr = "E"
    
    For Each brdGrp In grpCollection
        Call getGroupStartAndEndRowByGroupName(boardStyleSheet, CStr(brdGrp), startRow, endRow)
        
        If startRow <> -1 Then
            boardNoIndex = getboradNoColumNumber(boardStyleSheet, startRow + 1, CStr(brdGrp))
            boardStyleSheet.Cells(startRow + 1, boardNoIndex).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, Board_Index)
            Call refreshCommentText(ThisWorkbook.ActiveSheet.Cells(1, Board_Index), getResByKey("RxuBoard"))
            insertRxuBoardColumn = True
            Exit Function
        End If
    Next
'            For index = startRow + 2 To endRow
'                If isCustomMatchRow(temBtsName, boardStylesheet, CStr(brdGrp), index) Then
'                    If brdStr = "" Then
'                        brdStr = boardStylesheet.Cells(index, boardNoIndex).value
'                    Else
'                        brdStr = brdStr + "," + boardStylesheet.Cells(index, boardNoIndex).value
'                    End If
'                End If
'            Next
'NextLoop:
'        Next brdGrp

        
'        Dim tem_name As Variant
'        Dim rowCount As Long
'        Dim temArray() As String
'        rowCount = 0
'        For Each tem_name In btsNameRowCountMap.KeyCollection
'            temArray = Split(tem_name, "#")
'            If temArray(0) = temBtsName Then
'                rowCount = rowCount + btsNameRowCountMap.GetAt(tem_name)
'            End If
'        Next
'
'        Dim cellRang As range
'        If rowCount <> 0 Then
'            Set cellRang = ThisWorkbook.ActiveSheet.range(charStr + CStr(tempShtStartRow) + ":" + charStr + CStr(rowCount + tempShtStartRow - 1))
'            If brdStr <> "" Then
'                boardValueStr = brdStr
'                If Not btsNameBrdNoMap.haskey(temBtsName) Then
'                    Call btsNameBrdNoMap.SetAt(temBtsName, boardValueStr)
'                End If
'                With cellRang.Validation
'                   .Delete
'                   .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=brdStr
'                End With
'                ThisWorkbook.ActiveSheet.Columns(Board_Index).EntireColumn.AutoFit
'            End If
'            tempShtStartRow = rowCount + tempShtStartRow
'        End If
End Function

Private Function findBoardStyleSheetByBtsName(ByRef btsName As String) As String
    Dim groupName As String
    Dim columnName As String
    Dim mainSheet As Worksheet
    Dim mainSheetName As String
    mainSheetName = GetMainSheetName
    Set mainSheet = ThisWorkbook.Worksheets(mainSheetName)
    
    Dim siteIndex As Long
    Dim brdStyleIndex As Long
    'site index
    groupName = ThisWorkbook.Worksheets(SEC_GRP_SHEET_NAME).Cells(1, 1).value
    columnName = ThisWorkbook.Worksheets(SEC_GRP_SHEET_NAME).Cells(2, 1).value
    siteIndex = findColNumByGrpNameAndColName(mainSheet, groupName, columnName)
    'brd index
    Call findBrdStyleGrpNameAndColName(mainSheetName, groupName, columnName)
    brdStyleIndex = findColNumByGrpNameAndColName(mainSheet, groupName, columnName)
    
    Dim row As Long
    For row = 2 To mainSheet.range("a1048576").End(xlUp).row
         If mainSheet.Cells(row, siteIndex).value = btsName Then
                findBoardStyleSheetByBtsName = mainSheet.Cells(row, brdStyleIndex).value
                Exit Function
         End If
    Next
    findBoardStyleSheetByBtsName = ""
End Function

Private Function findBoardStyleSheet(btsIndex As Long) As String
    Dim groupName As String
    Dim columnName As String
    Dim mainSheet As Worksheet
    Dim mainSheetName As String
    mainSheetName = GetMainSheetName
    Set mainSheet = ThisWorkbook.Worksheets(mainSheetName)
    
    Dim siteIndex As Long
    Dim brdStyleIndex As Long
    'site index
    groupName = ThisWorkbook.Worksheets(SEC_GRP_SHEET_NAME).Cells(1, 1).value
    columnName = ThisWorkbook.Worksheets(SEC_GRP_SHEET_NAME).Cells(2, 1).value
    siteIndex = findColNumByGrpNameAndColName(mainSheet, groupName, columnName)
    'brd index
    Call findBrdStyleGrpNameAndColName(mainSheetName, groupName, columnName)
    brdStyleIndex = findColNumByGrpNameAndColName(mainSheet, groupName, columnName)
    
    Dim row As Long
    For row = 2 To mainSheet.range("a1048576").End(xlUp).row
         If mainSheet.Cells(row, siteIndex).value = SITE_NAME Then
                findBoardStyleSheet = mainSheet.Cells(row, brdStyleIndex).value
                btsIndex = row
                Exit Function
         End If
    Next
    findBoardStyleSheet = ""
End Function

Private Function findColNumByGrpNameAndColName(sh As Worksheet, groupName As String, columnName As String)
    Dim m_colNum As Long
    For m_colNum = 1 To sh.range("XFD2").End(xlToLeft).column
        If get_GroupName(sh.name, m_colNum) = groupName Then
            If GetDesStr(columnName) = GetDesStr(sh.Cells(2, m_colNum).value) Then
                findColNumByGrpNameAndColName = m_colNum
                Exit For
            End If
        End If
    Next
End Function

Private Sub findBrdStyleGrpNameAndColName(sheetName As String, groupName As String, columanName As String)
    Dim relationSheet As Worksheet
    Dim rowCount As Long
    Dim index As Long
    
    Set relationSheet = ThisWorkbook.Worksheets("RELATION DEF")
    rowCount = relationSheet.range("a1048576").End(xlUp).row
    For index = 2 To rowCount
        If relationSheet.Cells(index, 1).value = sheetName And LCase(relationSheet.Cells(index, 4).value) = "true" And _
            LCase(relationSheet.Cells(index, 5).value) = "false" Then
                groupName = relationSheet.Cells(index, 2).value
                columanName = relationSheet.Cells(index, 3).value
                Exit Sub
            Exit Sub
        End If
    Next
End Sub

Private Function findBrdGroups() As Collection
    Dim relationSheet As Worksheet
    Dim rowCount As Long
    Dim index As Long
    Dim grpCollection As New Collection
    Set relationSheet = ThisWorkbook.Worksheets("RELATION DEF")
    rowCount = relationSheet.range("a1048576").End(xlUp).row
    For index = 2 To rowCount
        If UCase(relationSheet.Cells(index, 6).value) = "RRU" Or UCase(relationSheet.Cells(index, 6).value) = "RFU" Or _
            UCase(relationSheet.Cells(index, 6).value) = "AARU" Then
                If Not Contains(grpCollection, UCase(relationSheet.Cells(index, 6).value)) Then
                    grpCollection.Add Item:=relationSheet.Cells(index, 2).value, key:=UCase(relationSheet.Cells(index, 6).value)
                End If
        End If
    Next
    Set findBrdGroups = grpCollection
End Function

'Private Sub getStartAndEndRowNum(brdSheetName As String, ByRef groupName As String, ByRef groupNameStartRowNumber As Long, ByRef groupNameEndRowNumber As Long)
'    Dim sh As Worksheet
'    Set sh = ThisWorkbook.Worksheets(brdSheetName)
'    groupNameStartRowNumber = getGroupNameStartRowNumber(sh, groupName)
'    If groupNameStartRowNumber = -1 Then
'        groupNameEndRowNumber = -1
'    Else
'        groupNameEndRowNumber = groupNameStartRowNumber + getCurrentRegionRowsCount(sh, groupNameStartRowNumber) - 1
'    End If
'End Sub

Private Function getCurrentRegionRowsCount(ByRef ws As Worksheet, ByRef startRowNumber As Long) As Long
    Dim rowNumber As Long
    Dim rowscount As Long
    rowscount = 1
    For rowNumber = startRowNumber + 1 To 2000
        If rowIsBlank(ws, rowNumber) = True Then
            Exit For
        Else
            rowscount = rowscount + 1
        End If
    Next rowNumber
    getCurrentRegionRowsCount = rowscount
End Function

Private Function isCustomRxuBoardNo(ByRef boardStyleSheetName As String, ByRef groupName As String, ByRef rowNumber As Long) As Boolean
    Dim tempboardStyleData As CBoardStyleData
    If tempboardStyleData Is Nothing Then Set tempboardStyleData = New CBoardStyleData
    Call tempboardStyleData.init
    
    isCustomRxuBoardNo = True
    
    Dim baseStationData As CBaseStationData
    Set baseStationData = New CBaseStationData
    Dim boardStyleNeMap As CMap
    Dim neName As String
    Dim baseNeCustomInfoMap As CMapValueObject
    Dim customColValue As CMap
    Dim functionNeMap As CMap
    
    Call baseStationData.init
    Set functionNeMap = baseStationData.functionNeMap
    neName = getNeNamebyFunctionName(functionNeMap)
    Set baseNeCustomInfoMap = baseStationData.baseNeCustomInfoMap
    
    If Not baseNeCustomInfoMap.hasKey(neName) Then Exit Function
    Set customColValue = baseNeCustomInfoMap.GetAt(neName)
    
    Dim keyStr As String
    Dim boardStyleCustomMocInfoMap As CMap
    Dim mocName As String
    Dim customInfo As String
    Dim customInfoArry() As String
    Dim boardStylecustomColletter As String
    Dim boardStylecustomColName As String
    Dim boardStyleSheet As Worksheet
    Dim boardstylecellvalue As String
    
    Set boardStyleSheet = ThisWorkbook.Worksheets(boardStyleSheetName)
    
    mocName = tempboardStyleData.getMocNameByGroupName(groupName)
    Set boardStyleCustomMocInfoMap = tempboardStyleData.getBoardStyleCustomMocInfoMap
    
    If mocName <> "" Then
        keyStr = mocName + "," + groupName
        If boardStyleCustomMocInfoMap.hasKey(keyStr) Then
            customInfo = boardStyleCustomMocInfoMap.GetAt(keyStr)
            If customInfo <> "" Then
                customInfoArry = Split(customInfo, ",")
                boardStylecustomColName = customInfoArry(1)
                boardStylecustomColletter = customInfoArry(2)
                boardstylecellvalue = boardStyleSheet.range(boardStylecustomColletter & rowNumber).value
                isCustomRxuBoardNo = isExistSce(customColValue, boardstylecellvalue, boardStylecustomColName)
            End If
        End If

    End If

End Function

Private Function isExistSce(ByRef customColValue As CMap, ByRef scenValue As String, ByRef boardStylecustomColName As String) As Boolean
    Dim scenValueArry() As String
    Dim baseCustomColValue As String
    Dim index As Long
    
    isExistSce = True
    If Not customColValue.hasKey(boardStylecustomColName) Then Exit Function
    baseCustomColValue = customColValue.GetAt(boardStylecustomColName)
    If scenValue = "" And baseCustomColValue = "" Then Exit Function
    If scenValue = "" And baseCustomColValue <> "" Then
        isExistSce = False
        Exit Function
    End If
    scenValueArry = Split(scenValue, ",")

    For index = LBound(scenValueArry) To UBound(scenValueArry)
     If scenValueArry(index) = baseCustomColValue Then
        isExistSce = True
        Exit Function
     End If
    Next
    isExistSce = False
End Function


'插入天线端口列
Private Sub insertAntenneColumn()
    Dim referencedString As String
    Dim tmpsheet As Worksheet
    Dim rowNum As Long
    Dim mocName As String
    Dim attrName As String
    Dim columnName As String
    Dim antnIndex As Long
    Dim cellsheet As Worksheet
    Dim charStr As String
    
    referencedString = "R0A,R0B,R0C,R0D,R0E,R0F,R0G,R0H,R1A,R1B,R1C,R1D,R2A,R2B,R2C,R2D,R3A,R3B,R3C,R3D,R4A,R4B,R4C,R4D,R5A,R5B,R5C,R5D,R6A,R6B,R6C,R6D,R7A,R7B,R7C,R7D,NULL"
    
    Set cellsheet = ThisWorkbook.Worksheets(SEC_GRP_SHEET_NAME)
    Call getAntenneMocNameAndAttr(mocName, attrName)
    If SEC_GRP_SHEET_NAME = "EUPRBSECTOREQMGROUP" Or SEC_GRP_SHEET_NAME = getResByKey("EUPRBSECTOREQMGROUP") Then
        mocName = "EuPrbSectorEqmGroup"
    End If
    'columnName = findColumnFromRelationDef(SEC_GRP_SHEET_NAME, mocName, attrName)
    antnIndex = getColNum(SEC_GRP_SHEET_NAME, 2, attrName, mocName)
    
'    If ANT_CELL_TYPE = 0 Or ANT_CELL_TYPE = 4 Then
    Ante_Index = 6
    charStr = "F"
'    Else
'        Ante_Index = 4
'        charStr = "D"
'    End If
    
    cellsheet.Cells(2, antnIndex).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, Ante_Index)
    Call refreshCommentText(ThisWorkbook.ActiveSheet.Cells(1, Ante_Index), getResByKey("Antenne"))
    
    Dim antnRang As range
    Set antnRang = ThisWorkbook.ActiveSheet.range(charStr + "2:" + charStr + CStr(2 + ROW_COUNT + 3))
    If referencedString <> "" Then
            antnValueStr = referencedString
            With antnRang.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=referencedString
            End With
            ThisWorkbook.ActiveSheet.Columns(Ante_Index).EntireColumn.AutoFit
    End If
End Sub

'插入天线端口模式列
Private Sub insertSecGrpAnteModelColumn()
         Dim charStr As String
         Dim referencedString As String
         
'        If ANT_CELL_TYPE = 0 Or ANT_CELL_TYPE = 4 Then
        Model_Index = 7
        charStr = "G"
'        Else
'            Model_Index = 5
'            charStr = "E"
'        End If
        referencedString = "RXTX_MODE,RX_MODE,TX_MODE,NULL"
        
        ThisWorkbook.ActiveSheet.Cells(1, Board_Index).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, Model_Index)
        ThisWorkbook.ActiveSheet.Cells(1, Model_Index).value = getResByKey("anteModel")
        Call refreshCommentText(ThisWorkbook.ActiveSheet.Cells(1, Model_Index), getResByKey("AnteModelInfo"))

        Dim cellRang As range
        modelValueStr = referencedString
        Set cellRang = ThisWorkbook.ActiveSheet.range(charStr + "2:" + charStr + CStr(2 + ROW_COUNT + 3))
            With cellRang.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=referencedString
            End With
        ThisWorkbook.ActiveSheet.Columns(Model_Index).EntireColumn.AutoFit
End Sub

'插入基带设备编号列
'Private Function insertBaseEqmColumn() As Boolean
'        Dim mocName As String
'        Dim attrName As String
'        Dim sectorColumnName As String
'        Dim columnIndex As Long
'        Dim cellsheet As Worksheet
'
'        Set cellsheet = ThisWorkbook.Worksheets(SEC_GRP_SHEET_NAME)
'        Call getBaseEqmMocNameAndAttrName(mocName, attrName)
'        'sectorColumnName = findColumnFromRelationDef(SEC_GRP_SHEET_NAME, mocName, attrName)
'        columnIndex = getColNum(SEC_GRP_SHEET_NAME, 2, attrName, mocName)
'        If columnIndex < 0 Then
'            insertBaseEqmColumn = False
'            Exit Function
'        End If
'        BaseEqm_Index = 6
'        cellsheet.Cells(2, columnIndex).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, BaseEqm_Index)
'        ThisWorkbook.ActiveSheet.Cells(1, BaseEqm_Index).Comment.text text:=getResByKey("LTEBASEBANDEQMID")
'        ThisWorkbook.ActiveSheet.Columns(BaseEqm_Index).EntireColumn.AutoFit
'        insertBaseEqmColumn = True
'End Function

'调整临时页签格式
Private Sub AdjustTemSheetStyle()
    Dim tmpsheet As Worksheet
    Dim sheetRange As range
    Dim maxColLen As Long
    Set tmpsheet = ThisWorkbook.Worksheets(getResByKey("Temp_Adjust_Sheet"))
    
    With tmpsheet.Cells.Font
        .name = "Arial"
        .Size = 10
    End With
    maxColLen = getColMaxLength(tmpsheet)
    With tmpsheet.UsedRange
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With
    
    '刷新批注，自适应调整批注框大小
    Dim maxColumnNumber As Long
    maxColumnNumber = tmpsheet.range("XFD1").End(xlToLeft).column
    Call refreshComment(tmpsheet.range(tmpsheet.Cells(1, 1), tmpsheet.Cells(1, maxColumnNumber)), True)
End Sub

Private Function getColMaxLength(sheet As Worksheet) As Long
    getColMaxLength = sheet.range("XFD1").End(xlToLeft).column
End Function

'插入扇区设备组标识
Private Function insertSectorEqmGrpIdColumn() As Boolean
        Dim mocName As String
        Dim attrName As String
        Dim sectorColumnName As String
        Dim columnIndex As Long
        Dim cellsheet As Worksheet
        Dim referencedString As String
        Dim charStr As String
        
        referencedString = ""
        
        charStr = "C"
        
        Set cellsheet = ThisWorkbook.Worksheets(SEC_GRP_SHEET_NAME)
        Call getSecGrpSectoreqmGrpIdMocNameAndAttrName(mocName, attrName)
        If SEC_GRP_SHEET_NAME = "EUPRBSECTOREQMGROUP" Or SEC_GRP_SHEET_NAME = getResByKey("EUPRBSECTOREQMGROUP") Then
            mocName = "EuPrbSectorEqmGroup"
        End If
        'sectorColumnName = findColumnFromRelationDef(SEC_GRP_SHEET_NAME, mocName, attrName)
        columnIndex = getColNum(SEC_GRP_SHEET_NAME, 2, attrName, mocName)
        If columnIndex < 0 Then
            insertSectorEqmGrpIdColumn = False
            SectorEqmGrpId_Index = 0
            Exit Function
        End If
        SectorEqmGrpId_Index = 3
        cellsheet.Cells(2, columnIndex).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, SectorEqmGrpId_Index)
        'ThisWorkbook.ActiveSheet.Cells(1, SectorEqmGrpId_Index).Comment.text text:=getResByKey("SECTOREQMCOMBINEGRPID")
        
        
        Dim index As Long
        Dim temBtsName As Variant
        
        For Each temBtsName In selectBtsNameCol
            For index = 2 To cellsheet.range("a1048576").End(xlUp).row
                If cellsheet.Cells(index, 1).value = CStr(temBtsName) And cellsheet.Cells(index, 2).value <> "RMV" Then
                    If referencedString <> "" Then
                        referencedString = referencedString + "," + cellsheet.Cells(index, columnIndex).value
                    Else
                        referencedString = cellsheet.Cells(index, columnIndex).value
                    End If
                End If
            Next
        Next
        
'        Dim cellRang As range
        sectoreqmValueStr = referencedString
'        Set cellRang = ThisWorkbook.ActiveSheet.range(charStr + "2:" + charStr + CStr(2 + ROW_COUNT + 3))
'        With cellRang.Validation
'            .Delete
'            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=referencedString
'        End With
        
        ThisWorkbook.ActiveSheet.Columns(SectorEqmGrpId_Index).EntireColumn.AutoFit
        
        insertSectorEqmGrpIdColumn = True
End Function

'插入子卡免配置列
Private Sub insertVMprfAutoAfgColumn()
         Dim charStr As String
         Dim referencedString As String
         Dim cellsheet As Worksheet
         Dim mocName As String
         Dim attrName As String
         Dim columnIndex As Long
         
        charStr = "H"
        referencedString = "Y,NA"
        
        Set cellsheet = ThisWorkbook.Worksheets(SEC_GRP_SHEET_NAME)
        Call getVMprfAutoAfgMocNameAndAttrName(mocName, attrName)
        If SEC_GRP_SHEET_NAME = "EUPRBSECTOREQMGROUP" Or SEC_GRP_SHEET_NAME = getResByKey("EUPRBSECTOREQMGROUP") Then
            mocName = "EuPrbSectorEqmGroup"
        End If
        'sectorColumnName = findColumnFromRelationDef(SEC_GRP_SHEET_NAME, mocName, attrName)
        columnIndex = getColNum(SEC_GRP_SHEET_NAME, 2, attrName, mocName)
        
        VMprfAutoAfg_Index = 8
        cellsheet.Cells(2, columnIndex).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, VMprfAutoAfg_Index)
        'ThisWorkbook.ActiveSheet.Cells(1, Model_Index).value = getResByKey("anteModel")
        'ThisWorkbook.ActiveSheet.Cells(1, Model_Index).Comment.text text:=getResByKey("VMPRFAUTOCFG")

        Dim cellRang As range
        Set cellRang = ThisWorkbook.ActiveSheet.range(charStr + "2:" + charStr + CStr(2 + ROW_COUNT + 3))
            With cellRang.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=referencedString
            End With
        ThisWorkbook.ActiveSheet.Columns(Model_Index).EntireColumn.AutoFit
End Sub
'向临时页签写入数据
Private Sub writeDataToTemSheet(cellId_secEqmGrpIdMap As Collection)
    'Dim tempColl As Collection
    Dim cell As CLampsiteAntennes
    Dim row As Variant
    Dim index As Long
    Dim sflag As Boolean
    Dim boardStr As String
    Dim antnStr As String
    Dim rsModel As String
    Dim tempSheet As Worksheet
    Dim antenneCollection As Collection
    Dim rangeStr As String
    Dim rowVal As Long
    Dim valueMapIndex As Long
    Dim temIndex As Long
    Dim cellId As Variant
    Dim temKeyVal As String
    sflag = True
    Set tempSheet = ThisWorkbook.Worksheets(getResByKey("Temp_Adjust_Sheet"))
    index = 2
    
    For Each cellId In Cell_Row_Map.KeyCollection
        If Contains(valueMap, CStr(cellId)) Then
            Dim tempColl As Collection
            For Each tempColl In valueMap(CStr(cellId))
                For Each cell In tempColl
                    If cell.antennes = "" Or cell.sectorIds = "" Or cell.sectorEqmGrpIds = "" Then
                        GoTo NextLoop
                    End If
            
                    Set antenneCollection = cell.getAntenneCollection(sflag, rowVal)
                    If sflag = False Then
                        Call deleteSecGrpTempSheet
                        ThisWorkbook.Worksheets(SEC_GRP_SHEET_NAME).rows(rowVal).Select
                        Exit Sub
                    End If
                    For Each row In antenneCollection
                        boardStr = row(3)
                        antnStr = row(4)
                        rsModel = row(5)
                        rangeStr = row(6)
                        If checkBoardBasedValid(boardStr) = False Or (Trim(antnStr) <> "" And checkDataValid(antnValueStr, antnStr) = False) _
                            Or (Trim(rsModel) <> "" And checkDataValid(modelValueStr, rsModel) = False) Then
                            Call MsgBox(getResByKey("wrongDataInput") & rangeStr, vbInformation, getResByKey("Warning"))
                            Call deleteSecGrpTempSheet
                            ThisWorkbook.Worksheets(SEC_GRP_SHEET_NAME).range(rangeStr).Select
                            Exit Sub
                        End If
                        tempSheet.Cells(index, BTSName_Index).value = row(0)
                        tempSheet.Cells(index, Cell_Index).value = row(1)
                            
                        temKeyVal = row(0) + "*" + row(1)
                        
                        tempSheet.Cells(index, Sector_Index).value = row(2)
                        tempSheet.Cells(index, Board_Index).value = boardStr
                        tempSheet.Cells(index, Ante_Index).value = antnStr
                        tempSheet.Cells(index, Model_Index).value = rsModel
                        '写扇区设备组Id
                        tempSheet.Cells(index, SectorEqmGrpId_Index).value = row(10)
                        '更新单元格下拉列表，根据CellId更新
'                        Call updateSecGrpIdValidList(tempSheet, index, SectorEqmGrpId_Index, temKeyVal)
                        '写子卡免配置
                        tempSheet.Cells(index, VMprfAutoAfg_Index).value = row(11)
                        index = index + 1
                    Next
NextLoop:
                Next cell
            Next
        End If
    Next
End Sub

Public Sub updateSecGrpIdValidListMain(ByVal sheet As Worksheet, ByVal target As range)
    If target.rows.count <> 1 Or target.Columns.count <> 1 Then Exit Sub
    '所选单元格为扇区设备组id列
    If target.column = SectorEqmGrpId_Index Then
        Dim temCellId As String
        Dim temBtsName As String
        Dim temKey As String
        
        temBtsName = sheet.Cells(target.row, BTSName_Index)
        temCellId = sheet.Cells(target.row, Cell_Index)
        temKey = temBtsName + "*" + temCellId
        Call updateSecGrpIdValidList(sheet, target.row, target.column, temKey)
    End If
    
End Sub

Private Sub updateSecGrpIdValidList(tempSheet As Worksheet, row_index As Long, col_index As Long, temCellId As String)
    Dim sectorEqmGrpIdStr As String
    sectorEqmGrpIdStr = ""
    
    If Contains(cellId_secEqmGrpIdMap, temCellId) Then
        Dim temVal As Variant
        For Each temVal In cellId_secEqmGrpIdMap(temCellId)
            If sectorEqmGrpIdStr = "" Then
                sectorEqmGrpIdStr = CStr(temVal)
            Else
                sectorEqmGrpIdStr = sectorEqmGrpIdStr + "," + CStr(temVal)
            End If
        Next
    End If
    
    With tempSheet.Cells(row_index, col_index).Validation
        .Delete
        .Add Type:=xlValidateList, formula1:=sectorEqmGrpIdStr
    End With
End Sub

Private Function checkBoardBasedValid(tVal As String) As Boolean
        If tVal = "" Then
            checkBoardBasedValid = True
            Exit Function
        End If
        Dim tCount As Long
        Dim strArray() As String
        strArray = Split(tVal, "_")
        tCount = UBound(strArray)
        If tCount = 3 Then
            checkBoardBasedValid = True
        Else
            checkBoardBasedValid = False
        End If
End Function
Private Function checkDataValid(tValueStr As String, tVal As String) As Boolean
        Dim index As Long
        Dim strArray() As String
        strArray = Split(tValueStr, ",")
        For index = LBound(strArray) To UBound(strArray)
              If strArray(index) = tVal Then
                checkDataValid = True
                Exit Function
                End If
        Next
        checkDataValid = False
End Function

Public Function isLampsiteLteSectoreqmGrpId() As Boolean
    Dim actSheetName As String
    Dim cellsheet As Worksheet
    Dim mocName As String
    Dim attrName As String
    Dim sectorColumnName As String
    Dim columnIndex As Long
    isLampsiteLteSectoreqmGrpId = False
    'actSheetName = getResByKey("A176")
    For Each cellsheet In ThisWorkbook.Worksheets
        Call getSecGrpSectoreqmGrpIdMocNameAndAttrName(mocName, attrName)
        If SEC_GRP_SHEET_NAME = "EUPRBSECTOREQMGROUP" Or SEC_GRP_SHEET_NAME = getResByKey("EUPRBSECTOREQMGROUP") Then
            mocName = "EuPrbSectorEqmGroup"
        End If
        'sectorColumnName = findColumnFromRelationDef(actSheetName, mocName, attrName)
        columnIndex = getColNum(SEC_GRP_SHEET_NAME, 2, attrName, mocName)
        If columnIndex > 0 Then
            isLampsiteLteSectoreqmGrpId = True
            Exit Function
        End If
    Next
End Function

'回写数据到小区扇区设备组
Private Sub writeBackDataToSecGrp()
    On Error GoTo ErrorHandler
    
    'DTS2019050703315，回写数据前清空粘贴板
    Application.CutCopyMode = False
    Dim cellInfoMap As CMapValueObject
    Dim error As Boolean
    Set errCollect = New Collection
    If WRITESUCCESS = True Then
        error = checkUserData()
        If error = False Then
            Exit Sub
        End If
'    If ANT_CELL_TYPE = 2 Then
    Call writeCellData
'        ElseIf CELL_TYPE = 1 And SectorEqm_Index <> 0 Then
'            Call writeUMTSCellData
'        Else
'            Set cellInfoMap = genCellInfoMap()
'            Call sortMapByKey(cellInfoMap, error)
'            Call writeCellData(cellInfoMap)
'        End If
    End If
    WRITESUCCESS = False
    Call deleteSecGrpTempSheet
    Exit Sub
ErrorHandler:
    Call deleteSecGrpTempSheet
End Sub

'检查用户填写数据是否符合
Private Function checkUserData() As Boolean
    Dim tmpsheet As Worksheet
    Dim errRangeCol As Collection
    Dim maxRow As Long
    Dim rowNum As Long
    Dim cellStr As String
    Dim sectorStr As String
    Dim boardStr As String
    Dim trxStr As String
    Dim antnStr As String
    Dim modelStr As String
    Dim baseEqmStr As String
    Dim sectoreqmproStr As String
    Dim vmprfAutoAfg As String
    Dim btsNameStr As String
    
    Dim errflag As Boolean
    Dim lineStr As String
    Dim tVal As Variant
    Dim eRange As range
    Dim dupCol As Collection
    Dim keyStr As String
    Dim umtSkeyStr As String
    Dim cellTrxMap As CMapValueObject  '校验GSM小区的频点必须都要配置
    'Dim trxSectorMap As New CMapValueObject '校验GSM共小区情况下每个频点配置的扇区都要一致
    
    Dim sectorBaseEqmMap As CMapValueObject
    Dim sectorBaseEqm As String
    Set sectorBaseEqmMap = New CMapValueObject
    
    Set cellTrxMap = New CMapValueObject
    checkUserData = True
    Set tmpsheet = ThisWorkbook.Worksheets(getResByKey("Temp_Adjust_Sheet"))
    Set errRangeCol = New Collection
    Set dupCol = New Collection
    
    maxRow = tmpsheet.range("a1048576").End(xlUp).row
    For rowNum = 2 To maxRow
        btsNameStr = tmpsheet.Cells(rowNum, BTSName_Index).value
        cellStr = tmpsheet.Cells(rowNum, Cell_Index).value
        sectorStr = tmpsheet.Cells(rowNum, Sector_Index).value
        boardStr = tmpsheet.Cells(rowNum, Board_Index).value
        antnStr = tmpsheet.Cells(rowNum, Ante_Index).value

        modelStr = tmpsheet.Cells(rowNum, Model_Index).value
        vmprfAutoAfg = tmpsheet.Cells(rowNum, VMprfAutoAfg_Index).value

        errflag = checkInputData(tmpsheet, rowNum, btsNameStr, cellStr, sectorStr, boardStr, antnStr, modelStr, errRangeCol)
        If errflag Then
            '用于检查扇区ID和子卡免配置的一一对应，相同的扇区ID，子卡免配置相同
            'sectorBaseEqm = cellStr + "_" + sectorStr
            sectorBaseEqm = boardStr
            Dim bMap As CMap
            Set bMap = New CMap
            If sectorBaseEqmMap.hasKey(sectorBaseEqm) Then
                Set bMap = sectorBaseEqmMap.GetAt(sectorBaseEqm)
                sectorBaseEqmMap.RemoveKey (sectorBaseEqm)
            End If
            Call bMap.SetAt(rowNum, vmprfAutoAfg)
            Call sectorBaseEqmMap.SetAt(sectorBaseEqm, bMap)

            keyStr = btsNameStr + "_" + cellStr + "_" + sectorStr + "_" + boardStr + "_" + antnStr + "_" + modelStr

            If Contains(dupCol, keyStr) Then
                lineStr = dupCol(keyStr) + "," + str(rowNum)
                dupCol.Remove (keyStr)
            Else
                lineStr = str(rowNum)
            End If
            dupCol.Add Item:=lineStr, key:=keyStr
        End If
    Next

    If errRangeCol.count() <> 0 Then
        Call MsgBox(getResByKey("recordError"), vbInformation, getResByKey("Warning"))
        For Each tVal In errRangeCol
            Set eRange = tmpsheet.range(tVal)
            eRange.Interior.colorIndex = 3
        Next
        Set errCollect = errRangeCol
        checkUserData = False
        Exit Function
    End If

    Set dupCollect = New Collection

    If CheckDataDuplicate(tmpsheet, dupCol) = False Then
        checkUserData = False
        Exit Function
    End If

    If CheckLteCellValid(tmpsheet, sectorBaseEqmMap) = False Then
        checkUserData = False
        Exit Function
    End If
    

End Function

Private Function CheckLteCellValid(tmpsheet As Worksheet, sectorBaseEqmMap As CMapValueObject) As Boolean
        Dim aVal As Variant
        Dim tMap As CMap
        Dim rowErrStr As String
        Dim errReportStr As String
        CheckLteCellValid = True
        
        For Each aVal In sectorBaseEqmMap.KeyCollection
            Set tMap = sectorBaseEqmMap.GetAt(aVal)
            Dim rowList As String
            rowList = ""
            If existDiffData(tMap, rowList) = True Then
                dupCollect.Add Item:=rowList, key:=rowList
                If rowErrStr = "" Then
                    rowErrStr = rowList
                    errReportStr = rowList
                Else
                    rowErrStr = rowErrStr + "," + rowList
                    errReportStr = errReportStr + ";" + rowList
                End If
            End If
        Next
        
        Dim lineSet As Variant
        Dim maxColLen As Long
        Dim tVal As Variant
        Dim rowNum  As Long
        maxColLen = getColMaxLength(tmpsheet)
        If rowErrStr <> "" Then
            'Call MsgBox(getResByKey("recordNotConsis") + errReportStr, vbInformation, getResByKey("Warning"))
            Call MsgBox(getResByKey("VMPRFAUTOCFGNotConsis"), vbInformation, getResByKey("Warning"))
            lineSet = Split(rowErrStr, ",")
            For Each tVal In lineSet
                rowNum = CLng(tVal)
                tmpsheet.range(Cells(rowNum, Cell_Index), Cells(rowNum, maxColLen)).Interior.colorIndex = 3
            Next
            CheckLteCellValid = False
            Exit Function
        End If
        
End Function

Private Function existDiffData(aMap As CMap, rowList As String) As Boolean
    existDiffData = False
    Dim keyVal As Variant
    Dim tVal As Variant
    
    Dim dataCollect As Collection
    Set dataCollect = New Collection
    For Each keyVal In aMap.KeyCollection
        tVal = aMap.GetAt(keyVal)
        If Not Contains(dataCollect, CStr(tVal)) Then
            dataCollect.Add Item:=CStr(tVal), key:=CStr(tVal)
        End If
        If rowList = "" Then
            rowList = CStr(keyVal)
        Else
            rowList = rowList + "," + CStr(keyVal)
        End If
    Next
    
    If dataCollect.count() <> 1 Then
        existDiffData = True
        Exit Function
    End If
End Function

Private Function checkInputData(sheet As Worksheet, lineNo As Long, btsName As String, cell As String, sector As String, board As String, antn As String, model As String, errRangeCol As Collection) As Boolean
        checkInputData = True
        Dim rangeStr As String
        
        If btsName = "" Or existInCollection(btsName, selectBtsNameCol) = False Then
            rangeStr = sheet.Cells(lineNo, BTSName_Index).address(False, False)
            errRangeCol.Add Item:=rangeStr, key:=rangeStr
            checkInputData = False
        End If
        
        If checkDataValid(cellValueStr, cell) = False Then
            rangeStr = sheet.Cells(lineNo, Cell_Index).address(False, False)
            errRangeCol.Add Item:=rangeStr, key:=rangeStr
            checkInputData = False
        End If
    
        If Trim(sector) = "" Or isAInteger(sector) = False Then
            rangeStr = sheet.Cells(lineNo, Sector_Index).address(False, False)
            errRangeCol.Add Item:=rangeStr, key:=rangeStr
            checkInputData = False
        End If
        
        'Lampsite特殊处理，槽位号为255自动替换为0处理。
        Dim temBoard As String
        temBoard = board
        board = getReplaceString(temBoard)
'        If checkDataValid(boardValueStr, board) = False Then
'            rangeStr = sheet.Cells(lineNo, Board_Index).address(False, False)
'            errRangeCol.Add Item:=rangeStr, key:=rangeStr
'            checkInputData = False
'        End If
            
        If checkDataValid(antnValueStr, antn) = False Then
            rangeStr = sheet.Cells(lineNo, Ante_Index).address(False, False)
            errRangeCol.Add Item:=rangeStr, key:=rangeStr
            checkInputData = False
        End If
        
        If checkDataValid(modelValueStr, model) = False Then
            rangeStr = sheet.Cells(lineNo, Model_Index).address(False, False)
            errRangeCol.Add Item:=rangeStr, key:=rangeStr
            checkInputData = False
        End If
        
'        If CELL_TYPE = 2 Then
'            If Trim(sector) = "" Or isAInteger(baseEqm) = False Then
'                rangeStr = sheet.Cells(lineNo, BaseEqm_Index).address(False, False)
'                errRangeCol.Add Item:=rangeStr, key:=rangeStr
'                checkInputData = False
'            End If
'        End If
        
                
'        If CELL_TYPE = 1 And SectorEqm_Index <> 0 Then
'            If checkDataValid(sectoreqmValueStr, sectoreqmproStr) = False Then
'                rangeStr = sheet.Cells(lineNo, SectorEqm_Index).address(False, False)
'                errRangeCol.Add Item:=rangeStr, key:=rangeStr
'                checkInputData = False
'            End If
'        End If
        
        
End Function

Private Function getReplaceString(temBoard) As String
    Dim temBoardArr() As String
    If temBoard <> "" Then
        temBoardArr = Split(temBoard, "_")
        If UBound(temBoardArr) = 3 Then
            If temBoardArr(2) = "255" Then
                temBoardArr(2) = "0"
            End If
        End If
        Dim index As Long
        For index = LBound(temBoardArr) To UBound(temBoardArr)
            If getReplaceString = "" Then
                getReplaceString = temBoardArr(index)
            Else
                getReplaceString = getReplaceString + "_" + temBoardArr(index)
            End If
        Next
    End If
End Function

Private Function isAInteger(ByRef tVal As String) As Boolean
    On Error GoTo ErrorHandler
    Dim k As Long
    If InStr(tVal, ".") <> 0 Then GoTo ErrorHandler
    k = CLng(tVal)
    If k <= 1048576 And k >= 0 Then
        isAInteger = True
    Else
        isAInteger = False
    End If
    Exit Function
ErrorHandler:
    isAInteger = False
End Function

Private Function CheckDataDuplicate(tmpsheet As Worksheet, dupCol As Collection) As Boolean
    CheckDataDuplicate = True
    
    Dim lineSet As Variant
    Dim selectStr As String
    Dim dupStr As String
    Dim rowNum As Long
    Dim tVal As Variant
    
    For Each tVal In dupCol
        If InStr(1, tVal, ",") <> 0 Then
            dupCollect.Add Item:=tVal, key:=tVal
            If dupStr = "" Then
                dupStr = tVal
                selectStr = tVal
            Else
                dupStr = dupStr + ";" + tVal
                selectStr = selectStr + "," + tVal
            End If
        End If
    Next
    
    Dim maxColLen As Long
    maxColLen = getColMaxLength(tmpsheet)
    If dupStr <> "" Then
        Call MsgBox(getResByKey("recordDuplicate") + dupStr, vbInformation, getResByKey("Warning"))
        lineSet = Split(selectStr, ",")
        For Each tVal In lineSet
            rowNum = CLng(tVal)
            tmpsheet.range(Cells(rowNum, Cell_Index), Cells(rowNum, maxColLen)).Interior.colorIndex = 3
        Next
        CheckDataDuplicate = False
        Exit Function
    End If
    
End Function

Private Sub writeCellData()
    Dim tmpsheet As Worksheet
    Dim cellsheet As Worksheet
    Dim mocName As String
    Dim attrName As String
    Dim columnName As String
    Dim antnIndex As Long
    Dim secIndex As Long
    Dim baseEqmIndex As Long
    Dim sectorEqmGrpIdIndex As Long
    Dim cellBeamModeIndex As Long
    Dim vmprfAutoCfgIndex As Long
    
    
    Set tmpsheet = ThisWorkbook.Worksheets(getResByKey("Temp_Adjust_Sheet"))
    
    Set cellsheet = ThisWorkbook.Worksheets(SEC_GRP_SHEET_NAME)
    
    '找小区ID所在的列
'    Dim constCellTempCol As Long
'    Call getCellMocNameAndAttrName(mocName, attrName)
'    constCellTempCol = getColNum(SEC_GRP_SHEET_NAME, 2, attrName, mocName)
    
    '获取天线端口所在列
    Call getAntenneMocNameAndAttr(mocName, attrName)
    If SEC_GRP_SHEET_NAME = "EUPRBSECTOREQMGROUP" Or SEC_GRP_SHEET_NAME = getResByKey("EUPRBSECTOREQMGROUP") Then
        mocName = "EuPrbSectorEqmGroup"
    End If
    antnIndex = getColNum(SEC_GRP_SHEET_NAME, 2, attrName, mocName)
    '获取扇区所在列
    Call getSectorMocNameAndAttr(mocName, attrName)
    If SEC_GRP_SHEET_NAME = "EUPRBSECTOREQMGROUP" Or SEC_GRP_SHEET_NAME = getResByKey("EUPRBSECTOREQMGROUP") Then
        mocName = "EuPrbSectorEqmGroup"
    End If
    secIndex = getColNum(SEC_GRP_SHEET_NAME, 2, attrName, mocName)
    
    '获取子卡免配置列
    Call getVMprfAutoAfgMocNameAndAttrName(mocName, attrName)
    If SEC_GRP_SHEET_NAME = "EUPRBSECTOREQMGROUP" Or SEC_GRP_SHEET_NAME = getResByKey("EUPRBSECTOREQMGROUP") Then
        mocName = "EuPrbSectorEqmGroup"
    End If
    vmprfAutoCfgIndex = getColNum(SEC_GRP_SHEET_NAME, 2, attrName, mocName)
    
'    If ANT_CELL_TYPE = 2 Then
    Call getSecGrpSectoreqmGrpIdMocNameAndAttrName(mocName, attrName)
    If SEC_GRP_SHEET_NAME = "EUPRBSECTOREQMGROUP" Or SEC_GRP_SHEET_NAME = getResByKey("EUPRBSECTOREQMGROUP") Then
            mocName = "EuPrbSectorEqmGroup"
        End If
    sectorEqmGrpIdIndex = getColNum(SEC_GRP_SHEET_NAME, 2, attrName, mocName)
'    End If
    
'    Call getAntenneMocNameAndAttr(mocName, attrName)
'    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
'    antnIndex = findColumnByName(cellsheet, columnName, 2)
'    '获取扇区所在列
'    Call getSectorMocNameAndAttr(mocName, attrName)
'    columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
'    secIndex = findColumnByName(cellsheet, columnName, 2)
'
'    If CELL_TYPE = 2 Then
'        Call getBaseEqmMocNameAndAttrName(mocName, attrName)
'        columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
'        baseEqmIndex = findColumnByName(cellsheet, columnName, 2)
'
'        Call getSectoreqmGrpIdMocNameAndAttrName(mocName, attrName)
'        columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
'        sectorEqmGrpIdIndex = findColumnByName(cellsheet, columnName, 2)
'
'        Call getCellbeamModeMocNameAndAttrName(mocName, attrName)
'        columnName = findColumnFromRelationDef(CELL_SHEET_NAME, mocName, attrName)
'        cellBeamModeIndex = findColumnByName(cellsheet, columnName, 2)
'
'    End If
    
    Dim sVal As Variant
    For Each sVal In Cell_Row_Map.KeyCollection
        Call writeOneCell(tmpsheet, cellsheet, CStr(sVal), secIndex, antnIndex, sectorEqmGrpIdIndex, vmprfAutoCfgIndex)
    Next
End Sub

Private Sub writeOneCell(tmpsheet As Worksheet, cellsheet As Worksheet, temKey As String, _
                                                    secIndex As Long, antnIndex As Long, sectorEqmGrpIdIndex As Long, vmprfAutoCfgIndex As Long)
    Dim maxRow As Long
    Dim rowNum As Long
    Dim sectorId As String
    Dim baseEqmId As String
    Dim boardAnte As String
    Dim sectorIds As String
    Dim sectorEqmGrpId As String
    Dim sectorEqmGrpIdStr As String
    Dim cellbeamMode As String
    Dim vmprfAutoCfgStr As String
    Dim vmprfAutoCfgStrs As String
    Dim temKeyArr() As String
    
    sectorIds = ""
    vmprfAutoCfgStr = ""
    
    
    Dim cellInRow As String
    cellInRow = Cell_Row_Map.GetAt(temKey)
    
    Dim sectorEqmGrpIdMap As CMap
    Set sectorEqmGrpIdMap = New CMap
    
    Dim sectorMap As CMap
    Set sectorMap = New CMap
    
    Dim vmprfAutoCfgMap As CMap
    Set vmprfAutoCfgMap = New CMap
    
    temKeyArr = Split(temKey, "*")
    If UBound(temKeyArr) < 0 Then Exit Sub
    
    maxRow = tmpsheet.range("a1048576").End(xlUp).row
    '同一小区下有多个扇区设备组标识
    For rowNum = 2 To maxRow
         If tmpsheet.Cells(rowNum, BTSName_Index).value = temKeyArr(0) And tmpsheet.Cells(rowNum, Cell_Index).value = temKeyArr(1) Then
            sectorId = tmpsheet.Cells(rowNum, Sector_Index).value
            sectorEqmGrpId = tmpsheet.Cells(rowNum, SectorEqmGrpId_Index).value
            If sectorEqmGrpIdMap.hasKey(sectorEqmGrpId) = False Then
                Call sectorEqmGrpIdMap.SetAt(sectorEqmGrpId, sectorId)
            End If
        End If
    Next
    '扇区设备组标识字符串拼接
    Dim sIndex As Variant
    For Each sIndex In sectorEqmGrpIdMap.KeyCollection
        If sectorEqmGrpIdStr = "" Then
            sectorEqmGrpIdStr = CStr(sIndex)
        Else
            sectorEqmGrpIdStr = sectorEqmGrpIdStr + "," + CStr(sIndex)
        End If
    Next
    
    Dim sectorEqmGrpIdArray() As String
    sectorEqmGrpIdArray = Split(sectorEqmGrpIdStr, ",")
    
    
    For rowNum = 2 To maxRow
         If tmpsheet.Cells(rowNum, BTSName_Index).value = temKeyArr(0) And tmpsheet.Cells(rowNum, Cell_Index).value = temKeyArr(1) Then
            sectorId = tmpsheet.Cells(rowNum, Sector_Index).value
            'vmprfAutoCfgStr = tmpsheet.Cells(RowNum, VMprfAutoAfg_Index).value
            'sectorEqmGrpId = tmpsheet.Cells(RowNum, SectorEqmGrpId_Index).value
            If sectorMap.hasKey(sectorId) = False Then
                Call sectorMap.SetAt(sectorId, sectorId)
            End If
'            If vmprfAutoCfgMap.haskey(sectorId) = False Then
'                Call vmprfAutoCfgMap.SetAt(sectorId, vmprfAutoCfgStr)
'            End If
        End If
    Next
    
    '扇区ID字符串拼接
    Dim sVal As Variant
    For Each sVal In sectorMap.KeyCollection
        If sectorIds = "" Then
            sectorIds = CStr(sVal)
        Else
            sectorIds = sectorIds + "," + CStr(sVal)
        End If
    Next
    
    '子卡免配置字符串拼接
'    Dim sLoop As Variant
'    For Each sLoop In vmprfAutoCfgMap.ValueCollection
'        If vmprfAutoCfgStrs = "" Then
'            vmprfAutoCfgStrs = CStr(sLoop)
'        Else
'            vmprfAutoCfgStrs = vmprfAutoCfgStrs + "," + CStr(sLoop)
'        End If
'    Next
    
    getSortedStr (sectorIds)
    'getSortedStr (vmprfAutoCfgStrs)
    
    
    Dim sectorArray() As String
    
    
    Dim antenna As String
    Dim antaStr As String
    antaStr = ""
    
    sectorArray = Split(sectorIds, ",")
    'sectorEqmGrpIdArray = Split(sectorEqmGrpIdStr, ",")
    Dim index As Long
    Dim sectorIndex As Long
    Dim sectorIdStr As String
    Dim vmprfAutoCfgs As String
    Dim vmprfAutoCfg As String
    For sectorIndex = LBound(sectorEqmGrpIdArray) To UBound(sectorEqmGrpIdArray)
        sectorIdStr = ""
        vmprfAutoCfgs = ""
        antaStr = ""
        Dim sectorCMap As CMap
        Set sectorCMap = New CMap
        Dim temVmprfAutoCfgCMap As CMap
        Set temVmprfAutoCfgCMap = New CMap
        
        For index = LBound(sectorArray) To UBound(sectorArray)
            antenna = ""
            For rowNum = 2 To maxRow
                If tmpsheet.Cells(rowNum, BTSName_Index).value = temKeyArr(0) And tmpsheet.Cells(rowNum, Cell_Index).value = temKeyArr(1) And tmpsheet.Cells(rowNum, SectorEqmGrpId_Index).value = sectorEqmGrpIdArray(sectorIndex) Then
                    sectorId = tmpsheet.Cells(rowNum, Sector_Index).value
                    vmprfAutoCfg = tmpsheet.Cells(rowNum, VMprfAutoAfg_Index).value
                    If sectorId = sectorArray(index) Then
                        
                        'sectorEqmGrpId = tmpsheet.Cells(RowNum, SectorEqmGrpId_Index).value
                    
                        boardAnte = tmpsheet.Cells(rowNum, Board_Index).value + "_" + tmpsheet.Cells(rowNum, Ante_Index).value + ":" + tmpsheet.Cells(rowNum, Model_Index).value
                        If antenna = "" Then
                            antenna = boardAnte
                        Else
                            antenna = antenna + "," + boardAnte
                        End If
                        

                        If sectorCMap.hasKey(sectorId) = False Then
                            Call sectorCMap.SetAt(sectorId, sectorId)
                        End If
                        

                        If temVmprfAutoCfgCMap.hasKey(sectorId) = False Then
                            Call temVmprfAutoCfgCMap.SetAt(sectorId, vmprfAutoCfg)
                        End If
                    End If
                    

                    
'                    If sectorIdStr = "" Then
'                        sectorIdStr = sectorId
'                    Else
'                        sectorIdStr = sectorIdStr + "," + sectorId
'                    End If
'                    If vmprfAutoCfgs = "" Then
'                        vmprfAutoCfgs = vmprfAutoCfg
'                    Else
'                        vmprfAutoCfgs = vmprfAutoCfgs + "," + vmprfAutoCfg
'                    End If
                End If
                
            Next
            If antenna <> "" Then
                If antaStr = "" Then
                    antaStr = antenna
                Else
                    antaStr = antaStr + ";" + antenna
                End If
            End If
            
        Next
        
        Dim temVal As Variant
        For Each temVal In sectorCMap.KeyCollection
            If sectorIdStr = "" Then
                sectorIdStr = CStr(temVal)
            Else
                sectorIdStr = sectorIdStr + "," + CStr(temVal)
            End If
        Next

        Dim temLoop As Variant
        For Each temLoop In temVmprfAutoCfgCMap.ValueCollection
            If vmprfAutoCfgs = "" Then
                vmprfAutoCfgs = CStr(temLoop)
            Else
                vmprfAutoCfgs = vmprfAutoCfgs + "," + CStr(temLoop)
            End If
        Next
        
        cellsheet.Cells(cellInRow, sectorEqmGrpIdIndex).value = sectorEqmGrpIdArray(sectorIndex)
        cellsheet.Cells(cellInRow, secIndex).value = sectorIdStr
        cellsheet.Cells(cellInRow, antnIndex).value = antaStr
        cellsheet.Cells(cellInRow, vmprfAutoCfgIndex).value = vmprfAutoCfgs
        cellInRow = cellInRow + 1
    Next

End Sub

Private Sub getSortedStr(ByRef infoStr As String)
    Dim strArray() As Variant
    Dim tmpArray As Variant
    Dim index As Long
    Dim tCount As Long
    Dim tmpStr As String
    tmpArray = Split(infoStr, ",")
    tCount = UBound(tmpArray) - LBound(tmpArray)
    ReDim strArray(tCount) As Variant
    For index = LBound(tmpArray) To UBound(tmpArray)
        strArray(index) = tmpArray(index)
    Next
    Call QuickSort(strArray())
    tmpStr = ""
    
    For index = LBound(strArray) To UBound(strArray)
        If tmpStr = "" Then
            tmpStr = strArray(index)
        Else
            tmpStr = tmpStr & "," & strArray(index)
        End If
    Next
    infoStr = tmpStr
End Sub


Private Sub insertBtsNameColumn()
    
    Dim cellsheet As Worksheet
    Dim btsNameColIndex As Long
    Dim mocName As String
    Dim attrName As String
    
    Call getBaseStationMocNameAndAttrNameForLampsite(mocName, attrName)
    btsNameColIndex = getColNum(SEC_GRP_SHEET_NAME, 2, attrName, mocName)
    
    BTSName_Index = 1
    Set cellsheet = ThisWorkbook.Worksheets(SEC_GRP_SHEET_NAME)
    cellsheet.Cells(2, btsNameColIndex).Copy Destination:=ThisWorkbook.ActiveSheet.Cells(1, BTSName_Index)
End Sub

Private Function existInCollection(strValue As Variant, strCollection As Collection) As Boolean
    Dim sItem As Variant
    If Trim(CStr(strValue)) = "" Then
        existInCollection = True
        Exit Function
    End If
    For Each sItem In strCollection
        If sItem = strValue Then
            existInCollection = True
            Exit Function
        End If
    Next
    existInCollection = False
End Function

Public Sub insertRxuBoardListForLampsite(ByRef ws As Worksheet, ByRef cellRange As range)
    If cellRange.rows.count <> 1 And cellRange.Columns.count <> 1 Then Exit Sub
    If cellRange.column <> Board_Index Then Exit Sub
    
    Dim btsName As String
    btsName = ws.Cells(cellRange.row, 1)
    If btsName = "" Then
        cellRange.Validation.Delete
        Exit Sub
    End If
    
    Call insertRxuBoardListByBtsName(btsName, cellRange)
End Sub


Private Sub insertRxuBoardListByBtsName(ByRef btsName As String, ByRef cellRang As range)
    Dim brdStr As String
    If btsNameBrdNoMap.hasKey(btsName) Then
        brdStr = btsNameBrdNoMap.GetAt(btsName)
    Else
        brdStr = findBoardNosStrByBtsName(btsName)
        Call btsNameBrdNoMap.SetAt(btsName, brdStr)
    End If
    
    If Len(brdStr) = 0 Then Exit Sub
            
    '插入List
    With cellRang.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=brdStr
    End With

End Sub

Private Function findBoardNosStrByBtsName(ByRef btsName As String) As String
    findBoardNosStrByBtsName = ""
    
    Dim brdStyleSheetName As String
    Dim brdStyleSheet As Worksheet
    
    brdStyleSheetName = findBoardStyleSheetByBtsName(btsName)
    If brdStyleSheetName = "" Then
        Call MsgBox(getResByKey("NoBoradStyle"), vbInformation, getResByKey("Warning"))
        'Call deleteTempSheet
        Exit Function
    End If
    
    Set brdStyleSheet = ThisWorkbook.Worksheets(brdStyleSheetName)
    
    Dim grpCollection As Collection
    Set grpCollection = findBrdGroups

    findBoardNosStrByBtsName = findBoardNosStrByBtsName_i(brdStyleSheet, grpCollection, btsName)
    
End Function


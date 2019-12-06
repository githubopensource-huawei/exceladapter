Attribute VB_Name = "CellSub"
Option Explicit
'「eNodeB Radio Data」页记录起始行
Private Const constRecordRow = 2
Private Const cellMocName As String = "GLoCell"
Private Const attrName As String = "CellTemplateName"
Private Const CellType As String = "GSM Local Cell"

Private Const logicCellMocName As String = "GCELL"
Private Const logicAttrName As String = "TemplateName"
Private Const logicCellType As String = "GSM Logic Cell"

Private Const UCellMocName As String = "ULOCELL"
Private Const UAttrName As String = "CellTemplateName"
Private Const UCellType As String = "UMTS Local Cell"

Private Const logicUCellMocName As String = "CELL"
Private Const logicUAttrName As String = "TemplateName"
Private Const logicUCellType As String = "UMTS Logic Cell"

Private Const LCellMocName As String = "Cell"
Private Const LAttrName As String = "CellTemplateName"
Private Const LCellType As String = "LTE Cell"

Function isCellSheet(sheetName As String) As Boolean
    If sheetName = "GSM Cell" Or sheetName = "UMTS Cell" Or sheetName = "LTE Cell" _
        Or sheetName = getResByKey("GSMCellSheetName") Or sheetName = getResByKey("UMTSCellSheetName") Or sheetName = getResByKey("LTECellSheetName") Then
        isCellSheet = True
        Exit Function
    End If
    isCellSheet = False
End Function
 
Function isTrasnPortSheet(sheetName As String) As Boolean
    If sheetName = "GSM Cell" Or sheetName = "UMTS Cell" Or sheetName = "LTE Cell" _
        Or sheetName = getResByKey("GSMCellSheetName") Or sheetName = getResByKey("UMTSCellSheetName") Or sheetName = getResByKey("LTECellSheetName") _
        Or sheetName = "GTRXGROUP" Or sheetName = getResByKey("GTRXGroupSheetName") Then
        isTrasnPortSheet = True
        Exit Function
    End If
    isTrasnPortSheet = False
End Function

Public Function isEuCellSectorEqmSht(sheetName As String) As Boolean
    If sheetName = "EUCELLSECTOREQM" Or sheetName = getResByKey("EUCELLSECTOREQM") Then
        isEuCellSectorEqmSht = True
        Exit Function
    End If
    isEuCellSectorEqmSht = False
End Function

Public Function isEuPrbSectorEqmSht(sheetName As String) As Boolean
    If sheetName = "EUPRBSECTOREQM" Or sheetName = getResByKey("EUPRBSECTOREQM") Then
        isEuPrbSectorEqmSht = True
        Exit Function
    End If
    isEuPrbSectorEqmSht = False
End Function

Public Function isEuPrbSectorEqmGrpSht(sheetName As String) As Boolean
    If sheetName = "EUPRBSECTOREQMGROUP" Or sheetName = getResByKey("EUPRBSECTOREQMGROUP") Then
        isEuPrbSectorEqmGrpSht = True
        Exit Function
    End If
    isEuPrbSectorEqmGrpSht = False
End Function


'需要区分处理LTE和Mrat的差异
Public Sub cellSheetSelectionEvent(ByVal sheet As Worksheet, ByVal target As range)
    Dim neType As String
    neType = getNeType()
    If neType = "LTE" Then
        Call Cell_Worksheet_SelectionChange(sheet, target)
    Else 'Mrat,UMTS
        Call CellSelectionChange(sheet, target)
    End If
End Sub

'定义设置「Cell Template」列下拉列表的事件
Public Sub CellSelectionChange(ByVal sheet As Worksheet, ByVal target As range)
        'init 物理GCell
        Call initCellTemplate(sheet, target, attrName, cellMocName, getResByKey(CellType))
        'init 逻辑GCell
        Call initCellTemplate(sheet, target, logicAttrName, logicCellMocName, getResByKey(logicCellType))
        
        'init 物理UCell
        Call initCellTemplate(sheet, target, UAttrName, UCellMocName, getResByKey(UCellType))
        'init 逻辑UCell
        Call initCellTemplate(sheet, target, logicUAttrName, logicUCellMocName, getResByKey(logicUCellType))
        
        'init 物理LCell
        Call initCellTemplate(sheet, target, LAttrName, LCellMocName, getResByKey(LCellType))
End Sub

Sub initCellTemplate(ByVal sheet As Worksheet, ByVal target As range, myAttrName As String, myCellMocName As String, CellType As String)
        '如果是LTE小区，则进行LTE小区特有的按条件过滤筛选，其余小区页按原有流程
        If CellType = getResByKey(LCellType) Then
            Call initLteCellTemplate(sheet, target, CellType)
            Exit Sub
        End If
        
        Dim m_Cell_Template As String
        
        Dim constCellTempCol As Long
        '「物理Cell Template」所在列
        constCellTempCol = getColNum(sheet.name, constRecordRow, myAttrName, myCellMocName)

        If constCellTempCol >= 0 And target.column = constCellTempCol And target.count = 1 And target.row > constRecordRow Then
            '获取「CellTemplate」列侯选值
            m_Cell_Template = getCellTemplate(CellType, sheet, target)
            If m_Cell_Template <> "" Then
                With target.Validation
                   .Delete
                   .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=m_Cell_Template
                End With
                If Not target.Validation.value Then
                    target.value = ""
                End If
            Else
                With target.Validation
                    .Delete
                    .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .inputTitle = ""
                    .ErrorTitle = ""
                    .inputMessage = ""
                    .ErrorMessage = ""
                    .IMEMode = xlIMEModeNoControl
                    .ShowInput = True
                    .ShowError = True
                End With
                    'Target.value = ""
            End If
        End If
End Sub

Sub initLteCellTemplate(ByVal sh As Worksheet, ByVal target As range, ByVal CellType As String)
    Dim m_Cell_Template As String
    Dim constBandwidthCol As Long, constSACol As Long, constFDDTDDCol As Long, constTxRxModeCol As Long, constCellTempCol As Long, constNBIOTFlagCol As Long
    
    '「LTE Cell」页「*DlBandwidth」所在列
    constBandwidthCol = getColNum(sh.name, constRecordRow, "DlBandWidth", "Cell")
    
    '「LTE Cell」页「SubframeAssignment」所在列
    constSACol = getColNum(sh.name, constRecordRow, "SubframeAssignment", "Cell")
    
    '「LTE Cell」页「*FddTddInd」所在列
    constFDDTDDCol = getColNum(sh.name, constRecordRow, "FddTddInd", "Cell")
    
    '「LTE Cell」页「*TxRxMode」所在列
    constTxRxModeCol = getColNum(sh.name, constRecordRow, "TxRxMode", "Cell")
      
    '「LTE Cell」页「*Cell Template」所在列
    constCellTempCol = getColNum(sh.name, constRecordRow, "CellTemplateName", "Cell")
    
    '「LTE Cell」页「*NB-IoT TA Flag」所在列
    If getNBIOTFlag = True Then
        constNBIOTFlagCol = getColNum(sh.name, constRecordRow, "NbCellFlag", "Cell")
    Else
        constNBIOTFlagCol = -1
    End If
    
    Dim bandWidthValue As String
    Dim saValue As String
    Dim fddTddValue As String
    Dim txRxModeValue As String
    Dim nbIotValue As String
    
    If constBandwidthCol = -1 Then
         bandWidthValue = ""
     Else
         bandWidthValue = Cells(target.row, constBandwidthCol).value
     End If
     
     If constTxRxModeCol = -1 Then
         txRxModeValue = ""
     Else
         txRxModeValue = Cells(target.row, constTxRxModeCol).value
     End If
     
     If constFDDTDDCol = -1 Then
         fddTddValue = ""
     Else
         fddTddValue = Cells(target.row, constFDDTDDCol).value
     End If
     
     If constSACol = -1 Then
         saValue = ""
     Else
         saValue = Cells(target.row, constSACol).value
     End If
     
     If constNBIOTFlagCol = -1 Then
        nbIotValue = "False"
     Else
        nbIotValue = Cells(target.row, constNBIOTFlagCol).value
     End If
     

    If target.column = constCellTempCol And target.count = 1 And target.row > constRecordRow Then
         '获取「CellTemplate」列侯选值
         m_Cell_Template = Get_LteCellTemplate_Related(bandWidthValue, txRxModeValue, fddTddValue, saValue, nbIotValue, sh, target, CellType)
        If m_Cell_Template <> "" Then
            With target.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=m_Cell_Template
            End With
            If Not target.Validation.value Then
                target.value = ""
            End If
        Else
            With target.Validation
                .Delete
                .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
                .IgnoreBlank = True
                .InCellDropdown = True
                .inputTitle = ""
                .ErrorTitle = ""
                .inputMessage = ""
                .ErrorMessage = ""
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
        End If
    End If
End Sub

'新的从「MappingCellTemplate」页获取「Cell Template」列侯选值，LTE Cell根据条件过滤
Function Get_LteCellTemplate_Related(DlBandwidth As String, TxRxMode As String, FDDTDD As String, SA As String, NBIoTFlag As String, sheet As Worksheet, cellRange As range, ByRef CellType As String) As String
    Dim m_rowNum As Long
    Dim m_Str As String
    Dim m_start As Long
    Dim m_end As Long
    m_start = 0
    
    m_Str = ""
    
    Dim DlBandwidth1 As String
    Select Case DlBandwidth
        Case "CELL_BW_N6"
            DlBandwidth1 = "1.4M"
        Case "CELL_BW_N15"
            DlBandwidth1 = "3M"
        Case "CELL_BW_N25"
            DlBandwidth1 = "5M"
        Case "CELL_BW_N50"
            DlBandwidth1 = "10M"
        Case "CELL_BW_N75"
            DlBandwidth1 = "15M"
        Case "CELL_BW_N100"
            DlBandwidth1 = "20M"
        Case Else
            DlBandwidth1 = ""
    End Select
    
    Dim FDDTDD1 As String
    Select Case FDDTDD
        Case "CELL_TDD"
            FDDTDD1 = "TDD"
        Case "CELL_FDD"
            FDDTDD1 = "FDD"
        Case Else
            FDDTDD1 = ""
    End Select

    Dim neType As String
    neType = getNeType()
    m_Str = ""
    
    '这个新增的容器用来去重复用的，控制器和基站小区模板解耦可能会导致小区模板页有重复模板
    Dim cellTemplateCol As New Collection
    Dim cellTemplate As String
    
    Dim BandwidthCol As Long, TxRxModeCol As Long, FDDTDDCol As Long, SACol As Long, CellPatternCol As Long, CellTypeCol As Long, NETypeCol As Long
    Dim MappingCellTemplate As Worksheet
    Set MappingCellTemplate = ThisWorkbook.Worksheets("MappingCellTemplate")
    
    Call getCellTemplateCol(BandwidthCol, TxRxModeCol, FDDTDDCol, SACol, CellPatternCol, CellTypeCol, NETypeCol)
    
    For m_rowNum = 2 To MappingCellTemplate.range(getColStr(CellPatternCol) & "1048576").End(xlUp).row
        If (DlBandwidth1 = MappingCellTemplate.Cells(m_rowNum, BandwidthCol).value Or DlBandwidth1 = "" Or MappingCellTemplate.Cells(m_rowNum, BandwidthCol).value = "") _
        And (TxRxMode = MappingCellTemplate.Cells(m_rowNum, TxRxModeCol).value Or TxRxMode = "" Or MappingCellTemplate.Cells(m_rowNum, TxRxModeCol).value = "") _
        And (FDDTDD1 = MappingCellTemplate.Cells(m_rowNum, FDDTDDCol).value Or FDDTDD1 = "" Or MappingCellTemplate.Cells(m_rowNum, FDDTDDCol).value = "") _
        And (SA = MappingCellTemplate.Cells(m_rowNum, SACol).value Or SA = "" Or MappingCellTemplate.Cells(m_rowNum, SACol).value = "") _
        And CellType = MappingCellTemplate.Cells(m_rowNum, CellTypeCol) _
        And neType = MappingCellTemplate.Cells(m_rowNum, NETypeCol) And (UCase(NBIoTFlag) = UCase("FALSE") Or NBIoTFlag = "") _
        Or ((UCase(NBIoTFlag) = UCase("TRUE") And UCase(MappingCellTemplate.Cells(m_rowNum, FDDTDDCol).value) = UCase("NB-IoT"))) Then
            cellTemplate = MappingCellTemplate.Cells(m_rowNum, CellPatternCol).value
            If Contains(cellTemplateCol, cellTemplate) Or cellTemplate = "" Then GoTo NextLoop
            cellTemplateCol.Add Item:=cellTemplate, key:=cellTemplate
            If m_Str = "" Then
                 m_Str = cellTemplate
            Else
                 m_Str = m_Str & "," & cellTemplate
            End If
        End If
NextLoop:
    Next
    
     If Len(m_Str) > 255 Then
        Dim groupName As String
        Dim columnName As String
        Dim valideDef As CValideDef
        Call getGroupAndColumnName(sheet, cellRange, groupName, columnName)
        Set valideDef = initDefaultDataSub.getInnerValideDef(sheet.name + "," + groupName + "," + columnName)
        If valideDef Is Nothing Then
            Set valideDef = addInnerValideDef(sheet.name, groupName, columnName, m_Str)
        Else
            Call modiflyInnerValideDef(sheet.name, groupName, columnName, m_Str, valideDef)
        End If
        m_Str = valideDef.getValidedef
    End If
    Get_LteCellTemplate_Related = m_Str
    
End Function

'从「MappingCellTemplate」页获取「Cell Template」列侯选值
Function getCellTemplate(myType As String, sheet As Worksheet, cellRange As range) As String
    Dim m_rowNum As Long
    Dim m_RowNum_Inner As Long
    Dim flag As Boolean
    Dim m_Str As String
    Dim m_start As Long
    Dim m_end As Long
    Dim neType As String
    
    neType = getNeType()
    m_Str = ""
    
    '这个新增的容器用来去重复用的，控制器和基站小区模板解耦可能会导致小区模板页有重复模板
    Dim cellTemplateCol As New Collection
    Dim cellTemplate As String
    Dim MappingCellTemplate As Worksheet
    Set MappingCellTemplate = ThisWorkbook.Worksheets("MappingCellTemplate")

    
    
    For m_rowNum = 2 To MappingCellTemplate.range("C1048576").End(xlUp).row
        cellTemplate = MappingCellTemplate.Cells(m_rowNum, 1).value
        If myType = MappingCellTemplate.Cells(m_rowNum, 2).value _
        And neType = MappingCellTemplate.Cells(m_rowNum, 3).value _
        And (Not Contains(cellTemplateCol, cellTemplate)) And cellTemplate <> "" Then
            cellTemplateCol.Add Item:=cellTemplate, key:=cellTemplate
            If m_Str = "" Then
                 m_Str = cellTemplate
            Else
                 m_Str = m_Str & "," & cellTemplate
            End If
         End If
    Next
    
    If Len(m_Str) > 255 Then
        Dim groupName As String
        Dim columnName As String
        Dim valideDef As CValideDef
        Call getGroupAndColumnName(sheet, cellRange, groupName, columnName)
        Set valideDef = initDefaultDataSub.getInnerValideDef(sheet.name + "," + groupName + "," + columnName)
        If valideDef Is Nothing Then
            Set valideDef = addInnerValideDef(sheet.name, groupName, columnName, m_Str)
        Else
            Call modiflyInnerValideDef(sheet.name, groupName, columnName, m_Str, valideDef)
        End If
        m_Str = valideDef.getValidedef
    End If
    getCellTemplate = m_Str
End Function

'LTE小区模板的处理，需要根据前面某些输入值显示
'定义设置「Cell Template」列下拉列表的事件
Sub Cell_Worksheet_SelectionChange(ByVal sh As Worksheet, ByVal target As range)
    Dim m_Cell_Template As String
    
    Dim constBandwidthCol As Long, constSACol As Long, constFDDTDDCol As Long, constTxRxModeCol As Long, constCellTempCol As Long, constNBIOTFlagCol As Long
    '「LTE Cell」页「*DlBandwidth」所在列
    constBandwidthCol = getColNum(sh.name, constRecordRow, "DlBandWidth", "Cell")
    
    '「LTE Cell」页「SubframeAssignment」所在列
    constSACol = getColNum(sh.name, constRecordRow, "SubframeAssignment", "Cell")
    
    '「LTE Cell」页「*FddTddInd」所在列
    constFDDTDDCol = getColNum(sh.name, constRecordRow, "FddTddInd", "Cell")
    
    '「LTE Cell」页「*TxRxMode」所在列
    constTxRxModeCol = getColNum(sh.name, constRecordRow, "TxRxMode", "Cell")
    
    '「LTE Cell」页「*NB-IoT TA Flag」所在列
    If getNBIOTFlag = True Then
        constNBIOTFlagCol = getColNum(sh.name, constRecordRow, "NbCellFlag", "Cell")
    Else
        constNBIOTFlagCol = -1
    End If
      
    '「LTE Cell」页「*Cell Template」所在列
    constCellTempCol = getColNum(sh.name, constRecordRow, "CellTemplateName", "Cell")
    Dim bandWidthValue As String
    Dim saValue As String
    Dim fddTddValue As String
    Dim txRxModeValue As String
    Dim nbIotValue As String
    If constBandwidthCol = -1 Then
         bandWidthValue = ""
     Else
         bandWidthValue = Cells(target.row, constBandwidthCol).value
     End If
     
     If constTxRxModeCol = -1 Then
         txRxModeValue = ""
     Else
         txRxModeValue = Cells(target.row, constTxRxModeCol).value
     End If
     
     If constFDDTDDCol = -1 Then
         fddTddValue = ""
     Else
         fddTddValue = Cells(target.row, constFDDTDDCol).value
     End If
     
     If constSACol = -1 Then
         saValue = ""
     Else
         saValue = Cells(target.row, constSACol).value
     End If
     
    If constNBIOTFlagCol = -1 Then
        nbIotValue = "False"
    Else
        nbIotValue = Cells(target.row, constNBIOTFlagCol).value
    End If
     

    If target.column = constCellTempCol And target.count = 1 And target.row > constRecordRow Then
         '获取「CellTemplate」列侯选值
        m_Cell_Template = Get_Template_Related(bandWidthValue, txRxModeValue, fddTddValue, saValue, nbIotValue, sh, target, getResByKey(LCellType))
        If m_Cell_Template <> "" Then
            With target.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=m_Cell_Template
            End With
            If Not target.Validation.value Then
                target.value = ""
            End If
        Else
            With target.Validation
                .Delete
                .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
                .IgnoreBlank = True
                .InCellDropdown = True
                .inputTitle = ""
                .ErrorTitle = ""
                .inputMessage = ""
                .ErrorMessage = ""
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
        End If
    End If
End Sub

'从「MappingCellTemplate」页获取「Cell Template」列侯选值
Function Get_Template_Related(DlBandwidth As String, TxRxMode As String, FDDTDD As String, SA As String, NBIoTFlag As String, sheet As Worksheet, cellRange As range, ByRef CellType As String) As String
    Dim m_rowNum As Long
    Dim m_Str As String
    Dim m_start As Long
    Dim m_end As Long
    
    Dim DlBandwidth1 As String, FDDTDD1 As String
    m_start = 0
    
    m_Str = ""
    Select Case DlBandwidth
        Case "CELL_BW_N6"
            DlBandwidth1 = "1.4M"
        Case "CELL_BW_N15"
            DlBandwidth1 = "3M"
        Case "CELL_BW_N25"
            DlBandwidth1 = "5M"
        Case "CELL_BW_N50"
            DlBandwidth1 = "10M"
        Case "CELL_BW_N75"
            DlBandwidth1 = "15M"
        Case "CELL_BW_N100"
            DlBandwidth1 = "20M"
        Case Else
            DlBandwidth1 = ""
    End Select
       
    Select Case FDDTDD
        Case "CELL_TDD"
            FDDTDD1 = "TDD"
        Case "CELL_FDD"
            FDDTDD1 = "FDD"
        Case Else
            FDDTDD1 = ""
    End Select

    m_Str = ""
    
    '这个新增的容器用来去重复用的，控制器和基站小区模板解耦可能会导致小区模板页有重复模板
    Dim cellTemplateCol As New Collection
    Dim cellTemplate As String
    
    Dim BandwidthCol As Long, TxRxModeCol As Long, FDDTDDCol As Long, SACol As Long, CellPatternCol As Long, CellTypeCol As Long, NETypeCol As Long
    Dim MappingCellTemplate As Worksheet
    Set MappingCellTemplate = ThisWorkbook.Worksheets("MappingCellTemplate")

    Call getCellTemplateCol(BandwidthCol, TxRxModeCol, FDDTDDCol, SACol, CellPatternCol, CellTypeCol, NETypeCol)
    
    For m_rowNum = 2 To MappingCellTemplate.range(getColStr(CellPatternCol) & "1048576").End(xlUp).row
        If (DlBandwidth1 = MappingCellTemplate.Cells(m_rowNum, BandwidthCol).value Or DlBandwidth1 = "" Or MappingCellTemplate.Cells(m_rowNum, BandwidthCol).value = "") _
        And (TxRxMode = MappingCellTemplate.Cells(m_rowNum, TxRxModeCol).value Or TxRxMode = "" Or MappingCellTemplate.Cells(m_rowNum, TxRxModeCol).value = "") _
        And (FDDTDD1 = MappingCellTemplate.Cells(m_rowNum, FDDTDDCol).value Or FDDTDD1 = "" Or MappingCellTemplate.Cells(m_rowNum, FDDTDDCol).value = "") _
        And CellType = MappingCellTemplate.Cells(m_rowNum, CellTypeCol) _
        And (SA = MappingCellTemplate.Cells(m_rowNum, SACol).value Or SA = "" Or MappingCellTemplate.Cells(m_rowNum, SACol).value = "") And (UCase(NBIoTFlag) = UCase("FALSE") Or NBIoTFlag = "") _
        Or ((UCase(NBIoTFlag) = UCase("TRUE") And UCase(MappingCellTemplate.Cells(m_rowNum, FDDTDDCol).value) = UCase("NB-IoT"))) Then
            cellTemplate = MappingCellTemplate.Cells(m_rowNum, CellPatternCol).value
            If Contains(cellTemplateCol, cellTemplate) Or cellTemplate = "" Then GoTo NextLoop
            cellTemplateCol.Add Item:=cellTemplate, key:=cellTemplate
            If m_Str = "" Then
                 m_Str = cellTemplate
            Else
                 m_Str = m_Str & "," & cellTemplate
            End If
         End If
NextLoop:
    Next
    
     If Len(m_Str) > 255 Then
        Dim groupName As String
        Dim columnName As String
        Dim valideDef As CValideDef
        Call getGroupAndColumnName(sheet, cellRange, groupName, columnName)
        Set valideDef = initDefaultDataSub.getInnerValideDef(sheet.name + "," + groupName + "," + columnName)
        If valideDef Is Nothing Then
            Set valideDef = addInnerValideDef(sheet.name, groupName, columnName, m_Str)
        Else
            Call modiflyInnerValideDef(sheet.name, groupName, columnName, m_Str, valideDef)
        End If
        m_Str = valideDef.getValidedef
    End If
    Get_Template_Related = m_Str
    
End Function



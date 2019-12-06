Attribute VB_Name = "CellSub"
'「eNodeB Radio Data」页记录起始行
Private Const constRecordRow = 2

Function isCellSheet(sheetName As String) As Boolean
    If sheetName = getResByKey("LTE Cell") Or sheetName = getResByKey("Cell Sector Equipment") Or sheetName = getResByKey("PRB Sector Equipment") Then
        isCellSheet = True
        Exit Function
    End If
    isCellSheet = False
End Function

Function isTrasnPortSheet(sheetName As String) As Boolean
    If sheetName = getResByKey("LTE Cell") Or sheetName = getResByKey("Cell Sector Equipment") Or sheetName = getResByKey("PRB Sector Equipment") Then
        isTrasnPortSheet = True
        Exit Function
    End If
    isTrasnPortSheet = False
End Function

'定义设置「Cell Template」列下拉列表的事件
Sub Cell_Worksheet_SelectionChange(ByVal sh As Worksheet, ByVal target As range)
    Dim m_Cell_Template As String
    
    If getNBIOTFlag = True Then
        constNBIOTFlag = getColNum(sh.name, constRecordRow, "NbCellFlag", "Cell")
    Else
        constNBIOTFlag = -1
    End If
    
    '「LTE Cell」页「*DlBandwidth」所在列
    constBandwidthCol = getColNum(sh.name, constRecordRow, "DlBandWidth", "Cell")
    
    '「LTE Cell」页「SubframeAssignment」所在列
    constSACol = getColNum(sh.name, constRecordRow, "SubframeAssignment", "Cell")
    
    '「LTE Cell」页「*FddTddInd」所在列
    constFddTddCol = getColNum(sh.name, constRecordRow, "FddTddInd", "Cell")
    
    '「LTE Cell」页「*TxRxMode」所在列
    constTxRxModeCol = getColNum(sh.name, constRecordRow, "TxRxMode", "Cell")
      
    '「LTE Cell」页「*Cell Template」所在列
    constCellTempCol = getColNum(sh.name, constRecordRow, "CellTemplateName", "Cell")
    Dim bandWidthValue As String
    Dim saValue As String
    Dim fddTddValue As String
    Dim txRxModeValue As String
    Dim nbiotValue As String
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
     
     If constFddTddCol = -1 Then
         fddTddValue = ""
     Else
         fddTddValue = Cells(target.row, constFddTddCol).value
     End If
     
     If constSACol = -1 Then
         saValue = ""
     Else
         saValue = Cells(target.row, constSACol).value
     End If
     
    If constNBIOTFlag = -1 Then
         nbiotValue = "false"
     Else
         nbiotValue = Cells(target.row, constNBIOTFlag).value
     End If
     

    If target.column = constCellTempCol And target.count = 1 And target.row > constRecordRow Then
         '获取「CellTemplate」列侯选值
        m_Cell_Template = Get_Template_Related(bandWidthValue, txRxModeValue, fddTddValue, saValue, nbiotValue, sh, target)
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
                .inputmessage = ""
                .ErrorMessage = ""
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
            target.value = ""
        End If
    End If
End Sub
'从「MappingCellTemplate」页获取「Cell Template」列侯选值
Function Get_Template_Related(DlBandwidth As String, TxRxMode As String, fddTdd As String, SA As String, NBIoTFlag As String, sheet As Worksheet, cellRange As range) As String
    Dim m_rowNum As Long
    Dim m_Str As String
    Dim m_start As Long
    Dim m_end As Long
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
       
    Select Case fddTdd
        Case "CELL_TDD"
            FDDTDD1 = "TDD"
        Case "CELL_FDD"
            FDDTDD1 = "FDD"
        Case "CELL_NB-IoT"
            FDDTDD1 = "NB-IoT"
        Case Else
            FDDTDD1 = ""
    End Select

    m_Str = ""
    
    '这个新增的容器用来去重复用的，控制器和基站小区模板解耦可能会导致小区模板页有重复模板
    Dim cellTemplateCol As New Collection
    Dim cellTemplate As String
    
    Dim BandwidthCol As Long, TxRxModeCol As Long, fddTddCol As Long, SACol As Long, CellPatternCol As Long, CellTypeCol As Long, NETypeCol As Long
    
    BandwidthCol = attrNameColNumInSpecialDef(Worksheets("MappingCellTemplate"), "Bandwidth")
    TxRxModeCol = attrNameColNumInSpecialDef(Worksheets("MappingCellTemplate"), "TxRxMode")
    fddTddCol = attrNameColNumInSpecialDef(Worksheets("MappingCellTemplate"), "FDD/TDD")
    SACol = attrNameColNumInSpecialDef(Worksheets("MappingCellTemplate"), "SA")
    CellPatternCol = attrNameColNumInSpecialDef(Worksheets("MappingCellTemplate"), "CellPattern")
    CellTypeCol = attrNameColNumInSpecialDef(Worksheets("MappingCellTemplate"), "CellType")
    NETypeCol = attrNameColNumInSpecialDef(Worksheets("MappingCellTemplate"), "NEType")
    
    
    For m_rowNum = 2 To getUsedRowCount(Worksheets("MappingCellTemplate"))
        If (DlBandwidth1 = Worksheets("MappingCellTemplate").Cells(m_rowNum, BandwidthCol).value Or DlBandwidth1 = "" Or Worksheets("MappingCellTemplate").Cells(m_rowNum, BandwidthCol).value = "") _
        And (FDDTDD1 = Worksheets("MappingCellTemplate").Cells(m_rowNum, fddTddCol).value Or FDDTDD1 = "" Or Worksheets("MappingCellTemplate").Cells(m_rowNum, fddTddCol).value = "") _
        And (SA = Worksheets("MappingCellTemplate").Cells(m_rowNum, SACol).value Or SA = "") _
        And (CellTypeCol <= 0 Or getResByKey("LTE Cell") = Worksheets("MappingCellTemplate").Cells(m_rowNum, CellTypeCol).value Or Len(Worksheets("MappingCellTemplate").Cells(m_rowNum, CellTypeCol).value) = 0) _
        And (UCase(NBIoTFlag) = UCase("FALSE") Or NBIoTFlag = "") _
        Or ((UCase(NBIoTFlag) = UCase("TRUE") And UCase(Worksheets("MappingCellTemplate").Cells(m_rowNum, fddTddCol).value) = UCase("NB-IoT"))) Then
            cellTemplate = Worksheets("MappingCellTemplate").Cells(m_rowNum, CellPatternCol).value
            If Contains(cellTemplateCol, cellTemplate) Or cellTemplate = "" Then GoTo NextLoop
            cellTemplateCol.Add Item:=cellTemplate, key:=cellTemplate
            If m_Str = "" Then
                 m_Str = cellTemplate
            ElseIf VBA.Trim(cellTemplate) <> "" Then
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









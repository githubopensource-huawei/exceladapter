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
        Or sheetName = getResByKey("GSM Logic Cell") Or sheetName = getResByKey("UMTS Logic Cell") Or sheetName = getResByKey("LTE Cell") Then
        isCellSheet = True
        Exit Function
    End If
    isCellSheet = False
End Function
 
Function isTrasnPortSheet(sheetName As String) As Boolean
    If sheetName = "GSM Cell" Or sheetName = "UMTS Cell" Or sheetName = "LTE Cell" _
        Or sheetName = getResByKey("GSM Logic Cell") Or sheetName = getResByKey("UMTS Logic Cell") Or getResByKey("LTE Cell") _
        Or sheetName = "GTRXGROUP" Or sheetName = getResByKey("GTRXGROUP") Then
        isTrasnPortSheet = True
        Exit Function
    End If
    isTrasnPortSheet = False
End Function

'需要区分处理LTE和Mrat的差异
Public Sub cellSheetSelectionEvent(ByVal sheet As Worksheet, ByVal Target As Range)
    Dim neType As String
    neType = getNeType()
    If neType = "LTE" Then
        Call Cell_Worksheet_SelectionChange(sheet, Target)
    Else 'Mrat,UMTS
        Call CellSelectionChange(sheet, Target)
    End If
End Sub

'定义设置「Cell Template」列下拉列表的事件
Public Sub CellSelectionChange(ByVal sheet As Worksheet, ByVal Target As Range)
        'init 物理GCell
        Call initCellTemplate(sheet, Target, attrName, cellMocName, getResByKey(CellType))
        'init 逻辑GCell
        Call initCellTemplate(sheet, Target, logicAttrName, logicCellMocName, getResByKey(logicCellType))
        
        'init 物理UCell
        Call initCellTemplate(sheet, Target, UAttrName, UCellMocName, getResByKey(UCellType))
        'init 逻辑UCell
        Call initCellTemplate(sheet, Target, logicUAttrName, logicUCellMocName, getResByKey(logicUCellType))
        
        'init 物理LCell
        Call initCellTemplate(sheet, Target, LAttrName, LCellMocName, getResByKey(LCellType))
End Sub

Sub initCellTemplate(ByVal sheet As Worksheet, ByVal Target As Range, myAttrName As String, myCellMocName As String, myType As String)
        Dim m_Cell_Template As String
        Dim constCellTempCol As Long
        '「物理Cell Template」所在列
        constCellTempCol = getColNum(sheet.name, constRecordRow, myAttrName, myCellMocName)
        
        If constCellTempCol >= 0 And Target.column = constCellTempCol And Target.count = 1 And Target.row > constRecordRow Then
            '获取「CellTemplate」列侯选值
            m_Cell_Template = getCellTemplate(myType, sheet, Target)
            If m_Cell_Template <> "" Then
                With Target.Validation
                   .Delete
                   .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=m_Cell_Template
                End With
                If Not Target.Validation.value Then
                    Target.value = ""
                End If
            Else
                With Target.Validation
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
                    Target.value = ""
            End If
        End If
End Sub

'从「MappingCellTemplate」页获取「Cell Template」列侯选值
Function getCellTemplate(myType As String, sheet As Worksheet, cellRange As Range) As String
    Dim m_rowNum As Long
    Dim m_RowNum_Inner As Long
    Dim flag As Boolean
    Dim m_Str As String
    Dim m_start As Long
    Dim m_end As Long
    Dim neType As String
    
    Dim MappingCellTemplate As Worksheet
    Set MappingCellTemplate = ThisWorkbook.Worksheets("MappingCellTemplate")
    
    neType = getNeType()
    m_Str = ""
      For m_rowNum = 2 To MappingCellTemplate.Range("a1048576").End(xlUp).row
        If myType = MappingCellTemplate.Cells(m_rowNum, 2).value _
        And neType = MappingCellTemplate.Cells(m_rowNum, 3).value Then
                   If m_Str = "" Then
                        m_Str = MappingCellTemplate.Cells(m_rowNum, 1).value
                   Else
                        m_Str = m_Str & "," & MappingCellTemplate.Cells(m_rowNum, 1).value
                   End If
         End If
    Next
    
    If Len(m_Str) > 256 Then
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
Sub Cell_Worksheet_SelectionChange(ByVal sh As Worksheet, ByVal Target As Range)
    Dim m_Cell_Template As String
    
    Dim constBandwidthCol As Long, constSACol As Long, constFDDTDDCol As Long, constTxRxModeCol As Long, constCellTempCol As Long
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
    Dim bandWidthValue As String
    Dim saValue As String
    Dim fddTddValue As String
    Dim txRxModeValue As String
    If constBandwidthCol = -1 Then
         bandWidthValue = ""
     Else
         bandWidthValue = Cells(Target.row, constBandwidthCol).value
     End If
     
     If constTxRxModeCol = -1 Then
         txRxModeValue = ""
     Else
         txRxModeValue = Cells(Target.row, constTxRxModeCol).value
     End If
     
     If constFDDTDDCol = -1 Then
         fddTddValue = ""
     Else
         fddTddValue = Cells(Target.row, constFDDTDDCol).value
     End If
     
     If constSACol = -1 Then
         saValue = ""
     Else
         saValue = Cells(Target.row, constSACol).value
     End If
     

    If Target.column = constCellTempCol And Target.count = 1 And Target.row > constRecordRow Then
         '获取「CellTemplate」列侯选值
        m_Cell_Template = Get_Template_Related(bandWidthValue, txRxModeValue, fddTddValue, saValue, sh, Target)
        If m_Cell_Template <> "" Then
            With Target.Validation
               .Delete
               .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=m_Cell_Template
            End With
        Else
            With Target.Validation
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
Function Get_Template_Related(DlBandwidth As String, TxRxMode As String, FDDTDD As String, SA As String, sheet As Worksheet, cellRange As Range) As String
    Dim m_rowNum As Long
    Dim m_Str As String
    Dim m_start As Long
    Dim m_end As Long
    
    Dim MappingCellTemplate As Worksheet
    Set MappingCellTemplate = ThisWorkbook.Worksheets("MappingCellTemplate")
    
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
    For m_rowNum = 2 To MappingCellTemplate.Range("e1048576").End(xlUp).row
        If (DlBandwidth1 = MappingCellTemplate.Cells(m_rowNum, 1).value Or DlBandwidth1 = "" Or MappingCellTemplate.Cells(m_rowNum, 1).value = "") _
        And (TxRxMode = MappingCellTemplate.Cells(m_rowNum, 2).value Or TxRxMode = "" Or MappingCellTemplate.Cells(m_rowNum, 2).value = "") _
        And (FDDTDD1 = MappingCellTemplate.Cells(m_rowNum, 3).value Or FDDTDD1 = "" Or MappingCellTemplate.Cells(m_rowNum, 3).value = "") _
        And (SA = MappingCellTemplate.Cells(m_rowNum, 4).value Or SA = "") Then
                   If m_Str = "" Then
                        m_Str = MappingCellTemplate.Cells(m_rowNum, 5).value
                   Else
                        m_Str = m_Str & "," & MappingCellTemplate.Cells(m_rowNum, 5).value
                   End If
         End If
    Next
    
     If Len(m_Str) > 256 Then
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




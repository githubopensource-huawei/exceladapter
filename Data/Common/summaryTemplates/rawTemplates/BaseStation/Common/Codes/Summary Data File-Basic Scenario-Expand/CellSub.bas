Attribute VB_Name = "CellSub"
Option Explicit

'��eNodeB Radio Data��ҳ��¼��ʼ��
Private Const constRecordRow = 2
Private Const cellMocName As String = "GLoCell"
Private Const attrName As String = "CellTemplateName"
Private Const cellType As String = "GSM Local Cell"

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


Private Const MCellMocName As String = "MCell"
Private Const MAttrName As String = "CellTemplateName"
Private Const MCellType As String = "NB-IoT Cell"

Private Const RCellMocName As String = "RFALoCell"
Private Const RAttrName As String = "CellTemplateName"
Private Const RCellType As String = "RFA Cell"

Private Const NRLocalCellMocName As String = "NRLoCell"
Private Const NRLocalCellAttrName As String = "CellTemplateName"
Private Const NRLocalCellCellType As String = "NR Local Cell"

Private Const NRDuCellMocName As String = "NRDUCell"
Private Const NRDuCellAttrName As String = "CellTemplateName"
Private Const NRDuCellCellType As String = "NR DU Cell"


Private Const NRCellMocName As String = "NRCell"
Private Const NRAttrName As String = "CellTemplateName"
Private Const NRCellType As String = "NR Cell"


Function isCellExist() As Boolean
    If IsSheetExist(getResByKey("GSMCell")) Or IsSheetExist(getResByKey("UMTSCell")) _
    Or IsSheetExist(getResByKey("LTECell")) Or IsSheetExist(getResByKey("RFA Cell")) _
    Or IsSheetExist(getResByKey("NB-IoTCell")) Or IsSheetExist(getResByKey("NR Cell")) _
    Or IsSheetExist(getResByKey("NR Local Cell")) Or IsSheetExist(getResByKey("NR DU Cell")) Then
        isCellExist = True
    Else
         isCellExist = False
    End If
End Function
Function isCellSheet(sheetName As String) As Boolean
    If sheetName = getResByKey("GSMCell") Or sheetName = getResByKey("UMTSCell") _
    Or sheetName = getResByKey("LTECell") Or sheetName = getResByKey("RFA Cell") _
    Or sheetName = getResByKey("NB-IoTCell") Or sheetName = getResByKey("NR Cell") _
    Or sheetName = getResByKey("NR Local Cell") Or sheetName = getResByKey("NR DU Cell") Then
        isCellSheet = True
        Exit Function
    End If
    isCellSheet = False
End Function
 
Function isTrasnPortSheet(sheetName As String) As Boolean
    If sheetName = getResByKey("GSMCell") Or sheetName = getResByKey("UMTSCell") _
    Or sheetName = getResByKey("LTECell") Or sheetName = getResByKey("RFA Cell") _
    Or sheetName = getResByKey("NB-IoTCell") _
    Or sheetName = getResByKey("GTRXGROUP") Or sheetName = getResByKey("GTRX") _
    Or sheetName = getResByKey("NB-IoT TRX") Or sheetName = getResByKey("NR Cell") _
    Or sheetName = getResByKey("NRDUCellTrp") Or sheetName = getResByKey("NR DU Cell") Or sheetName = getResByKey("NR Local Cell") _
    Or sheetName = getResByKey("Cell Sector Equipment") Or sheetName = getResByKey("PRB Sector Equipment") Then
        isTrasnPortSheet = True
        Exit Function
    End If
    isTrasnPortSheet = False
End Function

Function isGsmCellSheet(sheetName As String) As Boolean
    If sheetName = getResByKey("GSMCell") Then
        isGsmCellSheet = True
        Exit Function
    End If
    isGsmCellSheet = False
End Function

Function isLteCellSheet(ByRef sheetName As String) As Boolean
    If sheetName = getResByKey("LTECell") Then
        isLteCellSheet = True
    Else
        isLteCellSheet = False
    End If
End Function

'�������á�Cell Template���������б���¼�
Public Sub CellSelectionChange(ByVal sheet As Worksheet, ByVal target As range)
        'init ����GCell
        Call initCellTemplate(sheet, target, attrName, cellMocName, getResByKey(cellType))
        'init �߼�GCell
        Call initCellTemplate(sheet, target, logicAttrName, logicCellMocName, getResByKey(logicCellType))
        
        'init ����UCell
        Call initCellTemplate(sheet, target, UAttrName, UCellMocName, getResByKey(UCellType))
        'init �߼�UCell
        Call initCellTemplate(sheet, target, logicUAttrName, logicUCellMocName, getResByKey(logicUCellType))
        
        'init ����LCell
        Call initCellTemplate(sheet, target, LAttrName, LCellMocName, getResByKey(LCellType))
         'init ����MCell
        Call initCellTemplate(sheet, target, MAttrName, MCellMocName, getResByKey(MCellType))
        
         'init ����RCell
        Call initCellTemplate(sheet, target, RAttrName, RCellMocName, getResByKey(RCellType))
        
        Call initCellTemplate(sheet, target, NRLocalCellAttrName, NRLocalCellMocName, getResByKey(NRLocalCellCellType))
        
        Call initCellTemplate(sheet, target, NRDuCellAttrName, NRDuCellMocName, getResByKey(NRDuCellCellType))
        
        Call initCellTemplate(sheet, target, NRAttrName, NRCellMocName, getResByKey(NRCellType))
        
End Sub

Sub initCellTemplate(ByVal sheet As Worksheet, ByVal target As range, myAttrName As String, myCellMocName As String, cellType As String)
        '�����LTEС���������LTEС�����еİ���������ɸѡ������С��ҳ��ԭ������
        If cellType = getResByKey(LCellType) Then
            Call initLteCellTemplate(sheet, target, cellType)
            Exit Sub
        ElseIf cellType = getResByKey(NRLocalCellCellType) Or cellType = getResByKey(NRCellType) Or cellType = getResByKey(NRDuCellCellType) Then
            Call initNrCellTemplate(sheet, target, myCellMocName, cellType)
            Exit Sub
        End If
        
        Dim m_Cell_Template As String
        
        Dim constCellTempCol As Long
        '������Cell Template��������
        constCellTempCol = getColNum(sheet.name, constRecordRow, myAttrName, myCellMocName)

        If constCellTempCol >= 0 And target.column = constCellTempCol And target.count = 1 And target.row > constRecordRow Then
            '��ȡ��CellTemplate���к�ѡֵ
            m_Cell_Template = getCellTemplate(cellType, sheet, target)
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
                    'Target.value = ""
            End If
        End If
End Sub


Sub initNrCellTemplate(ByVal sh As Worksheet, ByVal target As range, ByVal myCellMocName As String, ByVal cellType As String)
    Dim m_Cell_Template As String
    Dim constBandwidthCol As Long, constSACol As Long, constFDDTDDCol As Long, constTxRxModeCol As Long, constCellTempCol As Long, constNBIOTFlagCol As Long
    
    Dim cellMocName As String
    cellMocName = myCellMocName
    
   
    
    '��NR Cell��ҳ��*DlBandwidth��������
    constBandwidthCol = getColNum(sh.name, constRecordRow, "DlBandwidth", cellMocName)
    'constBandwidthCol = -1
    
    '��NR Cell��ҳ��SubframeAssignment��������
    'constSACol = getColNum(sh.name, constRecordRow, "SubframeAssignment", "Cell")
    constSACol = -1
    
    '��NR Cell��ҳ��*FddTddInd��������
    constFDDTDDCol = getColNum(sh.name, constRecordRow, "DuplexMode", cellMocName)
    
    '��NR Cell��ҳ��*TxRxMode��������
    'constTxRxModeCol = getColNum(sh.name, constRecordRow, "TxRxMode", "Cell")
    constTxRxModeCol = -1
      
    '��NR Cell��ҳ��*Cell Template��������
    constCellTempCol = getColNum(sh.name, constRecordRow, "CellTemplateName", cellMocName)
    
    '��NR Cell��ҳ��*NB-IoT TA Flag��������
'    If getNBIOTFlag = True Then
'        constNBIOTFlagCol = getColNum(sh.name, constRecordRow, "NbCellFlag", "Cell")
'    Else
'        constNBIOTFlagCol = -1
'    End If
    constNBIOTFlagCol = -1
    
    Dim bandwidthValue As String
    Dim saValue As String
    Dim fddtddValue As String
    Dim txrxModeValue As String
    Dim NBIOTCellFlag As String
    If constBandwidthCol = -1 Then
         bandwidthValue = ""
     Else
         bandwidthValue = Cells(target.row, constBandwidthCol).value
     End If
     
     If constTxRxModeCol = -1 Then
         txrxModeValue = ""
     Else
         txrxModeValue = Cells(target.row, constTxRxModeCol).value
     End If
     
     If constFDDTDDCol = -1 Then
         fddtddValue = ""
     Else
         fddtddValue = Cells(target.row, constFDDTDDCol).value
     End If
     
     If constSACol = -1 Then
         saValue = ""
     Else
         saValue = Cells(target.row, constSACol).value
     End If
     
     If constNBIOTFlagCol = -1 Then
         NBIOTCellFlag = "FALSE"
     Else
         NBIOTCellFlag = Cells(target.row, constNBIOTFlagCol).value
     End If
     

    If target.column = constCellTempCol And target.count = 1 And target.row > constRecordRow Then
         '��ȡ��CellTemplate���к�ѡֵ
        m_Cell_Template = Get_NRCellTemplate_Related(bandwidthValue, txrxModeValue, fddtddValue, saValue, NBIOTCellFlag, sh, target, cellType)
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

Sub initLteCellTemplate(ByVal sh As Worksheet, ByVal target As range, ByVal cellType As String)
    Dim m_Cell_Template As String
    Dim constBandwidthCol As Long, constSACol As Long, constFDDTDDCol As Long, constTxRxModeCol As Long, constCellTempCol As Long, constNBIOTFlagCol As Long
    
    '��LTE Cell��ҳ��*DlBandwidth��������
    constBandwidthCol = getColNum(sh.name, constRecordRow, "DlBandWidth", "Cell")
    
    '��LTE Cell��ҳ��SubframeAssignment��������
    constSACol = getColNum(sh.name, constRecordRow, "SubframeAssignment", "Cell")
    
    '��LTE Cell��ҳ��*FddTddInd��������
    constFDDTDDCol = getColNum(sh.name, constRecordRow, "FddTddInd", "Cell")
    
    '��LTE Cell��ҳ��*TxRxMode��������
    constTxRxModeCol = getColNum(sh.name, constRecordRow, "TxRxMode", "Cell")
      
    '��LTE Cell��ҳ��*Cell Template��������
    constCellTempCol = getColNum(sh.name, constRecordRow, "CellTemplateName", "Cell")
    
    '��LTE Cell��ҳ��*NB-IoT TA Flag��������
    If getNBIOTFlag = True Then
        constNBIOTFlagCol = getColNum(sh.name, constRecordRow, "NbCellFlag", "Cell")
    Else
        constNBIOTFlagCol = -1
    End If
    
    Dim bandwidthValue As String
    Dim saValue As String
    Dim fddtddValue As String
    Dim txrxModeValue As String
    Dim NBIOTCellFlag As String
    If constBandwidthCol = -1 Then
         bandwidthValue = ""
     Else
         bandwidthValue = Cells(target.row, constBandwidthCol).value
     End If
     
     If constTxRxModeCol = -1 Then
         txrxModeValue = ""
     Else
         txrxModeValue = Cells(target.row, constTxRxModeCol).value
     End If
     
     If constFDDTDDCol = -1 Then
         fddtddValue = ""
     Else
         fddtddValue = Cells(target.row, constFDDTDDCol).value
     End If
     
     If constSACol = -1 Then
         saValue = ""
     Else
         saValue = Cells(target.row, constSACol).value
     End If
     
     If constNBIOTFlagCol = -1 Then
         NBIOTCellFlag = "FALSE"
     Else
         NBIOTCellFlag = Cells(target.row, constNBIOTFlagCol).value
     End If
     

    If target.column = constCellTempCol And target.count = 1 And target.row > constRecordRow Then
         '��ȡ��CellTemplate���к�ѡֵ
        m_Cell_Template = Get_LteCellTemplate_Related(bandwidthValue, txrxModeValue, fddtddValue, saValue, NBIOTCellFlag, sh, target, cellType)
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

'�ӡ�MappingCellTemplate��ҳ��ȡ��Cell Template���к�ѡֵ
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
    
    '�����������������ȥ�ظ��õģ��������ͻ�վС��ģ�������ܻᵼ��С��ģ��ҳ���ظ�ģ��
    Dim cellTemplateCol As New Collection
    Dim cellTemplate As String
    
    For m_rowNum = 2 To Worksheets("MappingCellTemplate").range("a65536").End(xlUp).row
        cellTemplate = Worksheets("MappingCellTemplate").Cells(m_rowNum, 1).value
        If (myType = Worksheets("MappingCellTemplate").Cells(m_rowNum, 2).value Or Len(Trim(Worksheets("MappingCellTemplate").Cells(m_rowNum, 2).value)) = 0) _
        And neType = Worksheets("MappingCellTemplate").Cells(m_rowNum, 3).value _
        And (Not Contains(cellTemplateCol, cellTemplate)) And cellTemplate <> "" Then
            cellTemplateCol.Add Item:=cellTemplate, key:=cellTemplate
            If m_Str = "" Then
                 m_Str = cellTemplate
            ElseIf VBA.Trim(cellTemplate) <> "" Then
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




'�µĴӡ�MappingCellTemplate��ҳ��ȡ��Cell Template���к�ѡֵ��NR Cell������������
Function Get_NRCellTemplate_Related(DlBandwidth As String, txrxMode As String, FDDTDD As String, sa As String, NBIoTFlag As String, sheet As Worksheet, cellRange As range, ByRef cellType As String) As String
    Dim m_rowNum As Long
    Dim m_Str As String
    Dim m_start As Long
    Dim m_end As Long
    m_start = 0
    
    m_Str = ""
    
    Dim DlBandwidth1 As String
    Select Case DlBandwidth
        Case "CELL_BW_N10"
            DlBandwidth1 = "10M"
        Case "CELL_BW_N15"
            DlBandwidth1 = "15M"
        Case "CELL_BW_N20"
            DlBandwidth1 = "20M"
        Case "CELL_BW_N40"
            DlBandwidth1 = "40M"
        Case "CELL_BW_N60"
            DlBandwidth1 = "60M"
        Case "CELL_BW_N80"
            DlBandwidth1 = "80M"
        Case "CELL_BW_N100"
            DlBandwidth1 = "100M"
        Case "CELL_BW_N100"
            DlBandwidth1 = "200M"
        Case "CELL_BW_10M"
            DlBandwidth1 = "10M"
        Case "CELL_BW_15M"
            DlBandwidth1 = "15M"
        Case "CELL_BW_20M"
            DlBandwidth1 = "20M"
        Case "CELL_BW_40M"
            DlBandwidth1 = "40M"
        Case "CELL_BW_60M"
            DlBandwidth1 = "60M"
        Case "CELL_BW_80M"
            DlBandwidth1 = "80M"
        Case "CELL_BW_100M"
            DlBandwidth1 = "100M"
        Case "CELL_BW_100M"
            DlBandwidth1 = "200M"
        Case Else
            DlBandwidth1 = ""
    End Select
    
    Dim FDDTDD1 As String
    Select Case FDDTDD
        Case "CELL_TDD"
            FDDTDD1 = "TDD"
        Case "CELL_FDD"
            FDDTDD1 = "FDD"
        Case "CELL_SUL"
             FDDTDD1 = "SUL"
        Case Else
            FDDTDD1 = ""
    End Select

    Dim neType As String
    neType = getNeType()
    m_Str = ""
    
    '�����������������ȥ�ظ��õģ��������ͻ�վС��ģ�������ܻᵼ��С��ģ��ҳ���ظ�ģ��
    Dim cellTemplateCol As New Collection
    Dim cellTemplate As String
    
    For m_rowNum = 2 To Worksheets("MappingCellTemplate").range("A65536").End(xlUp).row
        If (DlBandwidth1 = Worksheets("MappingCellTemplate").Cells(m_rowNum, 4).value Or DlBandwidth1 = "" Or Worksheets("MappingCellTemplate").Cells(m_rowNum, 4).value = "") _
        And (FDDTDD1 = Worksheets("MappingCellTemplate").Cells(m_rowNum, 6).value Or FDDTDD1 = "" Or Worksheets("MappingCellTemplate").Cells(m_rowNum, 6).value = "") _
        And (sa = Worksheets("MappingCellTemplate").Cells(m_rowNum, 7).value Or sa = "" Or Worksheets("MappingCellTemplate").Cells(m_rowNum, 7).value = "") _
        And (cellType = Worksheets("MappingCellTemplate").Cells(m_rowNum, 2) Or Len(Trim(Worksheets("MappingCellTemplate").Cells(m_rowNum, 2).value)) = 0) _
        And neType = Worksheets("MappingCellTemplate").Cells(m_rowNum, 3) _
        And (UCase(NBIoTFlag) = UCase("FALSE") Or NBIoTFlag = "") _
        Or ((UCase(NBIoTFlag) = UCase("TRUE") And UCase(Worksheets("MappingCellTemplate").Cells(m_rowNum, 6).value) = UCase("NB-IoT"))) Then
            cellTemplate = Worksheets("MappingCellTemplate").Cells(m_rowNum, 1).value
            If Contains(cellTemplateCol, cellTemplate) Or cellTemplate = "" Then GoTo NextLoop
            cellTemplateCol.Add Item:=cellTemplate, key:=cellTemplate
            If m_Str = "" Then
                 m_Str = Worksheets("MappingCellTemplate").Cells(m_rowNum, 1).value
            ElseIf VBA.Trim(Worksheets("MappingCellTemplate").Cells(m_rowNum, 1).value) <> "" Then
                 m_Str = m_Str & "," & Worksheets("MappingCellTemplate").Cells(m_rowNum, 1).value
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
    Get_NRCellTemplate_Related = m_Str
    
End Function


'�µĴӡ�MappingCellTemplate��ҳ��ȡ��Cell Template���к�ѡֵ��LTE Cell������������
Function Get_LteCellTemplate_Related(DlBandwidth As String, txrxMode As String, FDDTDD As String, sa As String, NBIoTFlag As String, sheet As Worksheet, cellRange As range, ByRef cellType As String) As String
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
    
    '�����������������ȥ�ظ��õģ��������ͻ�վС��ģ�������ܻᵼ��С��ģ��ҳ���ظ�ģ��
    Dim cellTemplateCol As New Collection
    Dim cellTemplate As String
    
    Dim BandwidthCol As Long, TxRxModeCol As Long, FDDTDDCol As Long, SACol As Long, CellPatternCol As Long, CellTypeCol As Long, NETypeCol As Long
    BandwidthCol = attrNameColNumInSpecialDef(Worksheets("MappingCellTemplate"), "Bandwidth")
    TxRxModeCol = attrNameColNumInSpecialDef(Worksheets("MappingCellTemplate"), "TxRxMode")
    FDDTDDCol = attrNameColNumInSpecialDef(Worksheets("MappingCellTemplate"), "FDD/TDD")
    SACol = attrNameColNumInSpecialDef(Worksheets("MappingCellTemplate"), "SA")
    CellPatternCol = attrNameColNumInSpecialDef(Worksheets("MappingCellTemplate"), "CellPattern")
    CellTypeCol = attrNameColNumInSpecialDef(Worksheets("MappingCellTemplate"), "CellType")
    NETypeCol = attrNameColNumInSpecialDef(Worksheets("MappingCellTemplate"), "NEType")
    
    
    For m_rowNum = 2 To getUsedRowCount(Worksheets("MappingCellTemplate"))
        If (DlBandwidth1 = Worksheets("MappingCellTemplate").Cells(m_rowNum, BandwidthCol).value Or DlBandwidth1 = "" Or Worksheets("MappingCellTemplate").Cells(m_rowNum, BandwidthCol).value = "") _
        And (FDDTDD1 = Worksheets("MappingCellTemplate").Cells(m_rowNum, FDDTDDCol).value Or FDDTDD1 = "" Or Worksheets("MappingCellTemplate").Cells(m_rowNum, FDDTDDCol).value = "") _
        And (sa = Worksheets("MappingCellTemplate").Cells(m_rowNum, SACol).value Or sa = "" Or Worksheets("MappingCellTemplate").Cells(m_rowNum, SACol).value = "") _
        And (cellType = Worksheets("MappingCellTemplate").Cells(m_rowNum, CellTypeCol) Or Len(Trim(Worksheets("MappingCellTemplate").Cells(m_rowNum, CellTypeCol).value)) = 0) _
        And neType = Worksheets("MappingCellTemplate").Cells(m_rowNum, NETypeCol) _
        And (UCase(NBIoTFlag) = UCase("FALSE") Or NBIoTFlag = "") _
        Or ((UCase(NBIoTFlag) = UCase("TRUE") And UCase(Worksheets("MappingCellTemplate").Cells(m_rowNum, FDDTDDCol).value) = UCase("NB-IoT"))) Then
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
    Get_LteCellTemplate_Related = m_Str
    
End Function










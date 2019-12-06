Attribute VB_Name = "CellSub"
Option Explicit

'��eNodeB Radio Data��ҳ��¼��ʼ��
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
        Or sheetName = getResByKey("A91") Or sheetName = getResByKey("A92") Or sheetName = getResByKey("A93") Then
        isCellSheet = True
        Exit Function
    End If
    isCellSheet = False
End Function
 
Function isTrasnPortSheet(sheetName As String) As Boolean
    If sheetName = "GSM Cell" Or sheetName = "UMTS Cell" Or sheetName = "LTE Cell" _
        Or sheetName = getResByKey("A78") Or sheetName = getResByKey("A79") Or sheetName = getResByKey("A80") _
        Or sheetName = "GTRXGROUP" Or sheetName = getResByKey("A81") Then
        isTrasnPortSheet = True
        Exit Function
    End If
    isTrasnPortSheet = False
End Function

Function isGsmCellSheet(sheetName As String) As Boolean
    If sheetName = "GSM Cell" Or sheetName = getResByKey("A82") Then
        isGsmCellSheet = True
        Exit Function
    End If
    isGsmCellSheet = False
End Function

Function isLteCellSheet(ByRef sheetName As String) As Boolean
    If sheetName = "LTE Cell" Or sheetName = getResByKey("A83") Then
        isLteCellSheet = True
    Else
        isLteCellSheet = False
    End If
End Function

'�������á�Cell Template���������б���¼�
Public Sub CellSelectionChange(ByVal sheet As Worksheet, ByVal Target As range)
        'init ����GCell
        Call initCellTemplate(sheet, Target, attrName, cellMocName, getResByKey(CellType))
        'init �߼�GCell
        Call initCellTemplate(sheet, Target, logicAttrName, logicCellMocName, getResByKey(logicCellType))
        
        'init ����UCell
        Call initCellTemplate(sheet, Target, UAttrName, UCellMocName, getResByKey(UCellType))
        'init �߼�UCell
        Call initCellTemplate(sheet, Target, logicUAttrName, logicUCellMocName, getResByKey(logicUCellType))
        
        'init ����LCell
        Call initCellTemplate(sheet, Target, LAttrName, LCellMocName, getResByKey(LCellType))
End Sub

Sub initCellTemplate(ByVal sheet As Worksheet, ByVal Target As range, myAttrName As String, myCellMocName As String, CellType As String)
        '�����LTEС���������LTEС�����еİ���������ɸѡ������С��ҳ��ԭ������
        If CellType = getResByKey(LCellType) Then
            Call initLteCellTemplate(sheet, Target, CellType)
            Exit Sub
        End If
        
        Dim m_Cell_Template As String
        
        Dim constCellTempCol As Long
        '������Cell Template��������
        constCellTempCol = getColNum(sheet.name, constRecordRow, myAttrName, myCellMocName)

        If constCellTempCol >= 0 And Target.column = constCellTempCol And Target.count = 1 And Target.row > constRecordRow Then
            '��ȡ��CellTemplate���к�ѡֵ
            m_Cell_Template = getCellTemplate(CellType, sheet, Target)
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
                    'Target.value = ""
            End If
        End If
End Sub

Sub initLteCellTemplate(ByVal sh As Worksheet, ByVal Target As range, ByVal CellType As String)
    Dim m_Cell_Template As String
    Dim constBandwidthCol As Long, constSACol As Long, constFDDTDDCol As Long, constTxRxModeCol As Long, constCellTempCol As Long
    
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
         '��ȡ��CellTemplate���к�ѡֵ
        m_Cell_Template = Get_LteCellTemplate_Related(bandWidthValue, txRxModeValue, fddTddValue, saValue, sh, Target, CellType)
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

'�ӡ�MappingCellTemplate��ҳ��ȡ��Cell Template���к�ѡֵ
Function getCellTemplate(myType As String, sheet As Worksheet, cellRange As range) As String
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
    
    '�����������������ȥ�ظ��õģ��������ͻ�վС��ģ�������ܻᵼ��С��ģ��ҳ���ظ�ģ��
    Dim cellTemplateCol As New Collection
    Dim cellTemplate As String
    
    For m_rowNum = 2 To MappingCellTemplate.range("a1048576").End(xlUp).row
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

'�µĴӡ�MappingCellTemplate��ҳ��ȡ��Cell Template���к�ѡֵ��LTE Cell������������
Function Get_LteCellTemplate_Related(DlBandwidth As String, TxRxMode As String, FDDTDD As String, SA As String, sheet As Worksheet, cellRange As range, ByRef CellType As String) As String
    Dim m_rowNum As Long
    Dim m_Str As String
    Dim m_start As Long
    Dim m_end As Long
    
    Dim MappingCellTemplate As Worksheet
    Set MappingCellTemplate = ThisWorkbook.Worksheets("MappingCellTemplate")
    
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
    
    For m_rowNum = 2 To MappingCellTemplate.range("A1048576").End(xlUp).row
        If (DlBandwidth1 = MappingCellTemplate.Cells(m_rowNum, 4).value Or DlBandwidth1 = "" Or MappingCellTemplate.Cells(m_rowNum, 4).value = "") _
        And (TxRxMode = MappingCellTemplate.Cells(m_rowNum, 5).value Or TxRxMode = "" Or MappingCellTemplate.Cells(m_rowNum, 5).value = "") _
        And (FDDTDD1 = MappingCellTemplate.Cells(m_rowNum, 6).value Or FDDTDD1 = "" Or MappingCellTemplate.Cells(m_rowNum, 6).value = "") _
        And (SA = MappingCellTemplate.Cells(m_rowNum, 7).value Or SA = "" Or MappingCellTemplate.Cells(m_rowNum, 7).value = "") _
        And CellType = MappingCellTemplate.Cells(m_rowNum, 2) _
        And neType = MappingCellTemplate.Cells(m_rowNum, 3) Then
            cellTemplate = MappingCellTemplate.Cells(m_rowNum, 1).value
            If Contains(cellTemplateCol, cellTemplate) Or cellTemplate = "" Then GoTo NextLoop
            cellTemplateCol.Add Item:=cellTemplate, key:=cellTemplate
            If m_Str = "" Then
                 m_Str = MappingCellTemplate.Cells(m_rowNum, 1).value
            Else
                 m_Str = m_Str & "," & MappingCellTemplate.Cells(m_rowNum, 1).value
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












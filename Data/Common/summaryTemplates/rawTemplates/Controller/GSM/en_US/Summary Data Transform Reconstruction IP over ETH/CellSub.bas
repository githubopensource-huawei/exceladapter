Attribute VB_Name = "CellSub"
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
    If sheetName = getResByKey("GSMCell") Or sheetName = getResByKey("UMTSCell") Or sheetName = getResByKey("LTECell") Then
        isCellSheet = True
        Exit Function
    End If
    isCellSheet = False
End Function
 
Function isTrasnPortSheet(sheetName As String) As Boolean
    If sheetName = getResByKey("GSMCell") Or sheetName = getResByKey("UMTSCell") Or sheetName = getResByKey("LTECell") _
        Or sheetName = getResByKey("GTRXGROUP") Or sheetName = getResByKey("GTRX") Then
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

Sub initCellTemplate(ByVal sheet As Worksheet, ByVal target As range, myAttrName As String, myCellMocName As String, myType As String)
        Dim m_Cell_Template As String
        '「物理Cell Template」所在列
        constCellTempCol = getColNum(sheet.name, constRecordRow, myAttrName, myCellMocName)
        
        If constCellTempCol >= 0 And target.column = constCellTempCol And target.count = 1 And target.row > constRecordRow Then
            '获取「CellTemplate」列侯选值
            m_Cell_Template = getCellTemplate(myType, sheet, target)
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
      For m_rowNum = 2 To MappingCellTemplate.range("a1048576").End(xlUp).row
        If myType = MappingCellTemplate.Cells(m_rowNum, 2).value _
        And neType = MappingCellTemplate.Cells(m_rowNum, 3).value Then
                   If m_Str = "" Then
                        m_Str = MappingCellTemplate.Cells(m_rowNum, 1).value
                   Else
                        m_Str = m_Str & "," & MappingCellTemplate.Cells(m_rowNum, 1).value
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










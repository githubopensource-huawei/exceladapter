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


Private Const MCellMocName As String = "MCell"
Private Const MAttrName As String = "CellTemplateName"
Private Const MCellType As String = "NB-IoT Cell"

Function isCellExist() As Boolean
    If IsSheetExist("GSM Cell") Or IsSheetExist("UMTS Cell") Or IsSheetExist("LTE Cell") Or IsSheetExist("NB-IoT Cell") _
    Or IsSheetExist(getResByKey("A224")) Or IsSheetExist(getResByKey("A225")) Or IsSheetExist(getResByKey("A226")) Or IsSheetExist(getResByKey("A262")) Then
        isCellExist = True
    Else
         isCellExist = False
    End If
End Function

Function isBSTransPortSht(ByVal sheetName As String)
    isBSTransPortSht = False
    If sheetName = "Base Station Transport Data" Or sheetName = getResByKey("A124") Then
        isBSTransPortSht = True
        Exit Function
    End If
End Function

Function isCellSheet(sheetName As String) As Boolean
    If sheetName = "GSM Cell" Or sheetName = "UMTS Cell" Or sheetName = "LTE Cell" Or sheetName = "NB-IoT Cell" _
        Or sheetName = getResByKey("A227") Or sheetName = getResByKey("A228") Or sheetName = getResByKey("A229") Or sheetName = getResByKey("A262") Then
        isCellSheet = True
        Exit Function
    End If
    isCellSheet = False
End Function
 
Function isTrasnPortSheet(ByVal sheetName As String) As Boolean
    If sheetName = "GSM Cell" Or sheetName = "UMTS Cell" Or sheetName = "LTE Cell" Or sheetName = "NB-IoT Cell" _
        Or sheetName = getResByKey("A230") Or sheetName = getResByKey("A231") Or sheetName = getResByKey("A232") Or sheetName = getResByKey("A262") _
        Or sheetName = "GTRXGROUP" Or sheetName = getResByKey("A233") Or sheetName = "NB-IoT TRX" Or sheetName = getResByKey("NB-IoT TRX") Then
        isTrasnPortSheet = True
        Exit Function
    End If
    isTrasnPortSheet = False
End Function

Function isGsmCellSheet(sheetName As String) As Boolean
    If sheetName = "GSM Cell" Or sheetName = getResByKey("A234") Then
        isGsmCellSheet = True
        Exit Function
    End If
    isGsmCellSheet = False
End Function

Function isLteCellSheet(ByRef sheetName As String) As Boolean
    If sheetName = "LTE Cell" Or sheetName = getResByKey("A235") Then
        isLteCellSheet = True
    Else
        isLteCellSheet = False
    End If
End Function

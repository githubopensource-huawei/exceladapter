VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMoiRowsManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private groupNameRowNumber_ As Long
Private startRowNumber_ As Long
Private endRowNumber_ As Long
Private maxColumnNumber_ As Long
Private dstMoiRange_ As Range
Private dstRowsRange_ As Range
Private needFillInRangeMap_ As CMapValueObject
Private needFillInRangeAddressMap_ As CMap

Public Function checkNeedFillInCells(ByRef emptyCell As Range, ByRef emptyCellAddressString As String) As Boolean
    checkNeedFillInCells = True
    Dim needFillInRange As Range
    Dim eachCell As Range
    For Each needFillInRange In needFillInRangeMap_.ValueCollection
        For Each eachCell In needFillInRange
            If Trim(eachCell.value) = "" Then
                If emptyCell Is Nothing Then Set emptyCell = eachCell
                checkNeedFillInCells = False
                Call makeEmptyCellAddressString(eachCell.address(False, False), emptyCellAddressString)
            End If
        Next eachCell
    Next needFillInRange
End Function

Private Sub makeEmptyCellAddressString(ByRef cellAddress As String, ByRef emptyCellAddressString As String)
    If emptyCellAddressString = "" Then
        emptyCellAddressString = cellAddress
    Else
        emptyCellAddressString = emptyCellAddressString & ", " & cellAddress
    End If
End Sub

Public Sub addNeedFillInRange(ByRef columnName As String, ByRef needFillInRange As Range)
    Call needFillInRangeMap_.SetAt(columnName, needFillInRange)
End Sub

Public Function rangeInNeedFillInRange(ByRef certainRange As Range) As Boolean
    rangeInNeedFillInRange = True
    If needFillInRangeAddressMap_ Is Nothing Then Call initNeedFillInRangeAddressMap
    
    Dim eachCell As Range
    For Each eachCell In certainRange
        If Not needFillInRangeAddressMap_.hasKey(eachCell.address(False, False)) Then
            rangeInNeedFillInRange = False
            Exit Function
        End If
    Next eachCell
End Function

Private Sub initNeedFillInRangeAddressMap()
    Set needFillInRangeAddressMap_ = New CMap
    Dim needFillInRange As Range
    Dim eachCell As Range
    For Each needFillInRange In needFillInRangeMap_.ValueCollection
        For Each eachCell In needFillInRange
            Call needFillInRangeAddressMap_.SetAt(eachCell.address(False, False), eachCell.address(False, False))
        Next eachCell
    Next needFillInRange
End Sub

Public Function rangeInAddingRows(ByRef certainRange As Range) As Boolean
    rangeInAddingRows = False
    Dim startRowNumber As Long, endRowNumber As Long
    startRowNumber = certainRange.row
    endRowNumber = startRowNumber + certainRange.Rows.count - 1
    If startRowNumber >= startRowNumber_ And endRowNumber <= endRowNumber_ Then
        rangeInAddingRows = True
    End If
End Function

Public Property Let groupNameRowNumber(ByRef RowNumber As Long)
    groupNameRowNumber_ = RowNumber
End Property

Public Property Let startRowNumber(ByRef RowNumber As Long)
    startRowNumber_ = RowNumber
End Property

Public Property Let endRowNumber(ByRef RowNumber As Long)
    endRowNumber_ = RowNumber
End Property

Public Property Let maxColumnNumber(ByRef columnNumber As Long)
    maxColumnNumber_ = columnNumber
End Property

Public Property Get groupNameRowNumber() As Long
    groupNameRowNumber = groupNameRowNumber_
End Property

Public Property Get startRowNumber() As Long
    startRowNumber = startRowNumber_
End Property

Public Property Get endRowNumber() As Long
    endRowNumber = endRowNumber_
End Property

Public Property Get getMoiRange()
    If dstMoiRange_ Is Nothing Then
        Set dstMoiRange_ = currentSheet.Range(currentSheet.Cells(startRowNumber_, 1), currentSheet.Cells(endRowNumber_, maxColumnNumber_))
    End If
    Set getMoiRange = dstMoiRange_
End Property

Public Property Get getMoiRowsRange()
    If dstRowsRange_ Is Nothing Then
        Set dstRowsRange_ = currentSheet.Rows(startRowNumber_ & ":" & endRowNumber_)
    End If
    Set getMoiRowsRange = dstRowsRange_
End Property

Private Sub Class_Initialize()
    Set needFillInRangeMap_ = New CMapValueObject
    startRowNumber_ = -1
    endRowNumber_ = -1
End Sub

Private Sub Class_Terminate()
    startRowNumber_ = -1
    endRowNumber_ = -1
End Sub


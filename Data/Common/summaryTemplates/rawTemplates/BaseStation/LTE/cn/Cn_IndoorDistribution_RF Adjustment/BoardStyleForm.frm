VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BoardStyleForm 
   Caption         =   "扩展行"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7830
   OleObjectBlob   =   "BoardStyleForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "BoardStyleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Option Explicit

Private Const MaxMocNumber As Long = 10

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub FinishButton_Click()
    Unload Me
End Sub

Private Sub NextButton_Click()
    Call makeNewRecords
    Me.Hide
End Sub

Private Sub makeNewRecords()
    Dim groupName As String
    Dim moiNumber As Long
    groupName = Me.MocNameComboBox.value
    moiNumber = CLng(Me.MocNumberComboBox.value)
    
    Dim groupNameStartRowNumber As Long, groupNameEndRowNumber As Long
    Call getGroupNameStartAndEndRowNumber(groupName, groupNameStartRowNumber, groupNameEndRowNumber)
    
    Set selectedGroupMappingDefData = boardStyleMappingDefMap.GetAt(groupName)
    
    Call addNewLines(groupNameStartRowNumber + 2, groupNameEndRowNumber + 1, moiNumber, selectedGroupMappingDefData.totalColumnNumber)
    'Call setNewRangesStyle(groupNameEndRowNumber + 1, groupNameEndRowNumber + moiNumber, )
    
End Sub

Private Sub addNewLines(ByVal srcRowNumber As Long, ByVal dstRowNumber As Long, ByVal moiNumber As Long, ByVal maxColumnNumber As Long)
    Call insertNewLines(srcRowNumber, dstRowNumber, moiNumber)
    Call setNewRangesStyle(srcRowNumber, dstRowNumber, moiNumber, maxColumnNumber)
End Sub

Private Sub insertNewLines(ByVal srcRowNumber As Long, ByVal dstRowNumber As Long, ByVal moiNumber As Long)
    Dim k As Long
    Dim srcRowRange As range, dstRowRange As range
    Set srcRowRange = currentSheet.rows(srcRowNumber)
    Set dstRowRange = currentSheet.rows(dstRowNumber)
    For k = 1 To moiNumber
        srcRowRange.Copy
        dstRowRange.Insert Shift:=xlShiftDown
    Next k
    Application.CutCopyMode = False
End Sub

Private Sub setNewRangesStyle(ByVal srcRowNumber As Long, ByVal dstRowNumber As Long, ByVal moiNumber As Long, ByVal maxColumnNumber As Long)
    Call initMoiRowsManager(srcRowNumber, dstRowNumber, moiNumber, maxColumnNumber)
    Call setNewRangesBackgroundColour(NewMoiRangeColorIndex)
    
    Dim endRowNumber As Long
    endRowNumber = dstRowNumber + moiNumber - 1
    Call setNeedFillInRangesStyles(dstRowNumber, endRowNumber)
    Call clearBoardNoRanges(dstRowNumber, endRowNumber)
    
    Call selectCertainCell(currentSheet, "A" & srcRowNumber - 2)
End Sub

Private Sub initMoiRowsManager(ByVal srcRowNumber As Long, ByVal dstRowNumber As Long, ByVal moiNumber As Long, ByVal maxColumnNumber As Long)
    moiRowsManager.groupNameRowNumber = srcRowNumber - 2
    moiRowsManager.startRowNumber = dstRowNumber
    moiRowsManager.endRowNumber = dstRowNumber + moiNumber - 1
    moiRowsManager.maxColumnNumber = maxColumnNumber
End Sub

'Private Sub setNewRangesStyles()
'    Dim newMoiRange As Range
'    Set newMoiRange = moiRowsManager.getMoiRange
'    newMoiRange.Interior.colorIndex = NewMoiRangeColorIndex
'End Sub

Private Sub setNeedFillInRangesStyles(ByVal startRowNumber As Long, ByVal endRowNumber As Long)
    Dim needFillColumnNameCol As Collection
    Set needFillColumnNameCol = selectedGroupMappingDefData.needFillColumnNameCol
    Dim columnNamePositionLetterMap As CMap
    Set columnNamePositionLetterMap = selectedGroupMappingDefData.columnNamePositionLetterMap
    Dim columnName As Variant
    Dim columnNamePositionLetter As String
    Dim needFillInRange As range
    For Each columnName In needFillColumnNameCol
        columnNamePositionLetter = columnNamePositionLetterMap.GetAt(columnName)
        Set needFillInRange = currentSheet.range(currentSheet.range(columnNamePositionLetter & startRowNumber), currentSheet.range(columnNamePositionLetter & endRowNumber))
        needFillInRange.Interior.colorIndex = NeedFillInRangeColorIndex
        needFillInRange.ClearContents
        Call moiRowsManager.addNeedFillInRange(CStr(columnName), needFillInRange)
    Next columnName
End Sub

Private Sub initMocNameComboBox()
    Set boardStyleMappingDefMap = boardStyleData.getBoardStyleMappingDefMap
    
    Dim groupName As String
    Dim groupNameCollection As Collection
    Set groupNameCollection = getKeyValueCollection(boardStyleMappingDefMap.KeyCollection)
    
    If currentSheet Is Nothing Then Set currentSheet = ThisWorkbook.ActiveSheet
    
    Dim rowNumber As Long
    With Me.MocNameComboBox
        For rowNumber = 1 To currentSheet.range("A1048576").End(xlUp).row
            groupName = currentSheet.range("A" & rowNumber).value
            If Contains(groupNameCollection, groupName) Then
                If rowNumber = 1 Then
                    .AddItem groupName
                ElseIf rowIsBlank(currentSheet, rowNumber - 1) Then
                    .AddItem groupName
                End If
            End If
        Next rowNumber
    End With
    
    Dim selectedGroupName As String
    Call getSelectedGroupName(selectedGroupName)
    If Me.MocNameComboBox.ListCount <> 0 Then
        If Contains(groupNameCollection, selectedGroupName) Then '默认显示选定的分组
            Me.MocNameComboBox.value = selectedGroupName
        Else
            Me.MocNameComboBox.listIndex = 0
        End If
    End If
End Sub

Private Sub getSelectedGroupName(ByRef selectedGroupName As String)
    Dim selectionRange As range
    Set selectionRange = Selection
    selectedGroupName = selectionRange(1).value
End Sub

Private Sub initMocNumberComboBox()
    Dim number As Long
    With Me.MocNumberComboBox
        For number = 1 To MaxMocNumber
            .AddItem number
        Next number
    End With
    If Me.MocNumberComboBox.ListCount <> 0 Then Me.MocNumberComboBox.listIndex = 0
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    Set addBoardStyleMoiInProcess = New CAddingBoardStyleMoc
    Set moiRowsManager = New CMoiRowsManager
    Call addBoardStyleMoiInProcess.init
    If boardStyleData Is Nothing Then
        Call initBoardStyleMappingDataPublic
    End If
    Call initMocNameComboBox
    Call initMocNumberComboBox
    Exit Sub
ErrorHandler:
End Sub

Private Sub UserForm_Terminate()
    Set addBoardStyleMoiInProcess = Nothing
    Set moiRowsManager = Nothing
End Sub

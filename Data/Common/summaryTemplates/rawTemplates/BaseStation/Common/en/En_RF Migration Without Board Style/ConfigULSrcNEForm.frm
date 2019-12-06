VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConfigULSrcNEForm 
   Caption         =   "Configure Source NE Columns"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7185
   OleObjectBlob   =   "ConfigULSrcNEForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "ConfigULSrcNEForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Private Const MaxMocNumber As Long = 10


Private Sub CancelButton_Click()
  Unload Me
End Sub

Private Sub SubmitButton_Click()
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    Call makeRecords
    Call setRaletionBoard(ws.name)
ErrorHandler:
Unload Me
End Sub

Public Sub makeRecords()
On Error GoTo ErrorHandler
    Dim neType As String
    Dim ULColunmNum As Long
    Dim typeNum As String
    typeNum = getNeType

    If typeNum = "LTE" Then neType = "eNodeB"
    If typeNum = "UMTS" Or typeNum = "NEW_UMTS" Then neType = "NodeB"
    If neType = "" Then Exit Sub

    
    ULColunmNum = CLng(Me.ULNameColNumComboBox.value)
    
    Dim ulAllCol As Long
    Dim migrationNemap As CMigrationNeMap
    Set migrationNemap = New CMigrationNeMap
   
    ulAllCol = migrationNemap.srcNeNameColNum(neType)

    Dim chaValue As Long
    chaValue = ULColunmNum - ulAllCol
    If chaValue > 0 Then
        Call migrationNemap.insertSrcNeNameColumn(neType, chaValue)
    Else
        chaValue = ulAllCol - ULColunmNum
        Call migrationNemap.delSrcNeNameColRec(neType, chaValue)
    End If

    Exit Sub
ErrorHandler:
End Sub


Private Sub initSourceNeNameNumberComboBox()
    Dim number As Long
    Dim neType As String
    Dim typeNum As String
    typeNum = getNeType
    
    For number = 1 To MaxMocNumber
        ULNameColNumComboBox.AddItem number
    Next number

    If typeNum = "LTE" Then neType = "eNodeB"
    If typeNum = "UMTS" Or typeNum = "NEW_UMTS" Then neType = "NodeB"

    If neType = "" Then Exit Sub
    
    Dim ulAllCol As Long
    Dim migrationNemap As CMigrationNeMap
    Set migrationNemap = New CMigrationNeMap
   
    ulAllCol = migrationNemap.srcNeNameColNum(neType)

    Me.ULNameColNumComboBox.value = ulAllCol

End Sub

Private Sub initboxName()
    Dim typeNum As String
    typeNum = getNeType
    If typeNum = "LTE" Then
        Me.ULNameNumberLabel.Caption = getResByKey("ENODEB_NE_NAME_COL")
    ElseIf typeNum = "UMTS" Or typeNum = "NEW_UMTS" Then
        Me.ULNameNumberLabel.Caption = getResByKey("NODEB_NE_NAME_COL")
    Else
    End If
End Sub

Private Sub UserForm_Initialize()
On Error GoTo ErrorHandler
    
    Call initSourceNeNameNumberComboBox
    Call initboxName
    Exit Sub
ErrorHandler:
End Sub




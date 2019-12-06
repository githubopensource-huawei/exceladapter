VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConfigMRATSrcNEForm 
   Caption         =   "Configure Source NE Columns"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7350
   OleObjectBlob   =   "ConfigMRATSrcNEForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "ConfigMRATSrcNEForm"
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
    Call setBoard(ws.name)
ErrorHandler:
Unload Me
End Sub

Public Sub makeRecords()
On Error GoTo ErrorHandler
    Dim neType As String
    Dim btsColunmNum As Long
    
    
    btsColunmNum = CLng(Me.BtsNameColNumComboBox.value)
    
    
    Dim btsstartColNum As Long, btsendColNum As Long
    
    btsstartColNum = 0
    btsendColNum = 0
    
    
    neType = "BTS"
    Call getStartendColNum(neType, btsstartColNum, btsendColNum)
    
    
    Dim btsAllCol As Long, nodebAllCol As Long, enodebAllCol As Long
    If btsendColNum <> 0 Then
        btsAllCol = btsendColNum - btsstartColNum
    Else
        btsAllCol = 0
    End If
    


    Dim chaValue As Long
    neType = "BTS"
    chaValue = btsColunmNum - btsAllCol
    If btsAllCol = 0 And btsColunmNum <> 0 Then
        Call makeNewGroupNameColunmRecs(neType, btsColunmNum)
        Call makeNewRadioNameColunmRec(neType)
    ElseIf btsAllCol <> 0 And (chaValue > 0) Then
        Call makeNameColunmRecs(neType, chaValue)
    ElseIf btsAllCol <> 0 And (chaValue < 0) Then
        chaValue = btsAllCol - btsColunmNum
        'If btsColunmNum = 1 Then chaValue = btsAllCol
        Call delNameColunmRecs(neType, chaValue)
        Call delRadioColunmRecs(neType)
    Else
    End If
    
    

    Exit Sub
ErrorHandler:
End Sub


Private Sub initSourceNeNameNumberComboBox()
    Dim number As Long
    For number = 1 To MaxMocNumber
        BtsNameColNumComboBox.AddItem number
        
    Next number
    
    Dim btsstartColNum As Long, btsendColNum As Long
    
    btsstartColNum = 0
    btsendColNum = 0
    
    
    Dim neType As String
    
    neType = "BTS"
    Call getStartendColNum(neType, btsstartColNum, btsendColNum)
    
    
    Dim btsAllCol As Long, nodebAllCol As Long, enodebAllCol As Long
    If btsendColNum <> 0 Then
        btsAllCol = btsendColNum - btsstartColNum
    Else
        btsAllCol = 0
    End If
    
    
    Me.BtsNameColNumComboBox.value = btsAllCol
    
    
End Sub

Private Sub UserForm_Initialize()
On Error GoTo ErrorHandler
    
    Call initSourceNeNameNumberComboBox

    Exit Sub
ErrorHandler:
End Sub


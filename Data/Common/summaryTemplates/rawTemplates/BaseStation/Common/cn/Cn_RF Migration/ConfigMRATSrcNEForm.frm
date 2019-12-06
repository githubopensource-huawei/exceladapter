VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConfigMRATSrcNEForm 
   Caption         =   "配置网元名称列"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7515
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
    
    Dim warnStr As String
    Dim errorStr As String
    warnStr = getResByKey("Warning")
    errorStr = getResByKey("CAN_NOT_BE_ZERO")
'    If isCnEnv(ws) Then
'        warnStr = "警告"
'        errorStr = "源网元列数不能同时为0！"
'    Else
'        warnStr = "Warning"
'        errorStr = "Source NE Columns count can not be 0 at the same time!"
'    End If
    
    If CLng(Me.BtsNameColNumComboBox.value) = 0 And CLng(Me.eNodebNameColNumComboBox.value) = 0 And CLng(Me.NodebNameColNumComboBox.value) = 0 Then
        Call MsgBox(errorStr, vbInformation, warnStr)
        Exit Sub
    End If
    
    Call makeRecords
    Call setRaletionBoard(ws.name)
ErrorHandler:
Unload Me
End Sub

Public Sub makeRecords()
On Error GoTo ErrorHandler
    Dim neType As String
    Dim btsColunmNum As Long
    Dim enodebColunmNum As Long
    Dim nodebColunmNum As Long
    Dim migrationNemap As CMigrationNeMap
    Set migrationNemap = New CMigrationNeMap
    
    btsColunmNum = CLng(Me.BtsNameColNumComboBox.value)
    enodebColunmNum = CLng(Me.eNodebNameColNumComboBox.value)
    nodebColunmNum = CLng(Me.NodebNameColNumComboBox.value)
    
    Dim btsAllCol As Long, nodebAllCol As Long, enodebAllCol As Long

    neType = "BTS"
    btsAllCol = migrationNemap.srcNeNameColNum(neType)
    neType = "NodeB"
    nodebAllCol = migrationNemap.srcNeNameColNum(neType)
    neType = "eNodeB"
    enodebAllCol = migrationNemap.srcNeNameColNum(neType)
    


    Dim chaValue As Long
    neType = "BTS"
    chaValue = btsColunmNum - btsAllCol
    If chaValue > 0 Then
        Call migrationNemap.insertSrcNeNameColumn(neType, chaValue)
    Else
        chaValue = btsAllCol - btsColunmNum
        Call migrationNemap.delSrcNeNameColRec(neType, chaValue)
    End If
    
    neType = "NodeB"
    chaValue = nodebColunmNum - nodebAllCol
    If chaValue > 0 Then
        Call migrationNemap.insertSrcNeNameColumn(neType, chaValue)
    Else
        chaValue = nodebAllCol - nodebColunmNum
        Call migrationNemap.delSrcNeNameColRec(neType, chaValue)
    End If
    
    neType = "eNodeB"
    chaValue = enodebColunmNum - enodebAllCol
    If chaValue > 0 Then
        Call migrationNemap.insertSrcNeNameColumn(neType, chaValue)
    Else
        chaValue = enodebAllCol - enodebColunmNum
        Call migrationNemap.delSrcNeNameColRec(neType, chaValue)
    End If

    Exit Sub
ErrorHandler:
End Sub


Private Sub initSourceNeNameNumberComboBox()
    Dim number As Long
    For number = 0 To MaxMocNumber
        BtsNameColNumComboBox.AddItem number
        NodebNameColNumComboBox.AddItem number
        eNodebNameColNumComboBox.AddItem number
    Next number
    
    Dim neType As String
    Dim migrationNemap As CMigrationNeMap
    Set migrationNemap = New CMigrationNeMap
    
    Dim btsAllCol As Long, nodebAllCol As Long, enodebAllCol As Long

    neType = "BTS"
    btsAllCol = migrationNemap.srcNeNameColNum(neType)
    neType = "NodeB"
    nodebAllCol = migrationNemap.srcNeNameColNum(neType)
    neType = "eNodeB"
    enodebAllCol = migrationNemap.srcNeNameColNum(neType)
    
    Me.BtsNameColNumComboBox.value = btsAllCol
    Me.NodebNameColNumComboBox.value = nodebAllCol
    Me.eNodebNameColNumComboBox.value = enodebAllCol
    
End Sub

Private Sub UserForm_Initialize()
On Error GoTo ErrorHandler
    Call initSourceNeNameNumberComboBox

    Exit Sub
ErrorHandler:
End Sub



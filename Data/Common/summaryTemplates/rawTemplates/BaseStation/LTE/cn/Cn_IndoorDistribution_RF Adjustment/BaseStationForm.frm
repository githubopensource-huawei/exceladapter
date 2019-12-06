VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BaseStationForm 
   Caption         =   "指定基站名称"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6495
   OleObjectBlob   =   "BaseStationForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "BaseStationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub OKButton_Click()
    Dim baseStationName As String
    baseStationName = Me.BaseStationList.value
    Call AddSectorEqm(baseStationName)
    Unload Me
End Sub

'从「当前激活页获得基站名称」
Private Sub Set_BaseStation_Related()
    Dim rowNum As Long
    Dim maxRow As String
    Dim nowSelection As range
    Dim index As Long
    Dim baseStationCollection As Collection
    Dim baseStationName As Variant
    Dim selectBtsName As String
    Set baseStationCollection = New Collection
    BaseStationList.Clear
    
    Set nowSelection = Selection
    selectBtsName = ActiveSheet.Cells(nowSelection.row, 1).value

    maxRow = ActiveSheet.range("a1048576").End(xlUp).row
    For rowNum = 3 To maxRow
        baseStationName = ActiveSheet.Cells(rowNum, 1).value
        If existInCollection(baseStationName, baseStationCollection) = False Then
            baseStationCollection.Add (baseStationName)
        End If
    Next
    
    For Each baseStationName In baseStationCollection
        BaseStationList.AddItem (baseStationName)
    Next
    
    If baseStationCollection.count <> 0 Then
        Me.BaseStationList.listIndex = getIndexInCollection(selectBtsName, baseStationCollection)
    End If
End Sub

Private Function existInCollection(strValue As Variant, strCollection As Collection) As Boolean
    Dim sItem As Variant
    If Trim(CStr(strValue)) = "" Then
        existInCollection = True
        Exit Function
    End If
    For Each sItem In strCollection
        If sItem = strValue Then
            existInCollection = True
            Exit Function
        End If
    Next
    existInCollection = False
End Function

Private Function getIndexInCollection(strValue As Variant, strCollection As Collection) As Long
    Dim sItem As Variant
    Dim index As Long
    index = 0
    For Each sItem In strCollection
        If sItem = strValue Then
            getIndexInCollection = index
            Exit Function
        End If
        index = index + 1
    Next
    getIndexInCollection = 0
End Function


Private Sub UserForm_Initialize()
    Call Set_BaseStation_Related
End Sub





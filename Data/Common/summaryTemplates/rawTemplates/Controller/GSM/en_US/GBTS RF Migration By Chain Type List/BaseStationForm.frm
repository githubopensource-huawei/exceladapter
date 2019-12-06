VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BaseStationForm 
   Caption         =   "Select Base Station"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6525
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
    Dim CellSheetName As String
    baseStationName = Me.BaseStationList.value
    CellSheetName = ActiveSheet.name
    If IsGBTSTemplate() Then
        Call AddTrxBinds(baseStationName, CellSheetName)
    Else
        Call AddSectorEqm(baseStationName)
    End If
    
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
    If IsGBTSTemplate() Then
        selectBtsName = ActiveSheet.Cells(nowSelection.row, getGcellBTSNameCol(ActiveSheet.name)).value
    End If
    
    If CELL_TYPE = 0 Or CELL_TYPE = 4 Then
        baseStationCollection.Add (selectBtsName)
    Else
        maxRow = ActiveSheet.range("a1048576").End(xlUp).row
        For rowNum = 3 To maxRow
            baseStationName = ActiveSheet.Cells(rowNum, 1).value
            '@gbts
            If IsGBTSTemplate() Then
                baseStationName = ActiveSheet.Cells(rowNum, getGcellBTSNameCol(ActiveSheet.name)).value
            End If
    
            If existInCollection(baseStationName, baseStationCollection) = False And Trim(baseStationName) <> "" Then
                baseStationCollection.Add (baseStationName)
            End If
        Next
    End If
    
    For Each baseStationName In baseStationCollection
        If Trim(baseStationName) <> "" Then
            BaseStationList.AddItem (baseStationName)
        End If
            
    Next
    
    If baseStationCollection.count <> 0 Then
        Me.BaseStationList.ListIndex = getIndexInCollection(selectBtsName, baseStationCollection)
    End If
End Sub


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
    Call Upt_Desc
    Call Set_BaseStation_Related
End Sub

Private Sub Upt_Desc()
    BaseStationForm.Caption = getResByKey("BaseStationForm.Caption")
    BaseStationNameBox.Caption = getResByKey("BaseStationNameBox.Caption")
    OKButton.Caption = getResByKey("OKButton.Caption")
    CancelButton.Caption = getResByKey("CancelButton.Caption")
End Sub




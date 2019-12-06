VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeleteFreqForm 
   Caption         =   "Delete Frequency"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5595
   OleObjectBlob   =   "DeleteFreqForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "DeleteFreqForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit
Private freqCollection As Collection
Private SelectedSiteName As String
Private SelectedCellName As String
Private SelectedFreq As String
Private SelectedFreqIndex As Long
Private row As Long
'从「当前激活页获得基站名称」
Private Sub Set_BaseStation_Related()
    Dim rowNum As Long
    Dim maxRow As String
    Dim nowSelection As Range
    Dim index As Long
    Dim baseStationCollection As Collection
    Dim baseStationName As Variant
    Dim selectBtsName As String
    Set baseStationCollection = New Collection
    BaseStationList.Clear
    
    Set nowSelection = Selection
    selectBtsName = ActiveSheet.Cells(nowSelection.row, getGcellBTSNameCol(ActiveSheet.name)).value

    maxRow = ActiveSheet.Range("B1048576").End(xlUp).row
    For rowNum = 3 To maxRow
        baseStationName = ActiveSheet.Cells(rowNum, getGcellBTSNameCol(ActiveSheet.name)).value
            
        If existInCollection(baseStationName, baseStationCollection) = False And Trim(baseStationName) <> "" Then
            baseStationCollection.Add (baseStationName)
        End If
    Next
    
    For Each baseStationName In baseStationCollection
        If Trim(baseStationName) <> "" Then
            BaseStationList.AddItem (baseStationName)
        End If
            
    Next
    
    If baseStationCollection.count <> 0 Then
        Me.BaseStationList.listIndex = getIndexInCollection(selectBtsName, baseStationCollection)
    End If
End Sub
'从基站名称获得小区名称
Private Function set_Cell_Related(btsName As String)
    Dim rowNum As Long
    Dim maxRow As Long
    Dim cellNameCollection As Collection
    Dim nowSelection As Range
    Dim cellName As Variant
    Dim selectCellName As String
    
    Set cellNameCollection = New Collection
    Set nowSelection = Selection
    CellList.Clear
    
    Set nowSelection = Selection
    selectCellName = ActiveSheet.Cells(nowSelection.row, getGcellCellNameCol(ActiveSheet.name)).value
    
    maxRow = ActiveSheet.Range("B1048576").End(xlUp).row
    For rowNum = 3 To maxRow
        If btsName = ActiveSheet.Cells(rowNum, getGcellBTSNameCol(ActiveSheet.name)).value Then
            cellName = ActiveSheet.Cells(rowNum, getGcellCellNameCol(ActiveSheet.name)).value
            If existInCollection(cellName, cellNameCollection) = False And Trim(cellName) <> "" Then
                cellNameCollection.Add (cellName)
            End If
        End If
    Next
    
    For Each cellName In cellNameCollection
        If Trim(cellName) <> "" Then
            CellList.AddItem (cellName)
        End If
    Next
    
    If cellNameCollection.count <> 0 Then
        Me.CellList.listIndex = getIndexInCollection(selectCellName, cellNameCollection)
    End If
    
End Function
'从基站名称和小区名称获得频点列表
Private Function setFreq_Related(btsName As String, cellName As String)
    Dim rowNum As Long
    Dim maxRow As Long
    
    Dim nowSelection As Range
    Dim selectFreq As String
    Dim bcch As String
    Dim nonbcch As String
    Dim trxNum As String
    Dim cellBand As String
    Dim freqArray() As String
    Dim nonbcchArray() As String
    Dim trxNumArray() As String
    
    Set freqCollection = New Collection
    Set nowSelection = Selection
    FrequencyList.Clear
    
    Set nowSelection = Selection
    
    maxRow = ActiveSheet.Range("B1048576").End(xlUp).row
    For rowNum = 3 To maxRow
        If btsName = ActiveSheet.Cells(rowNum, getGcellBTSNameCol(ActiveSheet.name)).value _
            And cellName = ActiveSheet.Cells(rowNum, getGcellCellNameCol(ActiveSheet.name)).value Then
            bcch = ActiveSheet.Cells(rowNum, getBcchCol(ActiveSheet.name)).value
            nonbcch = ActiveSheet.Cells(rowNum, getNonBcchCol(ActiveSheet.name)).value
            trxNum = ActiveSheet.Cells(rowNum, getTrxNumCol(ActiveSheet.name)).value
            cellBand = ActiveSheet.Cells(rowNum, getCellBandCol(ActiveSheet.name)).value
            
            Dim index As Long
            Dim freqs As String
            If (Trim(nonbcch) <> "") Then
                If Trim(bcch) <> "" Then
                    freqs = bcch + "," + nonbcch
                Else
                    freqs = nonbcch
                End If
            Else
                freqs = bcch
            End If
            
            Dim allFreqs As String
            allFreqs = freqs
            
            If Trim(trxNum) = "" Or CLng(trxNum) = 0 Then
                selectFreq = ""
            Else
                trxNumArray = Split(trxNum, ",")
            
                If UBound(trxNumArray) = 1 Then
                    Call changeFreqs(freqs, trxNum, cellBand)
                Else
                    Dim trxInd As Long
                    trxInd = CLng(trxNum)
                    freqs = cutTail(freqs, trxInd)
                End If
            
                If Trim(freqs) <> "" Then
                    freqArray = Split(freqs, ",")
                    selectFreq = freqArray(0)
                    For index = LBound(freqArray) To UBound(freqArray)
                        Dim onefreq As Variant
                        onefreq = freqArray(index)
                        If onefreq <> "" Then
                            FrequencyList.AddItem (onefreq)
                        End If
                    Next
                Else
                    selectFreq = ""
                End If
            End If
            
            Dim allFreqArray() As String
            If Trim(allFreqs) <> "" Then
                allFreqArray = Split(allFreqs, ",")
                For index = LBound(allFreqArray) To UBound(allFreqArray)
                    onefreq = allFreqArray(index)
                    freqCollection.Add (onefreq)
                Next
            End If
            Exit For
        End If
    Next
        
    If freqCollection.count <> 0 Then
        If Me.FrequencyList.ListCount <> 0 Then
            Me.FrequencyList.listIndex = getIndexInCollection(selectFreq, freqCollection)
        End If
    End If
    
End Function
Private Function getGcellBTSNameCol(shtName As String) As Long
    getGcellBTSNameCol = getColNum(shtName, 2, "BTSNAME", "GCELL")
End Function
Private Function getGcellCellNameCol(shtName As String) As Long
    getGcellCellNameCol = getColNum(shtName, 2, "CELLNAME", "GCELL")
End Function
Private Function getBcchCol(shtName As String) As Long
    getBcchCol = getColNum(shtName, 2, "BCCHFREQ", "TRXINFO")
End Function
Private Function getNonBcchCol(shtName As String) As Long
    getNonBcchCol = getColNum(shtName, 2, "NONBCCHFREQLIST", "TRXINFO")
End Function
Private Function getTrxNumCol(shtName As String) As Long
    getTrxNumCol = getColNum(shtName, 2, "TRXNUM", "TRXINFO")
End Function
Private Function getCellBandCol(shtName As String) As Long
    getCellBandCol = getColNum(shtName, 2, "TYPE", "GCELL")
End Function
Private Function existInCollection(strValue As Variant, strCollection As Collection) As Boolean
    Dim sItem As Variant
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


Private Sub Upt_Desc()
    DeleteFreqForm.Caption = getResByKey("DeleteFreqForm.Caption")
    BaseStationNameBox.Caption = getResByKey("BaseStationNameBox.Caption")
    CellNameBox.Caption = getResByKey("CellNameBox.Caption")
    FrequencyBox.Caption = getResByKey("FrequencyBox.Caption")
    CommitButton.Caption = getResByKey("CommitButton.Caption")
    CancelButton.Caption = getResByKey("CancelButton.Caption")
End Sub

Private Sub BaseStationList_Change()
    SelectedSiteName = Me.BaseStationList.value
    Call set_Cell_Related(SelectedSiteName)
    SelectedCellName = Me.CellList.value
        
    Call setFreq_Related(SelectedSiteName, SelectedCellName)
End Sub


Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub CellList_Change()
    SelectedSiteName = Me.BaseStationList.value
    SelectedCellName = Me.CellList.value
    Call setFreq_Related(SelectedSiteName, SelectedCellName)
    
End Sub
Private Sub CommitButton_Click()
    Dim baseStationName As String
    Dim CellSheetName As String
    Dim freq As String
    baseStationName = Me.BaseStationList.value
    CellSheetName = Me.CellList.value
    freq = Me.FrequencyList.value
    
    Dim rowNum As Long
    Dim maxRow As Long
    
    CellSheetName = ActiveSheet.name
    
    maxRow = ActiveSheet.Range("B1048576").End(xlUp).row
    For rowNum = 3 To maxRow
        If SelectedSiteName = ActiveSheet.Cells(rowNum, getGcellBTSNameCol(ActiveSheet.name)).value _
            And SelectedCellName = ActiveSheet.Cells(rowNum, getGcellCellNameCol(ActiveSheet.name)).value Then
            row = rowNum
            Exit For
        End If
    Next
    
    Call deleteFreqAndAssocMo(CellSheetName, row, SelectedFreqIndex, freqCollection)
    
    If iseGBTSTemp() Then
        Call deleteFreqAssoEqm(row, SelectedFreqIndex)
    End If
    
    Call MsgBox(getResByKey("deleteTrxSuccess"), vbInformation, getResByKey("success"))
        
    Unload Me
End Sub

Private Sub FrequencyList_Change()
    SelectedSiteName = Me.BaseStationList.value
    SelectedCellName = Me.CellList.value
    SelectedFreq = Me.FrequencyList.value
    
    SelectedFreqIndex = getIndexInCollection(SelectedFreq, freqCollection)
End Sub

Private Sub UserForm_Initialize()
    Call Upt_Desc
    Call Set_BaseStation_Related
    Dim btsName As String
    Dim cellName As String
    btsName = Me.BaseStationList.value
    
    Call set_Cell_Related(btsName)
    cellName = Me.CellList.value
        
    Call setFreq_Related(btsName, cellName)
End Sub
Private Function iseGBTSTemp() As Boolean
    Dim cover As String
    Dim key As String
    
    cover = getResByKey("Cover")
    key = ThisWorkbook.Worksheets(cover).Cells(2, 2).value
    If key = getResByKey("A173") Or key = "GBTS RF Adjustment Data Workbook" Then
        iseGBTSTemp = False
    Else
        iseGBTSTemp = True
    End If
End Function

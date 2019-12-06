VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAntennes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private p_cellId As String
Private p_sectorIds As String
Private p_antennes As String
Private p_baseEqmIds As String
Private p_rowCount As Long
Private p_range As String
Private p_trx As String
Private p_row As Long
Private p_rowCount_Trx As Long

Property Let cellId(str As String)
        p_cellId = str
End Property

Property Let trxId(str As String)
        Dim index As Long
        Dim trxIdArray() As String
        p_rowCount_Trx = 0
        trxIdArray = Split(str, ",")
        p_rowCount_Trx = UBound(trxIdArray) + 1
        p_trx = str
End Property

Property Let sectorIds(str As String)
        p_sectorIds = str
End Property

Property Let baseEqmIds(str As String)
        p_baseEqmIds = str
End Property

Property Let ranges(str As String)
        p_range = str
End Property

Property Let row(val As Long)
        p_row = val
End Property

Property Let antennes(str As String)
        Dim antenneIndex As Long
        Dim sectorArray() As String
        p_rowCount = 0
        sectorArray = Split(str, ";")
        For antenneIndex = LBound(sectorArray) To UBound(sectorArray)
              p_rowCount = p_rowCount + UBound(Split(sectorArray(antenneIndex), ",")) + 1
        Next
        p_antennes = str
End Property

Property Get cellId() As String
        cellId = p_cellId
End Property

Property Get trxId() As String
        trxId = p_trx
End Property

Property Get sectorIds() As String
       sectorIds = p_sectorIds
End Property

Property Get baseEqmIds() As String
       baseEqmIds = p_baseEqmIds
End Property

Property Get ranges() As String
       ranges = p_range
End Property

Property Get antennes() As String
        antennes = p_antennes
End Property

Property Get rowCount() As String
        rowCount = max(p_rowCount, p_rowCount_Trx)
End Property

Property Get row() As Long
        row = p_row
End Property

Function getAntenneCollection(sflag As Boolean, rowVal As Long) As Collection
    On Error GoTo ErrorHandler
    Dim sectorArray() As String
    Dim sectorAntenneArray() As String
    Dim trxArray() As String
    Dim antenneArray() As String
    Dim baseEqmArray() As String
    Dim innerSector() As String
    Dim index As Long
    Dim innerIndex As Long
    Dim boardNum As String
    Dim antenneNo As String
    Dim modelT As String
    Dim baseEqmStr As String
    Dim antenneCollection As Collection
    Dim arrayLen As Long
    Set antenneCollection = New Collection
    
    sflag = True
    If CELL_TYPE = 0 Or CELL_TYPE = 4 Then
        sectorArray = Split(p_sectorIds, ";")
    Else
        sectorArray = Split(p_sectorIds, ",")
    End If
    sectorAntenneArray = Split(p_antennes, ";")
    If CELL_TYPE = 2 Then
        baseEqmArray = Split(p_baseEqmIds, ",")
    End If
    trxArray = Split(p_trx, ",")
    Dim errMsg As String
    If checkDataMatch(sectorArray, sectorAntenneArray, baseEqmArray) Or checkGSMCellDataError(sectorArray, sectorAntenneArray, trxArray) Then
        errMsg = "errorRecordExists"
        If CELL_TYPE = 2 Then
            errMsg = "errorRecordExistsInLte"
        End If
        Call MsgBox(getResByKey(errMsg), vbInformation, getResByKey("Warning"))
        sflag = False
        rowVal = p_row
        Exit Function
    End If
    If CELL_TYPE = 0 Then
        arrayLen = UBound(trxArray)
        For index = LBound(trxArray) To arrayLen
            If index <= UBound(sectorAntenneArray) Then
                If Trim(sectorAntenneArray(index)) = "" Then
                        Call setValueIntoAntenneCollection(antenneCollection, p_cellId, sectorArray(index), "", "", "", p_range, trxArray(index), "")
                Else
                    antenneArray = Split(sectorAntenneArray(index), ",")
                    innerSector = Split(sectorArray(index), ",")
                    For innerIndex = LBound(antenneArray) To UBound(antenneArray)
                        If Trim(antenneArray(innerIndex)) = "" Then
                            Call setValueIntoAntenneCollection(antenneCollection, p_cellId, innerSector(innerIndex), "", "", "", p_range, trxArray(index), "")
                        Else
                            Call getBoardNumAndAntenneNo(antenneArray(innerIndex), boardNum, antenneNo, modelT)
                            Call setValueIntoAntenneCollection(antenneCollection, p_cellId, innerSector(innerIndex), boardNum, antenneNo, modelT, p_range, trxArray(index), "")
                        End If
                     Next
                 End If
            Else
                Call setValueIntoAntenneCollection(antenneCollection, p_cellId, "", "", "", "", p_range, trxArray(index), "")
            End If
        Next
    Else
        For index = LBound(sectorAntenneArray) To UBound(sectorAntenneArray)
                If Trim(sectorAntenneArray(index)) = "" Then
                    Call setValueIntoAntenneCollection(antenneCollection, p_cellId, sectorArray(index), "", "", "", p_range, "", "")
                Else
                    antenneArray = Split(sectorAntenneArray(index), ",")
                     For innerIndex = LBound(antenneArray) To UBound(antenneArray)
                        baseEqmStr = ""
                        If CELL_TYPE = 2 Then
                            baseEqmStr = baseEqmArray(index)
                        End If
                        
                        If Trim(antenneArray(innerIndex)) = "" Then
                            Call setValueIntoAntenneCollection(antenneCollection, p_cellId, sectorArray(index), "", "", "", p_range, "", baseEqmStr)
                        Else
                            Call getBoardNumAndAntenneNo(antenneArray(innerIndex), boardNum, antenneNo, modelT)
                            Call setValueIntoAntenneCollection(antenneCollection, p_cellId, sectorArray(index), boardNum, antenneNo, modelT, p_range, "", baseEqmStr)
                        End If
                     Next
                 End If
        Next
    End If
    Set getAntenneCollection = antenneCollection
    Exit Function
ErrorHandler:
    Set getAntenneCollection = New Collection
End Function

Private Function checkGSMCellDataError(sectorArray() As String, sectorAntenneArray() As String, trxArray() As String) As Boolean
    Dim index As Long
    Dim innerIndex As Long
    Dim antenneArray() As String
    Dim innerSector() As String
    
    checkGSMCellDataError = False
    If CELL_TYPE = 0 Or CELL_TYPE = 4 Then
        For index = LBound(sectorAntenneArray) To UBound(sectorAntenneArray)
            antenneArray = Split(sectorAntenneArray(index), ",")
            innerSector = Split(sectorArray(index), ",")
            If UBound(antenneArray) <> UBound(innerSector) Then
                checkGSMCellDataError = True
                Exit Function
            End If
        Next
    End If
End Function

Private Function checkDataMatch(sectorArray() As String, sectorAntenneArray() As String, baseEqmArray() As String) As Boolean
    checkDataMatch = False
    If CELL_TYPE = 2 Then
        If UBound(sectorArray) <> UBound(sectorAntenneArray) Or UBound(sectorArray) <> UBound(baseEqmArray) Then
            checkDataMatch = True
            Exit Function
        End If
    Else
        If UBound(sectorArray) <> UBound(sectorAntenneArray) Then
            checkDataMatch = True
            Exit Function
        End If
    End If
End Function



Private Sub setValueIntoAntenneCollection(antenneCollection As Collection, cellStr As String, sectorStr As String, boardStr As String, anteStr As String, rsModel As String, rangeStr As String, trxStr As String, baseEqmStr As String)
    Dim antenne(7) As String
    antenne(0) = cellStr
    antenne(1) = sectorStr
    antenne(2) = boardStr
    antenne(3) = anteStr
    antenne(4) = rsModel
    antenne(5) = rangeStr
    antenne(6) = trxStr
    antenne(7) = baseEqmStr
    Call antenneCollection.Add(antenne)
End Sub

Private Sub getBoardNumAndAntenneNo(ByVal str As String, boardNum As String, antenneNo As String, modelT As String)
    Dim twoStr() As String
    Dim boardAnte As String
    Dim strs() As String
    Dim index As Long
    twoStr = Split(str, ":")
    boardAnte = twoStr(0)
    modelT = twoStr(1)
    strs = Split(boardAnte, "_")
    antenneNo = strs(UBound(strs))
    boardNum = ""
    For index = LBound(strs) To UBound(strs) - 1
        If boardNum = "" Then
              boardNum = strs(index)
        Else
              boardNum = boardNum + "_" + strs(index)
        End If
    Next
End Sub

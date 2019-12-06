VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLampsiteAntennes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private p_btsName As String
Private p_cellId As String
Private p_sectorIds As String
Private p_antennes As String
Private p_baseEqmIds As String
Private p_sectorEqmIds As String
Private p_sectorEqmGrpIds As String
'Private p_cellbeamModes As String
'◊”ø®√‚≈‰÷√
Private p_vmprfAutoCfg As String

Private p_rowCount As Long
Private p_range As String
Private p_trx As String
Private p_row As Long
Private p_rowCount_Trx As Long
'

Property Let vmprfAutoCfg(str As String)
    p_vmprfAutoCfg = str
End Property
Property Get vmprfAutoCfg() As String
    vmprfAutoCfg = p_vmprfAutoCfg
End Property

Property Let btsName(str As String)
        p_btsName = str
End Property

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
Property Let sectorEqmIds(str As String)
        p_sectorEqmIds = str
End Property
Property Let sectorEqmGrpIds(str As String)
        p_sectorEqmGrpIds = str
End Property
'Property Let cellbeamModes(str As String)
'        p_cellbeamModes = str
'End Property

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

Property Get btsName() As String
        btsName = p_btsName
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
Property Get sectorEqmIds() As String
       sectorEqmIds = p_sectorEqmIds
End Property
Property Get sectorEqmGrpIds() As String
       sectorEqmGrpIds = p_sectorEqmGrpIds
End Property
'Property Get cellbeamModes() As String
'       cellbeamModes = p_cellbeamModes
'End Property

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
    'Dim baseEqmArray() As String
    Dim sectorEqmGrpIdArray() As String
    Dim sectorEqmArray() As String
'    Dim cellbeamModeArray() As String
    Dim innerSector() As String
    Dim index As Long
    Dim innerIndex As Long
    Dim boardNum As String
    Dim antenneNo As String
    Dim modelT As String
    Dim baseEqmStr As String
    Dim sectorEqmStr As String
    Dim sectorEqmGrpIdStr As String
    Dim cellbeamModeStr As String
    Dim antenneCollection As Collection
    Dim arrayLen As Long
    Dim tempSectoreqmArray() As String
    Dim tempsectorEqmStr As String
    Dim vmprfAutoCfgArray() As String
    
    Set antenneCollection = New Collection
    
    sflag = True
'    If ANT_CELL_TYPE = 0 Or ANT_CELL_TYPE = 4 Then
'        sectorArray = Split(p_sectorIds, ";")
'    Else
    sectorArray = Split(p_sectorIds, ",")
'    End If
    'ÃÏœﬂ∂Àø⁄
    sectorAntenneArray = Split(p_antennes, ";")
    '◊”ø®√‚≈‰÷√
    vmprfAutoCfgArray = Split(p_vmprfAutoCfg, ",")

    trxArray = Split(p_trx, ",")
    
    Dim errMsg As String
    If checkDataMatch(sectorArray, sectorAntenneArray, vmprfAutoCfgArray) Then
        errMsg = "lampsiteErrorRecordExists"
        If ANT_CELL_TYPE = 2 Then
            errMsg = "lampsiteErrorRecordExists"
        End If
        Call MsgBox(getResByKey(errMsg), vbInformation, getResByKey("Warning"))
        sflag = False
        rowVal = p_row
        Exit Function
    End If
    
        For index = LBound(sectorAntenneArray) To UBound(sectorAntenneArray)
                If Trim(sectorAntenneArray(index)) = "" Then
                    Call setValueIntoAntenneCollection(antenneCollection, p_btsName, p_cellId, sectorArray(index), "", "", "", p_range, "", "", "", "", "")
                Else
                    antenneArray = Split(sectorAntenneArray(index), ",")
                    baseEqmStr = ""
                    sectorEqmStr = ""
                    sectorEqmGrpIdStr = ""
                    cellbeamModeStr = ""
                    
                    If isLampsiteLteSectoreqmGrpId Then sectorEqmGrpIdStr = p_sectorEqmGrpIds

                    
                     For innerIndex = LBound(antenneArray) To UBound(antenneArray)
                        
                        If Trim(antenneArray(innerIndex)) = "" Then
                            Call setValueIntoAntenneCollection(antenneCollection, p_btsName, p_cellId, sectorArray(index), "", "", "", p_range, "", baseEqmStr, sectorEqmStr, sectorEqmGrpIdStr, vmprfAutoCfgArray(index))
                        Else
                            Call getBoardNumAndAntenneNo(antenneArray(innerIndex), boardNum, antenneNo, modelT)
                            Call setValueIntoAntenneCollection(antenneCollection, p_btsName, p_cellId, sectorArray(index), boardNum, antenneNo, modelT, p_range, "", baseEqmStr, sectorEqmStr, sectorEqmGrpIdStr, vmprfAutoCfgArray(index))
                        End If
                     Next
                 End If
        Next
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
    If ANT_CELL_TYPE = 0 Or ANT_CELL_TYPE = 4 Then
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

Private Function checkDataMatch(sectorArray() As String, sectorAntenneArray() As String, vmprfAutoCfgArray() As String) As Boolean
    checkDataMatch = False
'    If ANT_CELL_TYPE = 2 Then
        If UBound(sectorArray) <> UBound(sectorAntenneArray) Or UBound(sectorArray) <> UBound(vmprfAutoCfgArray) Then
            checkDataMatch = True
            Exit Function
        End If
'        If isLampsiteLteSectoreqmGrpId And UBound(sectorArray) <> UBound(sectorEqmGrpIdArray) Then
'            checkDataMatch = True
'            Exit Function
'        End If
'        If isLteCellBeamMode And UBound(sectorArray) <> UBound(cellbeamModeArray) Then
'            checkDataMatch = True
'            Exit Function
'        End If
'    ElseIf ANT_CELL_TYPE = 1 And isUmtsSectoreqmPo Then
'        If UBound(sectorArray) <> UBound(sectorAntenneArray) Or UBound(sectorArray) <> UBound(sectorEqmArray) Then
'            checkDataMatch = True
'            Exit Function
'        End If
'    Else
'        If UBound(sectorArray) <> UBound(sectorAntenneArray) Then
'            checkDataMatch = True
'            Exit Function
'        End If
'    End If
End Function



Private Sub setValueIntoAntenneCollection(antenneCollection As Collection, btsName As String, cellStr As String, sectorStr As String, boardStr As String, anteStr As String, rsModel As String, rangeStr As String, trxStr As String, baseEqmStr As String, sectorEqmStr As String, sectorEqmGrpIdStr As String, cellbeamModeStr As String)
    Dim antenne(11) As String
    antenne(0) = btsName
    antenne(1) = cellStr
    antenne(2) = sectorStr
    antenne(3) = boardStr
    antenne(4) = anteStr
    antenne(5) = rsModel
    antenne(6) = rangeStr
    antenne(7) = trxStr
    antenne(8) = baseEqmStr
    antenne(9) = sectorEqmStr
    antenne(10) = sectorEqmGrpIdStr
    antenne(11) = cellbeamModeStr
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


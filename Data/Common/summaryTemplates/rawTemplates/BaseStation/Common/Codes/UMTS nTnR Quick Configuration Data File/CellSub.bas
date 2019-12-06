Attribute VB_Name = "CellSub"
Option Explicit

Private g_SectorEqmPropertyFlag As Boolean

'interface called by GUI
Public Sub mergeSectorAndRXUAntForGUI()
    On Error Resume Next
    Dim cellSht As Worksheet
    For Each cellSht In ThisWorkbook.Worksheets
        If isCellSheet(cellSht.name) Then
            Call mergeSectorAndRXUAnt(cellSht)
            ThisWorkbook.Save
            Exit Sub
        End If
    Next
End Sub

Public Sub mergeSectorAndRXUAnt(ByRef cellSht As Worksheet)
    On Error Resume Next
    Dim sectorIdCol As Integer
    Dim rxuAntNoCol As Integer
    Dim sectorEqmPropertyCol As Integer
    
    g_SectorEqmPropertyFlag = False
    
    sectorIdCol = findColumnByName(cellSht, getResByKey("Sector_ID"), 2)
    rxuAntNoCol = findColumnByName(cellSht, getResByKey("RXUAntNo."), 2)
    sectorEqmPropertyCol = findColumnByName(cellSht, getResByKey("SectorEqmProperty"), 2)
    
    Dim sectors() As String
    Dim rxuAntNos() As String
    Dim sectorEqmProperties() As String
    
    Dim rowIdx As Integer
    For rowIdx = 3 To getUsedRow(cellSht)
        Dim strSector As String
        Dim strRxuAntNo As String
        Dim strSectorEqmProperty As String
        
        If sectorIdCol <> -1 Then strSector = cellSht.Cells(rowIdx, sectorIdCol)
        If rxuAntNoCol <> -1 Then strRxuAntNo = cellSht.Cells(rowIdx, rxuAntNoCol)
        If sectorEqmPropertyCol <> -1 Then
            strSectorEqmProperty = cellSht.Cells(rowIdx, sectorEqmPropertyCol)
            g_SectorEqmPropertyFlag = True
        End If
        
        If strSector <> "" And strRxuAntNo <> "" Then
            Dim resultSectors As String
            Dim resultRxuAntNos As String
            Dim resultSectorEqmProperty As String
            
            sectors = Split(strSector, ",")
            rxuAntNos = Split(strRxuAntNo, ";")
            sectorEqmProperties = Split(strSectorEqmProperty, ",")

            Call matchSectorAndRxuAntNo(sectors, rxuAntNos, sectorEqmProperties, resultSectors, resultRxuAntNos, resultSectorEqmProperty)
            cellSht.Cells(rowIdx, sectorIdCol) = resultSectors
            cellSht.Cells(rowIdx, rxuAntNoCol) = resultRxuAntNos
            If g_SectorEqmPropertyFlag = True Then cellSht.Cells(rowIdx, sectorEqmPropertyCol) = resultSectorEqmProperty
        End If
    Next
End Sub

Public Sub matchSectorAndRxuAntNo(ByRef sectors() As String, ByRef rxuAntNos() As String, ByRef sectorEqmProperties() As String, ByRef resultSectors As String, ByRef resultRxuAntNos As String, ByRef resultSectorEqmProperty As String)
    On Error Resume Next
    Dim pos As Integer
    
    resultSectors = sectors(0)
    resultRxuAntNos = rxuAntNos(0)
    If g_SectorEqmPropertyFlag = True Then resultSectorEqmProperty = sectorEqmProperties(0)
    
    For pos = LBound(sectors) + 1 To UBound(sectors)
        If sectors(pos) = sectors(pos - 1) Then
            If g_SectorEqmPropertyFlag = False Then
                resultRxuAntNos = resultRxuAntNos & "," & rxuAntNos(pos)
            ElseIf sectorEqmProperties(pos) = sectorEqmProperties(pos - 1) Then 'only g_SectorEqmPropertyFlag = true, sectorEqmProperties is available
                resultRxuAntNos = resultRxuAntNos & "," & rxuAntNos(pos)
            Else
                resultSectors = resultSectors & "," & sectors(pos)
                resultRxuAntNos = resultRxuAntNos & ";" & rxuAntNos(pos)
                If g_SectorEqmPropertyFlag = True Then resultSectorEqmProperty = resultSectorEqmProperty & "," & sectorEqmProperties(pos)
            End If
        Else
            resultSectors = resultSectors & "," & sectors(pos)
            resultRxuAntNos = resultRxuAntNos & ";" & rxuAntNos(pos)
            If g_SectorEqmPropertyFlag = True Then resultSectorEqmProperty = resultSectorEqmProperty & "," & sectorEqmProperties(pos)
        End If
    Next
End Sub



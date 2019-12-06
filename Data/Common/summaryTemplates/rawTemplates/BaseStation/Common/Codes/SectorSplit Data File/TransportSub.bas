Attribute VB_Name = "TransportSub"
Option Explicit

Public g_SectorInfoCol As Integer
Public g_FreqInfoCol As Integer
Public g_SplitSectorCol As Integer
Public g_SplitFreqCol As Integer

Public Function splitSectorSelected(ByRef sh As Worksheet, ByRef target As range) As Boolean
    splitSectorSelected = False
    Call getSplitSectorCol
    
    If sh.name = GetMainSheetName() And target.column = g_SplitSectorCol Then
        splitSectorSelected = True
    End If
End Function

Public Function splitFreqSelected(ByRef sh As Worksheet, ByRef target As range) As Boolean
    splitFreqSelected = False
    Call getSplitFreqCol
    
    If sh.name = GetMainSheetName() And target.column = g_SplitFreqCol Then
        splitFreqSelected = True
    End If
End Function

Public Sub getSectorInfoCol()
    g_SectorInfoCol = colNum(g_CurrentSheet, getResByKey("SectorInfo"))
End Sub

Public Sub getFreqInfoCol()
    g_FreqInfoCol = colNum(g_CurrentSheet, getResByKey("FreqInfo"))
End Sub

Public Sub getSplitSectorCol()
    g_SplitSectorCol = colNum(g_CurrentSheet, getResByKey("SplitSector"))
End Sub

Public Sub getSplitFreqCol()
    g_SplitFreqCol = colNum(g_CurrentSheet, getResByKey("SplitFreq"))
End Sub


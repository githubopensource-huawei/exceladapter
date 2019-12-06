VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCellSheetHyperlinkFlag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sheetNameFlagMap_ As New CMap

Public Sub setAllSheetsFlag(ByRef flag As Boolean)
    Dim sheetName As Variant
    For Each sheetName In sheetNameFlagMap_.KeyCollection
        Call sheetNameFlagMap_.SetAt(sheetName, flag)
    Next sheetName
End Sub

Public Sub setSheetFlag(ByRef sheetName As String, ByRef flag As Boolean)
    Call sheetNameFlagMap_.SetAt(sheetName, flag)
End Sub

Public Function getSheetFlag(ByRef sheetName As String) As Boolean
    getSheetFlag = sheetNameFlagMap_.GetAt(sheetName, False)
End Function

Public Function allSheetsFlag() As Boolean
    allSheetsFlag = True
    Dim flag As Variant
    For Each flag In sheetNameFlagMap_.ValueCollection
        allSheetsFlag = allSheetsFlag And flag
    Next flag
End Function



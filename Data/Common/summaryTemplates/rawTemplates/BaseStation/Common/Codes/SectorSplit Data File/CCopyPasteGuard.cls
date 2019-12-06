VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCopyPasteGuard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim isCutCopyMode As Integer
Dim clipBoardData As DataObject

Public Sub init()

End Sub

Private Sub Class_Initialize()
On Error GoTo ErrorHandler
    isCutCopyMode = Application.CutCopyMode
    If isCutCopyMode = xlCut Or isCutCopyMode = xlCopy Then
        Set clipBoardData = New DataObject
        clipBoardData.Clear
        clipBoardData.GetFromClipboard
        Debug.Print "Data copied: " & clipBoardData.GetText
    End If
    Exit Sub
ErrorHandler:
    Set clipBoardData = Nothing
End Sub

Private Sub Class_Terminate()
On Error GoTo ErrorHandler
    If isCutCopyMode = xlCut Or isCutCopyMode = xlCopy Then
        With clipBoardData
            .SetText ""
            .PutInClipboard
        End With
    End If
    Exit Sub
ErrorHandler:
    Set clipBoardData = Nothing
End Sub


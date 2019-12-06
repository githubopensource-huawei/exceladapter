VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTraceLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private fs As Object
Private ts As Object
Private logPath_ As String

Public Function openFile(ByRef logPath As String) As Boolean
    On Error GoTo ErrorHandler
    openFile = True
    Set ts = fs.OpenTextFile(logPath, 8, True)
    logPath_ = logPath
    Exit Function
ErrorHandler:
    openFile = False
End Function

Public Function closeFile() As Boolean
    Class_Terminate
End Function


Public Sub writeText(Optional ByRef s As String = "")
    On Error GoTo ErrorHandler
    ts.WriteLine ("[" & Now & "] " & s)
    Exit Sub
ErrorHandler:
    If ts Is Nothing Then
        Dim logPath As String
        logPath_ = ThisWorkbook.Path & "\Output.txt"
        If False = openFile(logPath_) Then Exit Sub
        Resume 0
    End If
End Sub

Public Function logPath() As String
    logPath = logPath_
End Function

Private Sub Class_Initialize()
    On Error Resume Next
    Set fs = CreateObject("Scripting.FileSystemObject")
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    ts.Close
    Set ts = Nothing
    Set fs = Nothing
End Sub

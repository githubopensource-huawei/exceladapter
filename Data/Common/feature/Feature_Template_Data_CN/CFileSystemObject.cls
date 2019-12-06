VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFileSystemObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private fs_ As Object

Public Function fileExist(ByRef filePath As String) As Boolean
    If fs_.FileExists(filePath) Then
        fileExist = True
    Else
        fileExist = False
    End If
End Function

Public Function deleteFile(ByRef filePath As String) As Boolean
    On Error GoTo ErrorHandler
    If fileExist(filePath) Then
        fs_.deleteFile (filePath)
    End If
    Exit Function
ErrorHandler:
End Function

Private Sub Class_Initialize()
    Set fs_ = CreateObject("Scripting.FileSystemObject")
End Sub

Private Sub Class_Terminate()
    Set fs_ = Nothing
End Sub

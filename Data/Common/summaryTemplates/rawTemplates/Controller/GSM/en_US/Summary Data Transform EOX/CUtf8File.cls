VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CUtf8File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private utfStream
    
Const adTypeText = 2
Const adSaveCreateOverWrite = 2
Private FileName
Private strCont As String

Public Function CreateXmlFile(strFileName As String) As Boolean
    Dim TextStream
    Set utfStream = CreateObject("ADODB.Stream")
    
    FileName = strFileName
    strCont = ""
    CreateXmlFile = True
End Function

Public Sub AppendLine(strLine As String, Optional bAlignIdent As Boolean = False)
    strCont = strCont & strLine & vbCrLf
End Sub

Public Sub CloseXmlFile()
    Dim pos
    Dim strPath As String

    '保证路径存在
    pos = InStrRev(FileName, "\")
    If pos > 0 Then
        strPath = Left(FileName, pos - 1)
        If Not NewDir(strPath) Then
            SysErr getResByKey("A157") & vbCrLf & strPath & vbCrLf & getResByKey("A158")
        End If
    End If
    
    With utfStream
        .Open
        .Type = adTypeText
        .Charset = "UTF-8"
        .Position = utfStream.Size
        .WriteText strCont
        .SaveToFile FileName, adSaveCreateOverWrite
        .Close
    End With
    
    strCont = ""
    FileName = ""
    Set utfStream = Nothing
    

End Sub





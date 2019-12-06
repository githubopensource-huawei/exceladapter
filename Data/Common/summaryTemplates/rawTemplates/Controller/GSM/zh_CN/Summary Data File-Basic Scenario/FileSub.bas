Attribute VB_Name = "FileSub"
#If VBA7 Then
Public Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByRef lpMultiByteStr As Any, _
    ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As Any, _
    ByVal cchWideChar As Long) As Long
Public Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Any, _
    ByVal cchWideChar As Long, _
    ByRef lpMultiByteStr As Any, _
    ByVal cchMultiByte As Long, _
    ByVal lpDefaultChar As String, _
    ByVal lpUsedDefaultChar As Long) As Long
#Else
Public Declare Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByRef lpMultiByteStr As Any, _
    ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long) As Long
Public Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByRef lpMultiByteStr As Any, _
    ByVal cchMultiByte As Long, _
    ByVal lpDefaultChar As String, _
    ByVal lpUsedDefaultChar As Long) As Long
#End If

    
Function readUTF8File(strFile As String) As String
    Dim bByte As Byte
    Dim ReturnByte() As Byte
    Dim lngBufferSize As Long
    Dim strBuffer As String
    Dim lngResult As Long
    Dim bHeader(1 To 3) As Byte
    Dim i As Long
 
    On Error GoTo errHandle
    If Dir(strFile) = "" Then Exit Function
 
     ' 以二进制打开文件
    Open strFile For Binary As #1
    ReDim ReturnByte(0 To LOF(1) - 1) As Byte
    ' 读取前三个字节
    Get #1, , bHeader(1)
    Get #1, , bHeader(2)
    Get #1, , bHeader(3)
    ' 判断前三个字节是否为BOM头
    If bHeader(1) = 239 And bHeader(2) = 187 And bHeader(3) = 191 Then
        For i = 3 To LOF(1) - 1
            Get #1, , ReturnByte(i - 3)
        Next i
    Else
        ReturnByte(0) = bHeader(1)
        ReturnByte(1) = bHeader(2)
        ReturnByte(2) = bHeader(3)
        For i = 3 To LOF(1) - 1
            Get #1, , ReturnByte(i)
        Next i
    End If
    ' 关闭文件
    Close #1
 
    ' 转换UTF-8数组为字符串
    lngBufferSize = UBound(ReturnByte) + 1
    strBuffer = String$(lngBufferSize, vbNullChar)
    lngResult = MultiByteToWideChar(65001, 0, ReturnByte(0), _
        lngBufferSize, StrPtr(strBuffer), lngBufferSize)
    readUTF8File = Left(strBuffer, lngResult)
 
    Exit Function
errHandle:
    readUTF8File = ""
    Exit Function
End Function

' 将输入文本写进UTF8格式的文本文件
' strInput：文本字符串
' strFile：保存的UTF8格式文件路径
' bBOM：True表示文件带"EFBBBF"头，False表示不带
Sub WriteUTF8File(strInput As String, strFile As String, Optional bBOM As Boolean = True)
    On Error GoTo errHandle
    Dim bByte As Byte
    Dim ReturnByte() As Byte
    Dim lngBufferSize As Long
    Dim lngResult As Long
    Dim TLen As Long
 
    ' 判断输入字符串是否为空
    If Len(strInput) = 0 Then Exit Sub
    ' 判断文件是否存在，如存在则删除
    If Dir(strFile) <> "" Then Kill strFile
 
    TLen = Len(strInput)
    lngBufferSize = TLen * 3 + 1
    ReDim ReturnByte(lngBufferSize - 1)
    lngResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strInput), TLen, _
        ReturnByte(0), lngBufferSize, vbNullString, 0)
    If lngResult Then
        lngResult = lngResult - 1
        ReDim Preserve ReturnByte(lngResult)
        Open strFile For Binary As #1
        If bBOM = True Then
            bByte = 239
            Put #1, , bByte
            bByte = 187
            Put #1, , bByte
            bByte = 191
            Put #1, , bByte
        End If
        Put #1, , ReturnByte
        Close #1
    End If
    Exit Sub
errHandle:
    Exit Sub
End Sub


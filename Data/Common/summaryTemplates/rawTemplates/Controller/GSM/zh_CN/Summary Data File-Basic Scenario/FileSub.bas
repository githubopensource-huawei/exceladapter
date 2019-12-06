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
 
     ' �Զ����ƴ��ļ�
    Open strFile For Binary As #1
    ReDim ReturnByte(0 To LOF(1) - 1) As Byte
    ' ��ȡǰ�����ֽ�
    Get #1, , bHeader(1)
    Get #1, , bHeader(2)
    Get #1, , bHeader(3)
    ' �ж�ǰ�����ֽ��Ƿ�ΪBOMͷ
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
    ' �ر��ļ�
    Close #1
 
    ' ת��UTF-8����Ϊ�ַ���
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

' �������ı�д��UTF8��ʽ���ı��ļ�
' strInput���ı��ַ���
' strFile�������UTF8��ʽ�ļ�·��
' bBOM��True��ʾ�ļ���"EFBBBF"ͷ��False��ʾ����
Sub WriteUTF8File(strInput As String, strFile As String, Optional bBOM As Boolean = True)
    On Error GoTo errHandle
    Dim bByte As Byte
    Dim ReturnByte() As Byte
    Dim lngBufferSize As Long
    Dim lngResult As Long
    Dim TLen As Long
 
    ' �ж������ַ����Ƿ�Ϊ��
    If Len(strInput) = 0 Then Exit Sub
    ' �ж��ļ��Ƿ���ڣ��������ɾ��
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


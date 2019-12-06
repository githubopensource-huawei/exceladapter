Attribute VB_Name = "Module1"
Option Explicit

Private Enum E_Language

    CHN = 1
    ENG = 2

End Enum


Private versionEndRow As Integer
Private mergeStartRow As Integer

Private language As E_Language

#If VBA7 Then
    Public Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
        ByVal CodePage As Long, _
        ByVal dwFlags As Long, _
        ByRef lpMultiByteStr As Any, _
        ByVal cchMultiByte As Long, _
        ByVal lpWideCharStr As Any, _
    ByVal cchWideChar As Long) As Long
#Else
    Public Declare Function MultiByteToWideChar Lib "kernel32" ( _
        ByVal CodePage As Long, _
        ByVal dwFlags As Long, _
        ByRef lpMultiByteStr As Any, _
        ByVal cchMultiByte As Long, _
        ByVal lpWideCharStr As Any, _
    ByVal cchWideChar As Long) As Long
#End If

Private Sub SortWorksheets()
'以升序排列工作表
    Dim sCount As Integer, i As Integer, j As Integer
    Application.ScreenUpdating = False
    sCount = Worksheets.Count
    If sCount = 1 Then Exit Sub
    For i = 1 To sCount - 1
        For j = i + 1 To sCount
            If Worksheets(j).Name < Worksheets(i).Name Then
                Worksheets(j).Move Before:=Worksheets(i)
            End If
        Next j
    Next i
End Sub

Private Sub NextSheet()
    If ActiveSheet.Index <> Worksheets.Count Then
        ActiveSheet.Next.Activate
    End If
End Sub


Private Sub setFrame(mergeStartRow As Integer, mergeEndRow As Integer)
    Range("A" + CStr(mergeStartRow) + ":" + "B" + CStr(mergeEndRow)).Select
    Selection.BorderAround Weight:=xlMedium
    With Selection.Borders(xlInsideVertical)
        .Weight = xlMedium
    End With
    Range("A" + CStr(mergeEndRow + 3)).Select
    
End Sub
Private Sub setFrameNone(mergeStartRow As Integer, mergeEndRow As Integer)
    Range("A" + CStr(mergeStartRow) + ":" + "B" + CStr(mergeEndRow)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A" + CStr(mergeEndRow + 2)).Select
End Sub

Private Sub readCfg()
    Dim CfgArray() As String
    
    CfgArray = Split(Trim(readUTF8File(ThisWorkbook.Path + "\Cfg.ini")), "*")
    If (CfgArray(0) = "en") Then
        language = ENG
    Else
        language = CHN
    End If
    
    versionEndRow = CInt(CfgArray(1))
    mergeStartRow = CInt(CfgArray(2))
End Sub

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
    MsgBox Err.Description, , "ERROR - " & Err.Number
    readUTF8File = ""
End Function











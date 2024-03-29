Attribute VB_Name = "Util"
Option Explicit

Public Const HyperLinkColorIndex = 6
Public Const BluePrintSheetColor = 5
Public Const MaxChosenSiteNum = 202

Public bIsEng As Boolean  '用于控制设置中英文

Private Const SW_SHOWNORMAL = 1
Public HasHistoryData As Boolean

Public Const NeType_GSM = "GSM"
Public Const NeType_UMTS = "UMTS"
Public Const NeType_LTE = "LTE"
Public Const NeType_MRAT = "MRAT"
Public Const NeType_USU = "USU"
Public Const NeType_ICS = "ICS"
Public Const NeType_CBS = "CBS"
Public Const NeType_5G = "NR"
Public Const NeType_DSA = "DSA"

Public Const SheetType_List = "LIST"
Public Const SheetType_Pattern = "PATTERN"
Public Const SheetType_Main = "MAIN"
Public Const SheetType_Common = "COMMON"
Public Const SheetType_Board = "BOARD"
Public Const SheetType_Iub = "IUB"

Global Const StartRow_Name As String = "StartRow"
Global Const EndRow_Name As String = "EndRow"
Global Const BaseSheetName_Name As String = "Base Sheet Name"

Private Const listShtTitleRow = 2


Public Function getNeType() As String
    On Error Resume Next
    Dim cover As String
    Dim key As String
    Dim reValue As String
    
    cover = getResByKey("Cover")
    key = ThisWorkbook.Worksheets(cover).Cells(2, 2).value
    reValue = getResByKey(key)
    
    If reValue = key Then
       reValue = "MRAT"
    End If
    
    Select Case reValue
        Case "GSM"
            getNeType = NeType_GSM
        Case "UMTS"
            getNeType = NeType_UMTS
        Case "LTE"
            getNeType = NeType_LTE
        Case "MRAT"
            getNeType = NeType_MRAT
        Case "USU"
            getNeType = NeType_USU
        Case "ICS"
            getNeType = NeType_ICS
        Case "DSA"
            getNeType = NeType_DSA
        Case Else
            getNeType = ""
    End Select
End Function

'调用方法：GetFileName("话统脚本, *.sql", "Open")
Public Function GetFileName(ByVal strFilter, ByVal strTitle, ByVal bMulSel As Boolean, vFileName As Variant) As Boolean
    
    Dim vRsp
    Dim i As Long
    
    GetFileName = False
    vRsp = Application.GetOpenFilename(FileFilter:=strFilter, title:=strTitle, MultiSelect:=bMulSel)
    If IsArray(vRsp) Then
        GetFileName = True
        
        ReDim vFileName(UBound(vRsp) - 1)
        For i = 1 To UBound(vRsp)
            vFileName(i - 1) = vRsp(i)
        Next i
    ElseIf vRsp <> False Then
        vFileName = vRsp
        GetFileName = True
    End If
End Function

Public Function ReadFile(ByVal strFileName As String) As Collection
    
    Dim FN As Long
    Dim strLine As String
    Dim colRead As New Collection
    
    FN = FreeFile()
    Open strFileName For Input As #FN
    
    While Not EOF(FN)
        Line Input #FN, strLine
        colRead.Add strLine
    Wend
    Close #FN
    
    Set ReadFile = colRead
    
End Function

Public Function WriteFile(ByVal strFileName As String, colWri As Collection, Optional bOverLay As Boolean = False, Optional bMarkSeq As Boolean = False) As Boolean
    Dim FN
    Dim v
    Dim strLine As String
    Dim Seq As Long
    Dim strPath As String
    Dim pos As Long
    
    '保证路径存在
    pos = InStrRev(strFileName, "\")
    If pos > 0 Then
        strPath = Left(strFileName, pos - 1)
        If Not NewDir(strPath) Then
            SysErr "File path" & vbCrLf & strPath & vbCrLf & getResByKey("Invalid")
            Exit Function
        End If
    End If
    
    On Error GoTo ErrExit
    
    If bOverLay Then KillFile strFileName
    
    FN = FreeFile()
    Open strFileName For Append As #FN

    Seq = 1
    For Each v In colWri
        strLine = v
        If bMarkSeq Then strLine = Seq & vbTab & strLine
        Print #FN, strLine
        Seq = Seq + 1
    Next v
    Close #FN
ErrExit:
    WriteFile = (Err = 0)
End Function

Public Function LikeEx(ByVal strExp, ByVal strCmp, Optional bCaseSensitive As Boolean = True) As Boolean
    
    If Not bCaseSensitive Then
        strExp = UCase(strExp)
        strCmp = UCase(strCmp)
    End If
    
    LikeEx = (strExp Like strCmp)
    
End Function

Public Function SubString(ByVal s As String, ByVal nFrom As Long, ByVal nTo As Long) As String
    
    Dim nLen As Long
    
    nLen = nTo - nFrom
    If nLen >= 0 Then nLen = nLen + 1
    
    If nLen > 0 Then SubString = Mid(s, nFrom, nLen)
    
End Function

Public Function SplitEx(ByVal s As String, ByVal Sep As String) As Collection
    Dim colRet As New Collection
    Dim i As Long
    Dim v
    
    If s <> "" Then
        v = Split(s, Sep)
        For i = 0 To UBound(v)
            colRet.Add v(i)
        Next i
    End If
    
    Set SplitEx = colRet
End Function

Public Function GetCell(shtX As Worksheet, ByVal r, ByVal C)
    GetCell = shtX.Cells(r, C)
End Function

'写入：True；未写入：False
Public Function SetCell(shtX As Worksheet, ByVal r, ByVal C, ByVal strCellVal) As Boolean
    Dim strOld As String
    
    strOld = GetCell(shtX, r, C)
    If strOld <> CStr(strCellVal) Then
        shtX.Cells(r, C) = strCellVal
        SetCell = True
    End If
End Function

Public Function IsNullRow(shtX As Worksheet, ByVal r As Long, Optional FromCol, Optional ToCol) As Boolean
    Dim nFromCol As Long
    Dim nToCol As Long
    
    nFromCol = IIf(IsMissing(FromCol), 1, CellCol2Int(FromCol))
    nToCol = IIf(IsMissing(ToCol), 3, CellCol2Int(ToCol))
    
    Dim C As Long
    Dim strCell As String
    
    IsNullRow = False
    For C = nFromCol To nToCol
        strCell = GetCell(shtX, r, C)
        If LeftMostMatch(strCell, "//") Then Exit For '该行为注释行，视为空行
        If strCell <> "" Then Exit Function
    Next C
    IsNullRow = True
End Function

Public Sub SetNullRow(shtX As Worksheet, ByVal r As Long, Optional FromCol, Optional ToCol)
    
    If IsMissing(FromCol) Then FromCol = "A"
    If IsMissing(ToCol) Then FromCol = "C"
    
    If Not IsNumeric(FromCol) Then FromCol = Asc(FromCol) - Asc("A") + 1
    If Not IsNumeric(ToCol) Then ToCol = Asc(ToCol) - Asc("A") + 1
    
    Dim C As Long
    
    For C = FromCol To ToCol
        Call SetCell(shtX, r, C, "")
    Next C

End Sub

Public Function Col2String(colX As Collection, Sep) As String
    
    Dim v
    Dim strRet As String
    
    strRet = ""
    For Each v In colX
        strRet = strRet & Sep & v
    Next v
    Col2String = Mid(strRet, Len(Sep) + 1)
    
End Function

Public Function Arr2String(arr As Variant, Sep) As String
    Dim i As Long
    Dim strRet As String
    
    For i = 0 To UBound(arr)
        strRet = strRet & Sep & arr(i)
    Next i
    Arr2String = Mid(strRet, Len(Sep) + 1)
End Function

Public Function FormatCellAddr(ByVal r, ByVal C)
    FormatCellAddr = "$" & CellCol2Str(C) & "$" & r
End Function

Public Function FormatRangeAddr(ByVal r0, ByVal r1, ByVal c0, ByVal c1)
    FormatRangeAddr = FormatCellAddr(r0, c0) & ":" & FormatCellAddr(r1, c1)
End Function

Public Function FormatCollection(ParamArray itm()) As Collection
    Dim colX As New Collection
    Dim i As Long
    
    For i = 0 To UBound(itm)
        colX.Add itm(i)
    Next i
    
    Set FormatCollection = colX
End Function

Public Sub SysErr(ByVal s As String)
    LimitMsgBoxText s
    MsgBox s, vbCritical, "Error"
End Sub

Public Sub SysWarn(ByVal s As String)
    LimitMsgBoxText s
    MsgBox s, vbExclamation, "Warning"
End Sub

Public Sub SysInfo(ByVal s As String)
    LimitMsgBoxText s
    MsgBox s, vbInformation, "Information"
End Sub

Public Function SysAsk(ByVal strAsk As String, Style As VbMsgBoxStyle) As VbMsgBoxResult
    LimitMsgBoxText strAsk
    SysAsk = MsgBox(strAsk, Style + vbQuestion, "Question?")
End Function

Public Sub UniqueCol(colX As Collection)
    Dim colTemp As New Collection
    Dim v

    On Error Resume Next
    For Each v In colX
        colTemp.Add v, "Key_" & v
    Next v
    Set colX = colTemp
    
    If Err <> 0 Then Err.Clear
End Sub

Public Sub KillFile(strFile As String)
    On Error Resume Next
    
    SetAttr strFile, vbNormal
    Kill strFile
End Sub

Public Function NewDir(ByVal strPath) As Boolean
    Dim dirAttr As VbFileAttribute
    
    dirAttr = vbNormal + vbDirectory + vbReadOnly
    If Dir(strPath, dirAttr) <> "" Then
        NewDir = True
        Exit Function
    End If
    
    Dim pos As Long
    Dim strtemp As String
    
    On Error Resume Next
    pos = InStr(1, strPath, "\")
    While pos > 0
        strtemp = Left(strPath, pos - 1)
        If Dir(strtemp, dirAttr) = "" Then MkDir strtemp
        pos = InStr(pos + 1, strPath, "\")
    Wend
    
    MkDir strPath
    
    NewDir = (Dir(strPath, dirAttr) <> "")
    If Err <> 0 Then Err.Clear
End Function

Public Function IsSameCol(ByVal Col1, ByVal Col2) As Boolean
    If IsNumeric(Col1) And IsNumeric(Col2) Then
        IsSameCol = (Val(Col1) = Val(Col2))
    ElseIf (Not IsNumeric(Col1)) And (Not IsNumeric(Col2)) Then
        IsSameCol = (Col1 = Col2)
    Else
        IsSameCol = (CellCol2Str(Col1) = CellCol2Str(Col2))
    End If
End Function

Public Function LookupSheet(ByVal strName As String, shtX As Worksheet) As Boolean
    strName = UCase(strName)
    LookupSheet = False
    For Each shtX In ThisWorkbook.Sheets
        If UCase(shtX.CodeName) = strName Then
            LookupSheet = True
            Exit For
        End If
    Next shtX
End Function

Public Sub HideSheetCol(shtX As Worksheet, ByVal col, Optional bHide As Boolean = True)
    Dim rng As range

    col = CellCol2Str(col)
    Set rng = MakeRange(shtX, 1, col)
    
    If bHide Then
        If IsColHide(shtX, col) Then Exit Sub
    Else
        If Not IsColHide(shtX, col) Then Exit Sub
    End If
    
    rng.EntireColumn.ColumnWidth = IIf(bHide, 0, 8)
End Sub

Public Function IsColHide(shtX As Worksheet, ByVal col) As Boolean
    Dim rng As range
    
    col = CellCol2Str(col)
    Set rng = MakeRange(shtX, 1, col)
    IsColHide = (rng.EntireColumn.ColumnWidth < 1)
End Function

Public Sub HideSheetRow(shtX As Worksheet, ByVal row, Optional bHide As Boolean = True)
    Dim rng As range
    
    Set rng = MakeRange(shtX, row, "A")
    
    If bHide Then
        If IsRowHide(shtX, row) Then Exit Sub
    Else
        If Not IsRowHide(shtX, row) Then Exit Sub
    End If
    
    rng.EntireRow.RowHeight = IIf(bHide, 0, 11.25)
End Sub

Public Sub DelSheetRow(shtX As Worksheet, ByVal r0 As Long, Optional r1)
    If IsMissing(r1) Then r1 = r0
    MakeRange(shtX, r0, "A", r1, "A").EntireRow.Delete
End Sub

Public Sub InsertSheetRow(shtX As Worksheet, ByVal row As Long, Optional Num = 1)
    MakeRange(shtX, row, "A", row + Num - 1).EntireRow.INSERT
End Sub

Public Function IsRowHide(shtX As Worksheet, ByVal row) As Boolean
    Dim rng As range
    
    Set rng = MakeRange(shtX, row, "A")
    IsRowHide = (rng.EntireRow.RowHeight < 1)
End Function

Public Function LeftMostMatch(ByVal s As String, ByVal strLeft As String) As Boolean
    LeftMostMatch = Left(s, Len(strLeft)) = strLeft
End Function
Public Function RightMostMatch(ByVal s As String, ByVal strRight As String) As Boolean
    RightMostMatch = (Right(s, Len(strRight)) = strRight)
End Function

Public Function InCollection(colX As Collection, ByVal vItem, Optional bCaseSentive As Boolean = True) As Boolean
    Dim v
    Dim bOk As Boolean
    
    If (Not bCaseSentive) And (Not IsNumeric(vItem)) Then vItem = UCase(vItem)
    
    For Each v In colX
        If IsNumeric(v) And IsNumeric(vItem) Then
            bOk = (Val(v) = Val(vItem))
        Else
            If bCaseSentive Then
                bOk = (CStr(v) = CStr(vItem))
            Else
                bOk = (UCase(v) = CStr(vItem))
            End If
        End If
        If bOk Then Exit For
    Next v
    
    InCollection = bOk
End Function

Public Function InArray(ByVal vItem, ParamArray Choice() As Variant) As Boolean
    Dim i As Long
    
    InArray = False
    For i = 0 To UBound(Choice)
        If vItem = Choice(i) Then
            InArray = True
            Exit For
        End If
    Next i
End Function

Public Function IsSubSet(colFather As Collection, colChild As Collection) As Boolean
    Dim v
    
    IsSubSet = False
    For Each v In colChild
        If Not InCollection(colFather, v) Then Exit Function
    Next v
    IsSubSet = True
End Function

Public Function HasOneOf(ByVal s As String, ByVal strCharSet As String) As Boolean
    Dim i As Long
    For i = 1 To Len(strCharSet)
        If InStr(1, s, Mid(strCharSet, i, 1)) > 0 Then
            HasOneOf = True
            Exit For
        End If
    Next i
End Function

Public Function IsSymbol(ByVal strSymbol As String) As Boolean
    Dim i As Long
    Dim ch As String
    Dim bOk As Boolean
    
    strSymbol = UCase(strSymbol)
    
    ch = Left(strSymbol, 1)
    bOk = (ch = "_")
    If Not bOk Then bOk = (("A" <= ch) And (ch <= "Z"))
    If Not bOk Then Exit Function
    
    For i = 1 To Len(strSymbol)
        ch = Mid(strSymbol, i, 1)
        bOk = (ch = "_")
        If Not bOk Then bOk = (("0" <= ch) And (ch <= "9")) Or (("A" <= ch) And (ch <= "Z"))
        If Not bOk Then Exit Function
    Next i
    IsSymbol = True
End Function

Public Function MakeRange(shtX As Worksheet, r0, c0, Optional r1, Optional c1) As range
    If IsMissing(r1) Then r1 = r0
    If IsMissing(c1) Then c1 = c0
    Set MakeRange = shtX.range(shtX.Cells(r0, CellCol2Str(c0)), shtX.Cells(r1, CellCol2Str(c1)))
End Function

Public Function max(ByVal a, ByVal b)
    max = IIf(a > b, a, b)
End Function

Public Function min(ByVal a, ByVal b)
    min = IIf(a < b, a, b)
End Function

Public Function LenEx(ByVal s As String) As Long
    Dim i As Long
    Dim nNum As Long
    Dim ch As String
    
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        nNum = nNum + 1
        If Asc(ch) < 0 Then nNum = nNum + 1
    Next i
    
    LenEx = nNum
    
End Function

Public Function HasBigChar(ByVal s As String) As Boolean
    Dim i As Long
    Dim ch As String
    
    HasBigChar = False
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If Asc(ch) < 0 Then
            HasBigChar = True
            Exit For
        End If
    Next i
End Function

Public Function IsMacroStr(ByVal s As String) As Boolean
    Dim i As Long
    Dim nAscII As Long
    Dim bOk As Boolean
    
    IsMacroStr = False
    s = UCase(Trim(s))
    If s = "" Then Exit Function
        
    nAscII = Asc(Left(s, 1))
    If Between(nAscII, Asc("A"), Asc("Z")) Or (nAscII = Asc("_")) Then   '首字母必须是A~Z或“_”
        For i = 2 To Len(s) '非首字符必须是: A~Z, 0~9, _ 等字符
            nAscII = Asc(Mid(s, i, 1))
            bOk = Between(nAscII, Asc("A"), Asc("Z"))
            If Not bOk Then bOk = Between(nAscII, Asc("0"), Asc("9"))
            If Not bOk Then bOk = (nAscII = Asc("_"))
            If Not bOk Then Exit Function
        Next i
        IsMacroStr = True
    End If
End Function

Public Function GetWorkSheet(wb As Workbook, ByVal strSheetName As String) As Worksheet
    Dim shtX  As Worksheet
    
    strSheetName = UCase(strSheetName)
    For Each shtX In wb.Sheets
        If (UCase(shtX.CodeName) = strSheetName) Or (UCase(shtX.name) = strSheetName) Then
            Set GetWorkSheet = shtX
            Exit For
        End If
    Next shtX
End Function

Public Sub AppendCollection(colOld As Collection, colAppend As Collection)
    Dim v
    
    For Each v In colAppend
        colOld.Add v
    Next v
End Sub

Public Sub CfmArray(v As Variant)
    Dim vOld
    If Not IsArray(v) Then
        vOld = v
        ReDim v(0)
        v(0) = vOld
    End If
End Sub

Public Sub Exchange(a, b)
    Dim tmp
    
    tmp = a
    a = b
    b = tmp
End Sub

Public Sub ShowInfo(colInfo As Collection, ByVal strInfo As String, Optional strFile)
    Dim strWri As String
    
    If Not IsMissing(strFile) Then strWri = ThisWorkbook.Path & "\" & strFile
    If strWri = "" Then
        SysInfo strInfo & vbCrLf & Col2String(colInfo, vbCrLf)
    Else
        strInfo = strInfo & vbCrLf & "The error information has been record in " & vbCrLf & strWri
        WriteFile strWri, colInfo, True, True
        SysInfo strInfo
        Shell "notepad.exe " & strWri, vbMaximizedFocus
    End If
End Sub

Private Sub LimitMsgBoxText(s As String)
    Const nMaxLen As Long = 250
    If Len(s) > nMaxLen Then s = Left(s, nMaxLen) & "..."
End Sub

Public Function OpenWorkBook(ByVal strFileName As String) As Workbook
    On Error GoTo ErrExit
    Dim bookRead As Workbook
    Set bookRead = Workbooks.Open(FileName:=strFileName)
    If Err = 0 Then bookRead.RunAutoMacros xlAutoOpen + xlAutoActivate
    
ErrExit:
    If Err = 0 Then
        Set OpenWorkBook = bookRead
    Else
        Set OpenWorkBook = Nothing
    End If
End Function

Public Sub CloseWorkBook(book As Workbook, bSaveChanges As Boolean)
    book.Close savechanges:=bSaveChanges
End Sub

Public Function GetRowStr(shtX As Worksheet, ByVal row As Long, ByVal ColFrom, ByVal ColTo) As String
    AssertEx (Len(ColFrom) = 1)
    AssertEx (Len(ColTo) = 1)
    
    Dim col As Long
    Dim s As String
    
    For col = Asc(ColFrom) To Asc(ColTo)
        s = s & vbTab & GetCell(shtX, row, Chr(col))
    Next col
    s = Mid(s, 2)
    GetRowStr = s
End Function

Public Function GetRowHeight(shtX As Worksheet, ByVal row As Long) As Variant
    GetRowHeight = MakeRange(shtX, row, "A").EntireRow.RowHeight
End Function

Public Sub SetRowHeight(shtX As Worksheet, ByVal row As Long, ByVal lHeight As Variant)
    MakeRange(shtX, row, "A").EntireRow.RowHeight = lHeight
End Sub

Public Sub SetColWidth(shtX As Worksheet, ByVal col, ByVal lWidth)
    MakeRange(shtX, 1, col).EntireColumn.ColumnWidth = lWidth
End Sub

Public Function GetColWidth(shtX As Worksheet, ByVal col) As Variant
    GetColWidth = MakeRange(shtX, 1, col).EntireColumn.ColumnWidth
End Function

Public Function CountSubStrNum(ByVal s As String, ByVal strSubStr As String) As Long
    Dim Num As Long
    Dim pos As Long
    Dim nSubStrLen As String
    
    nSubStrLen = Len(strSubStr)
    pos = InStr(1, s, strSubStr)
    
    While pos > 0
        Num = Num + 1
        pos = InStr(pos + nSubStrLen, s, strSubStr)
    Wend
    CountSubStrNum = Num
End Function

Public Sub AssertEx(Optional bCondition As Boolean = False)
    Debug.Assert (bCondition)
End Sub

'判断x是否介于[a, b]之间
Public Function Between(x, a, b) As Boolean
    Between = ((a <= x) And (x <= b))
End Function

'输入：1~256
Function CellCol2Str(ByVal C) As String
    Dim n0 As String
    Dim n1 As String
    
    If Not IsNumeric(C) Then
        CellCol2Str = UCase(C)
        Exit Function
    End If
    
    C = C - 1
    AssertEx Between(C, 0, 255)
    n0 = Chr((C Mod 26) + Asc("A"))
    C = C \ 26
    If C > 0 Then n1 = Chr(C + Asc("A") - 1)
    
    CellCol2Str = n1 & n0
End Function

'返回值：[1, 256]
Function CellCol2Int(C) As Long
    If IsNumeric(C) Then
        CellCol2Int = Val(C)
        Exit Function
    End If

    C = UCase(C)
    
    Dim d0 As Long
    Dim d1 As Long
    
    If Len(C) = 1 Then
        d0 = Asc(Left(C, 1)) - Asc("A")
    ElseIf Len(C) = 2 Then
        d1 = Asc(Left(C, 1)) - Asc("A") + 1
        d0 = Asc(Mid(C, 2)) - Asc("A")
    End If
    
    CellCol2Int = (d1 * 26 + d0) + 1
End Function

Public Sub FreeMem(v)
    If IsArray(v) Then Erase v
End Sub

Public Function MakeFileName(ByVal strPath, ByVal strName) As String
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    Call CfmDirExist(strPath)
    MakeFileName = strPath & strName
End Function

Public Function CfmDirExist(ByVal strPath As String) As Boolean
    Dim nPos As Long
    Dim strSubDir As String
    Dim ALL_DIR_ATTR As VbFileAttribute
    
    ALL_DIR_ATTR = vbHidden + vbNormal + vbReadOnly + vbSystem + vbDirectory
    
    strPath = Trim(strPath)
    If Right(strPath, 1) = "\" Then strPath = Left(strPath, Len(strPath) - 1)
    If Right(strPath, 1) = "$" Then strPath = strPath & "\"
    If Right(strPath, 1) = ":" Then strPath = strPath & "\"

    On Error GoTo ErrExit
    If Dir(strPath, ALL_DIR_ATTR) = "" Then
        nPos = IIf((Left(strPath, 2) = "\\") And (InStr(1, strPath, "$\") > 0), InStr(1, strPath, "$\") + 1, 3)
        Do
            nPos = InStr(nPos + 1, strPath, "\")
            If nPos = 0 Then Exit Do
            strSubDir = Left(strPath, nPos - 1)
            If Dir(strSubDir, ALL_DIR_ATTR) = "" Then MkDir strSubDir
        Loop
        MkDir strPath
    End If
    
ErrExit:
    CfmDirExist = (Err = 0)
End Function

Public Function CloneWorkBook(wbFrom As Workbook, ByVal strSaveFile As String) As Workbook
    Dim wbCopy As Workbook
    
    On Error Resume Next
    wbFrom.SaveCopyAs strSaveFile
    If Err <> 0 Then Exit Function
    
    Set wbCopy = Workbooks.Open(strSaveFile)
    If Err <> 0 Then Exit Function
    If wbCopy Is Nothing Then Exit Function
    
    wbCopy.RunAutoMacros xlAutoOpen
    Set CloneWorkBook = wbCopy
End Function

Public Sub ShowSheet(sh As Worksheet, bShow As Boolean)
    Dim eVis As XlSheetVisibility
    Dim bPreSaved As Boolean
    
    eVis = IIf(bShow, xlSheetVisible, xlSheetHidden)
    If (sh.Visible <> eVis) Then
        bPreSaved = ThisWorkbook.Saved
        sh.Visible = eVis
        ThisWorkbook.Saved = bPreSaved
    End If
    If bShow Then sh.Activate
End Sub

Public Sub writeSheet(sh As Worksheet, r As Long, nFromCol As Long, ParamArray cellVal() As Variant)
    Dim C As Long
    
    For C = 0 To UBound(cellVal)
        SetCell sh, r, C + nFromCol, cellVal(C)
    Next C
End Sub

'检查是否是空目录
Public Function IsBlankPath(ByVal strPath As String) As Boolean
    Dim strDir As String
    
    IsBlankPath = True
    
    strDir = Dir(strPath & "\*.*", vbDirectory + vbHidden + vbNormal + vbReadOnly + vbSystem)
    Do While strDir <> ""
        If (strDir <> ".") And (strDir <> "..") Then
            IsBlankPath = False
            Exit Do
        End If
        strDir = Dir()
    Loop
End Function

Public Function IsValidPath(ByVal strPath) As Boolean
    Dim strDrv As String
    Dim pos As Long
    Dim strInvalidCh As String
    Dim i As Long
    
    If strPath = "" Then Exit Function
    strInvalidCh = "/:*?<>|" & """"
    
    On Error Resume Next
    If Dir(strPath, vbDirectory) <> "" Then
        IsValidPath = True
        Exit Function
    End If
    Err.Clear
    
    pos = InStr(1, strPath, ":\")
    If pos > 0 Then
        strDrv = Left(strPath, pos + 1)
        If Dir(strDrv, vbDirectory) = "" Then Exit Function
        If Err <> 0 Then Exit Function
        
        strPath = Mid(strPath, pos + 2)
        For i = 1 To Len(strInvalidCh)
            If InStr(1, strPath, Mid(strInvalidCh, i, 1)) > 0 Then
                IsValidPath = False
                Exit Function
            End If
        Next i
        IsValidPath = True
    End If
End Function
'将平台给M2000接口文档中的测量指标拷贝到给M2000的接口文档中
Public Function CopyRosaM2000(shtSrc As Worksheet, shtDst As Worksheet, dstStartRow As Long, objNum As Long, srcStartRow As Long)

    InsertSheetRow shtDst, dstStartRow, objNum
    MakeRange(shtSrc, srcStartRow, "A", srcStartRow + objNum - 1).EntireRow.Copy
    MakeRange(shtDst, dstStartRow, "A", dstStartRow + objNum - 1).PasteSpecial
    
End Function
Public Function CopySheet(shtFrom As Worksheet, shtTo As Worksheet)
    shtTo.UsedRange.EntireRow.Delete
    
    shtFrom.UsedRange.Copy
    shtTo.Paste
    
    Dim r, C
    
    '调整行高、列宽
    For C = 1 To shtFrom.UsedRange.columns.count
        SetColWidth shtTo, C, GetColWidth(shtFrom, C)
    Next C
    
    shtTo.name = shtFrom.name
End Function

Public Function CompileMessage(ByVal szString As Long) As Boolean
    CompileMessage = True
End Function
     
Public Function CompileStatus(ByVal szString As Long) As Boolean
    CompileStatus = True
End Function

Public Function ColString2Num(ByVal colString As String) As Long
    ColString2Num = Asc(UCase(colString)) - Asc("A") + 1
End Function


Sub DisplayMessageOnStatusbar()
    Application.DisplayStatusBar = True '显示状态栏
    Application.StatusBar = "Running,please wait......" '状态栏显示信息
    Application.Cursor = xlWait
End Sub

Public Sub DisplayMessageOnStatusbaring(Num As Long)
    Application.StatusBar = "Running,please wait......Finish " & Num & "%!" '状态栏显示信息
End Sub

Public Sub EndDisplayMessageOnStatusbar()
    Application.Cursor = xlDefault
    Application.StatusBar = "Finished."  '状态栏显示信息
End Sub

Public Sub ReturnStatusbaring()
    Application.StatusBar = "Ready." '状态栏恢复正常
End Sub

Public Sub ClosingProStatusbaring()
    Application.StatusBar = getResByKey("ClosingAndWait") '状态栏恢复正常
End Sub

'装载用于添加「Base Station Transport Data」页「*Site Template」列侯选值的窗体
Sub addTemplate()
    If NodeBCommon.isNodeBComm And Not isCellExist Then
        MsgBox getResByKey("NotSupportAddTemplate"), vbExclamation, getResByKey("Warning")
    Else
        Load TemplateForm
        TemplateForm.Show
    End If
End Sub

Sub addIPRoute()
    Load IPRouteForm
    IPRouteForm.Show
End Sub

Sub addHyperlinks()
    Load HyperlinksForm
    HyperlinksForm.Show
End Sub

'获得垂直合并的组名代码
Public Function getVerticalGroupName(ByRef ws As Worksheet, ByRef rowNumber As Long, ByRef columnLetter As String, _
    ByRef groupStartRow As Long, ByRef groupEndRow As Long) As String
    Dim cellValue As String
    cellValue = ws.range(columnLetter & rowNumber).value
    Dim k As Long
    If cellValue = "" Then
        For k = rowNumber To 1 Step -1
            cellValue = ws.range(columnLetter & k).value
            If cellValue <> "" Then
                getVerticalGroupName = cellValue
                groupStartRow = k
                groupEndRow = getEndRowNumer(ws, columnLetter, k)
                Exit Function
            End If
        Next k
    Else
        getVerticalGroupName = cellValue
        groupStartRow = rowNumber
        groupEndRow = getEndRowNumer(ws, columnLetter, k)
    End If
End Function
'获得垂直合并的组结束行数代码
Public Function getEndRowNumer(ByRef ws As Worksheet, ByRef columnLetter As String, ByRef startRowNumber As Long) As Long
    Dim k As Long
    Dim maxRowNumber As Long
    maxRowNumber = ws.UsedRange.rows.count 'ws.Range("A2").End(xlToRight).Column
    For k = startRowNumber + 1 To maxRowNumber
        If ws.range(columnLetter & k).value <> "" Then
            getEndRowNumer = k - 1
            Exit Function
        End If
    Next k
    getEndRowNumer = maxRowNumber
End Function

'从指定sheet页的指定行，查找指定列，返回列号,通过attrName和MocName获取到
Public Function getColNum(sheetName As String, recordRow As Long, attrName As String, mocName As String) As Long
    On Error GoTo ErrorHandler
    Dim colName As String
    Dim grpName As String
    
    Dim flag As Boolean
    Dim mappingDef As Worksheet
    Dim ws As Worksheet
    If innerPositionMgr Is Nothing Then loadInnerPositions

    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    flag = False
    getColNum = -1
    
    Dim firstAddr As String
    Dim targetRange As range
    With mappingDef
        Set targetRange = .columns(innerPositionMgr.mappingDef_attrNameColNo).Find(attrName, lookat:=xlWhole, LookIn:=xlValues)
        If Not targetRange Is Nothing Then
            firstAddr = targetRange.address
            Do
                If UCase(targetRange.Offset(0, innerPositionMgr.mappingDef_mocNameColNo - innerPositionMgr.mappingDef_attrNameColNo).value) = UCase(mocName) _
                    And targetRange.Offset(0, innerPositionMgr.mappingDef_shtNameColNo - innerPositionMgr.mappingDef_attrNameColNo).value = sheetName Then
                        colName = .Cells(targetRange.row, innerPositionMgr.mappingDef_colNameColNo).value
                        grpName = .Cells(targetRange.row, innerPositionMgr.mappingDef_grpNameColNo).value
                        flag = True
                        Exit Do
                End If
                Set targetRange = .columns(innerPositionMgr.mappingDef_attrNameColNo).FindNext(targetRange)
            Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
        End If
    End With
    
    Set targetRange = Nothing
    firstAddr = ""
    If flag = True Then
        Set ws = ThisWorkbook.Worksheets(sheetName)
        With ws.rows(recordRow)
            Set targetRange = .Find(colName, lookat:=xlWhole, LookIn:=xlValues)
            If Not targetRange Is Nothing Then
                firstAddr = targetRange.address
                Do
                    If get_GroupName(sheetName, targetRange.column) = grpName Then
                        getColNum = targetRange.column
                        Exit Do
                    End If
                    Set targetRange = .FindNext(targetRange)
                Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
            End If
        End With
    End If
    
    Exit Function
ErrorHandler:
    Debug.Print "some exception in getColNum, " & Err.Description
End Function

Public Function GetMainSheetName() As String
    On Error Resume Next
    If innerPositionMgr Is Nothing Then loadInnerPositions
    
    Dim targetRange As range
    Set targetRange = ThisWorkbook.Worksheets("SHEET DEF").columns(innerPositionMgr.sheetDef_shtTypeColNo).Find("MAIN", lookat:=xlWhole, LookIn:=xlValues)
    If Not targetRange Is Nothing Then GetMainSheetName = targetRange.Offset(0, innerPositionMgr.sheetDef_shtNameColNo - innerPositionMgr.sheetDef_shtTypeColNo).value
End Function

Function GetCommonSheetName() As String
         On Error Resume Next
        Dim name As String
        Dim rowNum As Long
        Dim sheetDef As Worksheet
        Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
        
        For rowNum = 1 To sheetDef.range("a65536").End(xlUp).row
            If sheetDef.Cells(rowNum, 2).value = "COMMON" Then
                name = sheetDef.Cells(rowNum, 1).value
                Exit For
            End If
        Next
        GetCommonSheetName = name
End Function

'从普通页取得Group name
Public Function get_GroupName(sheetName As String, column As Long) As String
    Dim index As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    For index = column To 1 Step -1
        If ws.Cells(1, index).value <> "" Then
            get_GroupName = ws.Cells(1, index).value
            Exit Function
        End If
    Next
    get_GroupName = ""
End Function

'从普通页取得Colum name
Public Function get_ColumnName(ByVal sheetName As String, column As Long) As String
        Dim index As Long
        get_ColumnName = ThisWorkbook.Worksheets(sheetName).Cells(2, column)
End Function

Public Function isNum(contedType As String) As Boolean
    If contedType = "Integer" Or contedType = "UInteger" Then
        isNum = True
    Else
        isNum = False
    End If
End Function

Public Sub clearXLGray()
    Dim index, cloumIndex, commIndex, commCloumIndex As Long
    Dim worksh, sheetDef As New Worksheet
    If innerPositionMgr Is Nothing Then loadInnerPositions
    
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
    For index = 2 To sheetDef.range("a65536").End(xlUp).row
        Set worksh = ThisWorkbook.Worksheets(sheetDef.Cells(index, innerPositionMgr.sheetDef_shtNameColNo).value)
        If sheetDef.Cells(index, innerPositionMgr.sheetDef_shtTypeColNo) = "COMMON" Then
            For commIndex = 1 To worksh.range("a65536").End(xlUp).row
                For commCloumIndex = 1 To worksh.range("IV" + CStr(commIndex)).End(xlToLeft).column
                    If worksh.Cells(commIndex, commCloumIndex).Interior.colorIndex = 16 And _
                         worksh.Cells(commIndex, commCloumIndex).Interior.Pattern = xlGray16 Then
                            worksh.Cells(commIndex, commCloumIndex).Interior.colorIndex = xlNone
                            worksh.Cells(commIndex, commCloumIndex).Interior.Pattern = xlNone
                            'worksh.Cells(commIndex, commCloumIndex).Validation.ShowInput = True
                    End If
                Next
            Next
        ElseIf "Pattern" = sheetDef.Cells(index, innerPositionMgr.sheetDef_shtTypeColNo) Then
                
        Else
            For cloumIndex = 1 To worksh.range("IV" + CStr(3)).End(xlToLeft).column
                If worksh.Cells(3, cloumIndex).Interior.colorIndex = 16 And _
                    worksh.Cells(3, cloumIndex).Interior.Pattern = xlGray16 Then
                    worksh.Cells(3, cloumIndex).Interior.colorIndex = xlNone
                    worksh.Cells(3, cloumIndex).Interior.Pattern = xlNone
                    'worksh.Cells(3, cloumIndex).Validation.ShowInput = True
                End If
            Next
        End If
    Next
    Application.DisplayAlerts = False
    ThisWorkbook.Save
End Sub

Public Function isPatternSheet(sheetName As String) As Boolean
    Dim m_rowNum As Long
    If innerPositionMgr Is Nothing Then loadInnerPositions
    
    Dim sheetDef As Worksheet
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
    For m_rowNum = 2 To sheetDef.range("a65536").End(xlUp).row
        If sheetName = sheetDef.Cells(m_rowNum, innerPositionMgr.sheetDef_shtNameColNo).value Then
            If sheetDef.Cells(m_rowNum, innerPositionMgr.sheetDef_shtTypeColNo).value = "Pattern" Then
                isPatternSheet = True
            Else
                isPatternSheet = False
            End If
            Exit For
        End If
    Next
End Function

Sub clearStyles()
        Dim s As Style
        For Each s In ThisWorkbook.Styles
            If Not s.BuiltIn Then
                'Debug.Print s.Name
                Debug.Print s.name
                s.Delete '可以用来删除非内置样式
            End If
        Next
End Sub

Sub deleteWorkbookNames()
    Dim n As name
    For Each n In ThisWorkbook.Names
       Debug.Print n.index
        n.Delete
    Next
End Sub

'包含某个页签代码
Public Function containsASheet(ByRef wb As Workbook, ByRef sheetName As String, Optional ByRef ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    containsASheet = True
    Set ws = wb.Worksheets(sheetName)
    Exit Function
ErrorHandler:
    containsASheet = False
End Function

Public Function findCertainValRowNumberByTwoKeys(ByRef ws As Worksheet, ByVal columnLetter1 As String, ByRef cellVal1 As String, _
    ByVal columnLetter2 As String, ByRef cellVal2 As String, Optional ByVal startRow As Long = 1)
    
    Dim currentCellVal1 As String
    Dim currentCellVal2 As String
    Dim maxRowNumber As Long, k As Long
    maxRowNumber = ws.UsedRange.rows.count
    findCertainValRowNumberByTwoKeys = -1
    For k = startRow To maxRowNumber
        currentCellVal1 = ws.range(columnLetter1 & k).value
        currentCellVal2 = ws.range(columnLetter2 & k).value
        If currentCellVal1 = cellVal1 And currentCellVal2 = cellVal2 Then
            findCertainValRowNumberByTwoKeys = k
            Exit For
        End If
    Next
End Function

Public Sub setHyperlinkRangeFont(ByRef certainRange As range)
    With certainRange.Font
        .name = "Arial"
        .Size = 10
    End With
End Sub

'某行是否为空代码
Public Function rowIsBlank(ByRef ws As Worksheet, ByRef rowNumber As Long) As Boolean
    If Application.WorksheetFunction.CountBlank(ws.range("A" & rowNumber & ":IV" & rowNumber)) = 256 Then
        rowIsBlank = True
    Else
        rowIsBlank = False
    End If
End Function

Sub changeAlerts(ByRef flag As Boolean)
    Application.EnableEvents = flag
    Application.DisplayAlerts = flag
    Application.ScreenUpdating = flag
End Sub

Function is_Site(columnName As String) As Boolean
    is_Site = False
    If (columnName = getResByKey("*NODEB_NAME") Or columnName = getResByKey("*BTS_NAME") Or _
        columnName = getResByKey("*BASESTATION_NAME") Or columnName = getResByKey("*ENODEB_NAME") Or columnName = getResByKey("*USU_NAME") Or _
        columnName = getResByKey("*NBBSName") Or columnName = getResByKey("*ICSNAME") Or columnName = getResByKey("*eLTEName") Or _
        columnName = getResByKey("*RFANAME") Or columnName = getResByKey("*ENODEBEQMNAME") Or _
        columnName = getResByKey("*NRNAME") Or columnName = getResByKey("*NRNAME") Or columnName = getResByKey("*DSAName") _
        ) Then
        is_Site = True
    End If
End Function

Function is_Controller(columnName As String) As Boolean
    is_Controller = False
    If (columnName = getResByKey("*RNCName") Or columnName = getResByKey("*BSCName")) Then
        is_Controller = True
    End If
End Function

Function isOperationWs(ByRef ws As Worksheet) As Boolean
    isOperationWs = False

    If operationColNum(ws) = -1 Then Exit Function
    
    isOperationWs = True
End Function

Public Function getNBIOTFlag() As Boolean
    On Error GoTo ErrorHandler
    Dim mocColNum As Long, headRange As range, prbRange As range
    Set headRange = Worksheets("MAPPING DEF").range("A1:X1").Find(what:="Attribute Name", lookat:=xlWhole)
    If Not headRange Is Nothing Then
         mocColNum = headRange.column
    End If

    Dim rowEnd As Long, colstr As String
    colstr = getColStr(mocColNum)
    Set prbRange = Worksheets("MAPPING DEF").range(colstr & ":" & colstr).Find(what:="NbCellFlag", lookat:=xlWhole)
    If Not prbRange Is Nothing Then
        getNBIOTFlag = True
        Exit Function
    End If

ErrorHandler:
    getNBIOTFlag = False
End Function

Public Function getSheetDefNameColNum(ByRef titleName As String) As Long
    Dim sheetDef As Worksheet
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
    Dim maxColNum As Long, index As Long, lastColNum As Long
    maxColNum = sheetDef.range("IV1").End(xlToLeft).column
    For index = 1 To maxColNum
         If UCase(sheetDef.Cells(1, index).value) = UCase(titleName) Then
             getSheetDefNameColNum = index
             Exit Function
         End If
    Next
    getSheetDefNameColNum = -1
    lastColNum = -1
    If titleName = BaseSheetName_Name And lastColNum = -1 Then
        getSheetDefNameColNum = 6
    End If
    Exit Function
ErrorHandler:
    getSheetDefNameColNum = -1
End Function

Public Function getSrcSheetDefNameColNum(ByRef ws As Workbook, ByRef titleName As String) As Long
    Dim sheetDef As Worksheet
    Set sheetDef = ws.Worksheets("SHEET DEF")
    Dim maxColNum As Long, index As Long, lastColNum As Long
    maxColNum = sheetDef.range("IV1").End(xlToLeft).column
    For index = 1 To maxColNum
         If UCase(sheetDef.Cells(1, index).value) = UCase(titleName) Then
             getSrcSheetDefNameColNum = index
             Exit Function
         End If
    Next
    getSrcSheetDefNameColNum = -1
    lastColNum = -1
    If titleName = BaseSheetName_Name And lastColNum = -1 Then
        getSrcSheetDefNameColNum = 6
    End If
    Exit Function
ErrorHandler:
    getSrcSheetDefNameColNum = -1
End Function


Public Function getSrcSheetDefNameColNum4Worksheet(ByRef sheetDef As Worksheet, ByRef titleName As String) As Long
    Dim maxColNum As Long, index As Long, lastColNum As Long
    maxColNum = sheetDef.range("IV1").End(xlToLeft).column
    For index = 1 To maxColNum
         If UCase(sheetDef.Cells(1, index).value) = UCase(titleName) Then
             getSrcSheetDefNameColNum4Worksheet = index
             Exit Function
         End If
    Next
    getSrcSheetDefNameColNum4Worksheet = -1
    lastColNum = -1
    If titleName = BaseSheetName_Name And lastColNum = -1 Then
        getSrcSheetDefNameColNum4Worksheet = 6
    End If
    Exit Function
ErrorHandler:
    getSrcSheetDefNameColNum4Worksheet = -1
End Function

Public Function getUsedColumnCount(ByRef sheet As Worksheet, Optional ByRef attrNamerowIndex As Long = 2)
    getUsedColumnCount = sheet.range("IV" & attrNamerowIndex).End(xlToLeft).column
End Function

Public Function getUsedRowCount(ByRef sheet As Worksheet, Optional ByRef attrNamerowIndex As Long = 2)
    Dim colCount As Long
    colCount = getUsedColumnCount(sheet, attrNamerowIndex)
    getUsedRowCount = sheet.columns("A:" & getColStr(colCount)).Find(what:="*", searchorder:=xlByRows, searchdirection:=xlPrevious).row
End Function



Public Function findAttrName(ByRef attrName As String) As Boolean
    On Error GoTo ErrorHandler
    If Len(Trim(attrName)) < 1 Then
        findAttrName = False
        Exit Function
    End If
    
    Dim mocColNum As Long, headRange As range, prbRange As range
    Set headRange = Worksheets("MAPPING DEF").range("A1:X1").Find(what:="Column Name", lookat:=xlWhole)
    If Not headRange Is Nothing Then
         mocColNum = headRange.column
    End If
    
    Dim rowEnd As Long, colstr As String
    colstr = getColStr(mocColNum)
    Set prbRange = Worksheets("MAPPING DEF").range(colstr & ":" & colstr).Find(what:=attrName, lookat:=xlWhole)
    If Not prbRange Is Nothing Then
        findAttrName = True
        Exit Function
    End If
ErrorHandler:
findAttrName = False
End Function

Public Function findGroupName(ByRef groupName As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Len(Trim(groupName)) < 1 Then
        findGroupName = False
        Exit Function
    End If
    
    Dim mocColNum As Long, headRange As range, prbRange As range
    Set headRange = Worksheets("MAPPING DEF").range("A1:X1").Find(what:="Group Name", lookat:=xlWhole)
    If Not headRange Is Nothing Then
         mocColNum = headRange.column
    End If
    
    Dim rowEnd As Long, colstr As String
    colstr = getColStr(mocColNum)
    Set prbRange = Worksheets("MAPPING DEF").range(colstr & ":" & colstr).Find(what:=groupName, lookat:=xlWhole)
    If Not prbRange Is Nothing Then
        findGroupName = True
        Exit Function
    End If
ErrorHandler:
findGroupName = False
End Function

Public Function replaceGenel(ByRef strName As String) As String
 On Error GoTo ErrorHandler
 If InStr(strName, "*") > 0 Then
  replaceGenel = Replace(strName, "*", "~*")
  Exit Function
 End If
ErrorHandler:
replaceGenel = strName
End Function

Public Function isAttrRow_IUB(sht As Worksheet, ByVal rowIdx As Integer) As Boolean
    On Error GoTo ErrorHandler
    isAttrRow_IUB = True
    If sht.Cells(rowIdx, 1) <> "" Then Exit Function
    
    isAttrRow_IUB = False
    
    Exit Function
ErrorHandler:
    Debug.Print "some exception in isAttrRow_IUB, " & Err.Description
    isAttrRow_IUB = False
End Function

Public Function attrNameColNumInSpecialDef(ByRef sheet As Worksheet, ByRef attrName As String, Optional ByRef attrNamerowIndex As Long = 1) As Long
    On Error GoTo ErrorHandler
    attrNameColNumInSpecialDef = -1
    
    Dim targetRange As range
    Set targetRange = sheet.rows(attrNamerowIndex).Find(Trim(attrName), LookIn:=xlValues, lookat:=xlWhole)
    If Not targetRange Is Nothing Then attrNameColNumInSpecialDef = targetRange.column
    Exit Function
ErrorHandler:
    attrNameColNumInSpecialDef = -1
End Function

Public Function existsASheet(shtName As String) As Boolean
On Error GoTo ErrorHandler:
    existsASheet = True
    Dim sht As Worksheet
    Set sht = ThisWorkbook.Worksheets(shtName)
    Exit Function
ErrorHandler:
    existsASheet = False
End Function

Public Function IsSheetExist(sheetName As String) As Boolean
    Dim SheetNum, SheetCount As Long 'SheetCount每个原始数据文件的Sheet页总数
    SheetCount = ActiveWorkbook.Worksheets.count   '共有几个Sheet页
    For SheetNum = 1 To SheetCount
        If UCase(Worksheets(SheetNum).name) = UCase(sheetName) Then
            IsSheetExist = True
            Exit Function
        End If
    Next SheetNum
    IsSheetExist = False
End Function

Public Function GetBluePrintSheetName() As String '当前只支持一个
    GetBluePrintSheetName = ""
    
    Dim SheetNum, SheetCount As Long
    SheetCount = ActiveWorkbook.Worksheets.count
    For SheetNum = 1 To SheetCount
        If Worksheets(SheetNum).Tab.colorIndex = BluePrintSheetColor Then
            GetBluePrintSheetName = Worksheets(SheetNum).name
            Exit Function
        End If
    Next SheetNum
End Function

Public Sub InitTemplateVersion()
    bIsEng = getResByKey("Cover") = "Cover"
End Sub

Public Function isMultiVersionWb() As Boolean
    isMultiVersionWb = False
    If existsASheet(getResByKey("ModelDiffSht")) Then
        isMultiVersionWb = True
    End If
End Function

Public Function collectionJoin(coll As Collection, Optional delimiter As String = ",") As String
    On Error GoTo ErrorHandler
    collectionJoin = ""
    Dim del As String
    del = ""
    
    Dim it As Variant
    For Each it In coll
        collectionJoin = collectionJoin & del & CStr(it)
        del = delimiter
    Next
    Exit Function
ErrorHandler:
    Debug.Print "some exception in collectionJoin, " & Err.Description
    collectionJoin = ""
End Function

Public Function IsBluePrintSheetName(sheetName As String) As Boolean
    IsBluePrintSheetName = (Sheets(sheetName).Tab.colorIndex = BluePrintSheetColor)
End Function

Public Function getIndirectListValue(sheet As Worksheet, ByVal colNum As Long, rawListValue As String) As String
On Error GoTo ErrorHandler
    Dim groupName As String
    Dim columnName As String
    Dim valideDef As CValideDef
    
    Call getGrpAndColName(sheet, colNum, groupName, columnName)
    
    Set valideDef = initDefaultDataSub.getInnerValideDef(sheet.name + "," + groupName + "," + columnName)
    
    If valideDef Is Nothing Then
        Set valideDef = addInnerValideDef(sheet.name, groupName, columnName, rawListValue)
    Else
        Call modiflyInnerValideDef(sheet.name, groupName, columnName, rawListValue, valideDef)
    End If
    
    getIndirectListValue = valideDef.getValidedef
    
    Exit Function
ErrorHandler:
    Debug.Print "some exception in getInderectListValue, " & Err.Description
End Function

Private Sub getGrpAndColName(sht As Worksheet, ByVal colNum As Long, grpName As String, colName As String)
    Dim col As Long
    With sht
        colName = .Cells(listShtTitleRow, colNum).value
        For col = colNum To 1 Step -1
            If .Cells(1, col).value <> "" Then
                grpName = .Cells(1, col).value
                Exit For
            End If
        Next
    End With
End Sub

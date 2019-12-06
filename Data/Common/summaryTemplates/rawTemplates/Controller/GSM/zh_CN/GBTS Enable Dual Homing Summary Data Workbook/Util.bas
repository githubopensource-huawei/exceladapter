Attribute VB_Name = "Util"
Option Explicit

Global Const g_strInvalidChar4Sql As String = "'"
Global Const g_strInvalidChar4PureFileName  As String = "\/:*?<>|"""
Global Const g_strInvalidChar4Path          As String = "/*?<>|"""

Private Const SW_SHOWNORMAL = 1
Public HasHistoryData As Boolean


'调用方法：GetFileName("话统脚本, *.sql", "Open")
Public Function GetFileName(ByVal strFilter, ByVal strTitle, ByVal bMulSel As Boolean, vFileName As Variant) As Boolean
    
    Dim vRsp
    Dim i As Long
    
    GetFileName = False
    vRsp = Application.GetOpenFilename(FileFilter:=strFilter, Title:=strTitle, MultiSelect:=bMulSel)
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
            SysErr "File path" & vbCrLf & strPath & vbCrLf & "invalid！"
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

Public Function GetCell(shtX As Worksheet, ByVal r, ByVal c)
    GetCell = shtX.Cells(r, c)
End Function

'写入：True；未写入：False
Public Function SetCell(shtX As Worksheet, ByVal r, ByVal c, ByVal strCellVal) As Boolean
    Dim strOld As String
    
    strOld = GetCell(shtX, r, c)
    If strOld <> CStr(strCellVal) Then
        shtX.Cells(r, c) = strCellVal
        SetCell = True
    End If
End Function

Public Function IsNullRow(shtX As Worksheet, ByVal r As Long, Optional FromCol, Optional ToCol) As Boolean
    Dim nFromCol As Long
    Dim nToCol As Long
    
    nFromCol = IIf(IsMissing(FromCol), 1, CellCol2Int(FromCol))
    nToCol = IIf(IsMissing(ToCol), 3, CellCol2Int(ToCol))
    
    Dim c As Long
    Dim strCell As String
    
    IsNullRow = False
    For c = nFromCol To nToCol
        strCell = GetCell(shtX, r, c)
        If LeftMostMatch(strCell, "//") Then Exit For '该行为注释行，视为空行
        If strCell <> "" Then Exit Function
    Next c
    IsNullRow = True
End Function

Public Sub SetNullRow(shtX As Worksheet, ByVal r As Long, Optional FromCol, Optional ToCol)
    
    If IsMissing(FromCol) Then FromCol = "A"
    If IsMissing(ToCol) Then FromCol = "C"
    
    If Not IsNumeric(FromCol) Then FromCol = Asc(FromCol) - Asc("A") + 1
    If Not IsNumeric(ToCol) Then ToCol = Asc(ToCol) - Asc("A") + 1
    
    Dim c As Long
    
    For c = FromCol To ToCol
        Call SetCell(shtX, r, c, "")
    Next c

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

Public Function FormatCellAddr(ByVal r, ByVal c)
    FormatCellAddr = "$" & CellCol2Str(c) & "$" & r
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
    MakeRange(shtX, row, "A", row + Num - 1).EntireRow.Insert
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
    
    If Not IsMissing(strFile) Then strWri = ThisWorkbook.path & "\" & strFile
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
Function CellCol2Str(ByVal c) As String
    Dim n0 As String
    Dim n1 As String
    
    If Not IsNumeric(c) Then
        CellCol2Str = UCase(c)
        Exit Function
    End If
    
    c = c - 1
    AssertEx Between(c, 0, 255)
    n0 = Chr((c Mod 26) + Asc("A"))
    c = c \ 26
    If c > 0 Then n1 = Chr(c + Asc("A") - 1)
    
    CellCol2Str = n1 & n0
End Function

'返回值：[1, 256]
Function CellCol2Int(c) As Long
    If IsNumeric(c) Then
        CellCol2Int = Val(c)
        Exit Function
    End If

    c = UCase(c)
    
    Dim d0 As Long
    Dim d1 As Long
    
    If Len(c) = 1 Then
        d0 = Asc(Left(c, 1)) - Asc("A")
    ElseIf Len(c) = 2 Then
        d1 = Asc(Left(c, 1)) - Asc("A") + 1
        d0 = Asc(Mid(c, 2)) - Asc("A")
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
    Dim c As Long
    
    For c = 0 To UBound(cellVal)
        SetCell sh, r, c + nFromCol, cellVal(c)
    Next c
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
    
    Dim r, c
    
    '调整行高、列宽
    For c = 1 To shtFrom.UsedRange.columns.count
        SetColWidth shtTo, c, GetColWidth(shtFrom, c)
    Next c
    
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

Sub copyENodeBNameId()
    Dim sheetCount As Long
    Dim iIndext As Long
    Dim jIndext As Long
    Dim iRow As Long
    Dim rowCount As Long
    Dim bFind As Boolean
    Dim sheetCur  As Worksheet
    Dim sheetEquip  As Worksheet
    sheetCount = ThisWorkbook.Sheets.count
    Set sheetEquip = ThisWorkbook.Worksheets("Equipment")
   
    
    '遍历所有sheet
    For iIndext = 4 To sheetCount
        Set sheetCur = ThisWorkbook.Worksheets(iIndext)
        If sheetCur.Cells(3, 1).value = "eNodeBName" Then
            '遍历所有enodeb name
            iRow = 4
            While sheetEquip.Cells(iRow, 1).value <> ""
                bFind = False
                rowCount = 1
                While sheetCur.Cells(rowCount, 1).value <> ""
                    rowCount = rowCount + 1
                Wend
                For jIndext = 4 To rowCount - 1
                     
                    '找到enodb name
                    If sheetEquip.Cells(iRow, 1).value & "" = sheetCur.Cells(jIndext, 1).value & "" Then
                        bFind = True
                    End If
                Next
                If bFind = False Then
                    sheetCur.Cells(rowCount, 1).value = sheetEquip.Cells(iRow, 1).value
                    sheetCur.Cells(rowCount, 2).value = sheetEquip.Cells(iRow, 2).value
                End If
                iRow = iRow + 1
            Wend
        End If
    Next
           
End Sub

Sub DisplayMessageOnStatusbar()
    Application.DisplayStatusBar = True '显示状态栏
    Application.StatusBar = "Running,please wait......" '状态栏显示信息

End Sub

Public Sub DisplayMessageOnStatusbaring(Num As Long)
    Application.StatusBar = "Running,please wait......Finish " & Num & "%!" '状态栏显示信息
End Sub

Public Sub EndDisplayMessageOnStatusbar()
    Application.StatusBar = "Finished."  '状态栏显示信息
End Sub

Public Sub ReturnStatusbaring()
    Application.StatusBar = "Ready." '状态栏恢复正常
End Sub


Public Sub ClosingProStatusbaring()
    Application.StatusBar = "Closing the application，please wait......" '状态栏恢复正常
End Sub
'从指定sheet页的指定行，查找指定列，返回列号,通过attrName和MocName获取到
Public Function getColNum(sheetName As String, recordRow As Long, attrName As String, mocName As String) As Long
    On Error Resume Next
    Dim m_colNum As Long
    Dim m_rowNum As Long
    Dim colName As String
    Dim colGroupName As String
    
    Dim flag As Boolean
    Dim mappingDef As Worksheet
    Dim ws As Worksheet
    
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    flag = False
    getColNum = -1
    For m_rowNum = 2 To mappingDef.range("a1048576").End(xlUp).row
        If UCase(attrName) = UCase(mappingDef.Cells(m_rowNum, 5).value) _
           And UCase(sheetName) = UCase(mappingDef.Cells(m_rowNum, 1).value) _
           And UCase(mocName) = UCase(mappingDef.Cells(m_rowNum, 4).value) Then
            colName = mappingDef.Cells(m_rowNum, 3).value
            colGroupName = mappingDef.Cells(m_rowNum, 2).value
            flag = True
            Exit For
        End If
    Next
    If flag = True Then
        Set ws = ThisWorkbook.Worksheets(sheetName)
        For m_colNum = 1 To ws.range("XFD" + CStr(recordRow)).End(xlToLeft).column
            If get_GroupName(sheetName, m_colNum) = colGroupName Then
                If GetDesStr(colName) = GetDesStr(ws.Cells(recordRow, m_colNum).value) Then
                    getColNum = m_colNum
                    Exit For
                End If
            End If
        Next
    End If
End Function
Public Function GetMainSheetName() As String
       On Error Resume Next
        Dim name As String
        Dim rowNum As Long
        Dim sheetDef As Worksheet
        Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
        For rowNum = 1 To sheetDef.range("a1048576").End(xlUp).row
            If sheetDef.Cells(rowNum, 2).value = "MAIN" Then
                name = sheetDef.Cells(rowNum, 1).value
                Exit For
            End If
        Next
        GetMainSheetName = name
End Function

Function GetCommonSheetName() As String
         On Error Resume Next
        Dim name As String
        Dim rowNum As Long
        Dim sheetDef As Worksheet
        Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
        
        For rowNum = 1 To sheetDef.range("a1048576").End(xlUp).row
            If sheetDef.Cells(rowNum, 2).value = "COMMON" Then
                name = sheetDef.Cells(rowNum, 1).value
                Exit For
            End If
        Next
        GetCommonSheetName = name
End Function

'从普通页取得Group name
Public Function get_GroupName(sheetName As String, colNum As Long) As String
        Dim index As Long
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(sheetName)
        For index = colNum To 1 Step -1
            If Not isEmpty(ws.Cells(1, index).value) Then
                get_GroupName = ws.Cells(1, index).value
                Exit Function
            End If
        Next
        get_GroupName = ""
End Function

'从普通页取得Colum name
Public Function get_ColumnName(ByVal sheetName As String, colNum As Long) As String
        Dim index As Long
        get_ColumnName = ThisWorkbook.Worksheets(sheetName).Cells(2, colNum)
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
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
    For index = 2 To sheetDef.range("a1048576").End(xlUp).row
            Set worksh = ThisWorkbook.Worksheets(sheetDef.Cells(index, 1).value)
            If sheetDef.Cells(index, 2) = "COMMON" Then
                    For commIndex = 1 To worksh.range("a1048576").End(xlUp).row
                            For commCloumIndex = 1 To worksh.range("XFD" + CStr(commIndex)).End(xlToLeft).column
                                If worksh.Cells(commIndex, commCloumIndex).Interior.colorIndex = 16 And _
                                     worksh.Cells(commIndex, commCloumIndex).Interior.Pattern = xlGray16 Then
                                        worksh.Cells(commIndex, commCloumIndex).Interior.colorIndex = xlNone
                                        worksh.Cells(commIndex, commCloumIndex).Interior.Pattern = xlNone
                                        'worksh.Cells(commIndex, commCloumIndex).Validation.ShowInput = True
                                End If
                            Next
                    Next
            ElseIf "Pattern" = sheetDef.Cells(index, 2) Then
                    
            Else
                    For cloumIndex = 1 To worksh.range("XFD" + CStr(3)).End(xlToLeft).column
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
    Dim sheetDef As Worksheet
    Set sheetDef = ThisWorkbook.Worksheets("SHEET DEF")
    For m_rowNum = 2 To sheetDef.range("a1048576").End(xlUp).row
        If sheetName = sheetDef.Cells(m_rowNum, 1).value Then
            If sheetDef.Cells(m_rowNum, 2).value = "Pattern" Then
                isPatternSheet = True
            Else
                isPatternSheet = False
            End If
            Exit For
        End If
    Next
End Function

Function getColStr(ByVal NumVal As Long) As String
    Dim str As String
    Dim strs() As String
    
    If NumVal > 256 Or NumVal < 1 Then
        getColStr = ""
    Else
        str = Cells(NumVal).address
        strs = Split(str, "$", -1)
        getColStr = strs(1)
    End If
End Function

 Function Contains(coll As Collection, key As String) As Boolean
        On Error GoTo NotFound
        Call coll(key)
        Contains = True
        Exit Function
NotFound:
        Contains = False
End Function

'包含某个页签代码
Public Function containsASheet(ByRef wb As Workbook, ByRef sheetName As String) As Boolean
    On Error GoTo ErrorHandler
    containsASheet = True
    Dim sheet As Worksheet
    Set sheet = wb.Worksheets(sheetName)
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

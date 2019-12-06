Attribute VB_Name = "AddLnkModule"
Option Explicit

Private Enum E_Language

    CHN = 1
    ENG = 2

End Enum

'Const Parameters
'Private Const NENAME As String = "NENAME"
'Private Const CELLID As String = "CELLID"
'Private Const SITENAME As String = "SITENAME"
Private Const MERGETROW As Integer = 2
Private Const DISPLAYNAMEROW As Integer = 3
Private Const SHORTNAMEROW As Integer = 4

Private versionEndRow As Integer
Private mergeStartRow As Integer
Private mocCnt As Integer
Private attrStartRow As Integer
Private subFeatureSize As Integer
Private mustGiveArr As Variant
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


Public Sub AddLnk()
    
    'Stop
    Application.ScreenUpdating = False  ' 关闭屏幕更新可加快宏的执行速度
    Application.DisplayAlerts = False

    Call readCfg

    Call SortWorksheets '排序
    
    Call insertRowForTemplateSheet   '为Template页签插入一行

    ' 激活Cover页到第一页
    Dim coverSheet As Worksheet
    Set coverSheet = ThisWorkbook.Worksheets(getResByKey("Home"))
    coverSheet.Move Before:=ActiveWorkbook.Sheets(1)

    Dim helpSheet As Worksheet
    Set helpSheet = ThisWorkbook.Worksheets(getResByKey("Help"))
    helpSheet.Move Before:=ActiveWorkbook.Sheets(1)

    'coverSheet.Tab.ColorIndex = 3 'Set Color
    'coverSheet.Activate

    coverSheet.Activate
    coverSheet.columns(1).ColumnWidth = 20
    coverSheet.columns(2).ColumnWidth = 40
    If subFeatureSize > 0 Then
        Dim i As Integer
        For i = 3 To subFeatureSize + 1
            coverSheet.columns(i).ColumnWidth = 40
        Next i
    End If
    
    If subFeatureSize > 0 Then
        Call setFrameNone4FeaturePkg(2, 4, 2, subFeatureSize) '清除大颗粒表格多余边框

        Call setFrameNone4FeaturePkg(5, 6, 0, subFeatureSize)
        
        Call setFrameNone4FeaturePkg(9, 10, 0, subFeatureSize) '先清边框，再加边框
        
        Call setFrameNone4FeaturePkg(11, 11, 2, subFeatureSize)
        Call setFrameNone4FeaturePkg(12, versionEndRow, 2, subFeatureSize)
        
        Call setSubFeatureFrame(7, 7, subFeatureSize) '子Feature特性加边框
        Call setSubFeatureFrame(8, 8, subFeatureSize)
    
        Call setFrame(2, 2) 'Feature 信息加边框
        Call setFrame(3, 3) 'Feature 信息加边框
        Call setFrame(4, 4) 'Feature 信息加边框
        Call setFrame(11, 11) 'Version信息加边框
        Call setFrame(12, versionEndRow) 'Version信息加边框
    Else
        Call setFrameNone(4, 5) 'Feature 空格清边框
    
        Call setFrameNone(versionEndRow + 1, versionEndRow + 2) 'Feature 空格清边框
        Call setFrame(2, 3) 'Feature 信息加边框
        
        Call setFrame(6, 6) 'Version信息加边框
        Call setFrame(7, versionEndRow) 'Version信息加边框
    End If

    
    

    'Hilight Tip Info Merge
    Dim tileRange As Range
    Set tileRange = Range("A" + CStr(versionEndRow + 2) + ":" + "B" + CStr(versionEndRow + 2) + ":" + "C" + CStr(versionEndRow + 2))
    tileRange.Merge
    tileRange.Borders(xlEdgeLeft).LineStyle = xlNone
    tileRange.Borders(xlEdgeTop).LineStyle = xlNone
    tileRange.Borders(xlEdgeBottom).LineStyle = xlNone
    tileRange.Borders(xlEdgeRight).LineStyle = xlNone
    Dim sheetCount As Long
    sheetCount = Sheets.count

    'Cover页处理--增加到各个Sheet的链接
    Dim sheetIndex As Long
    Dim row As Integer
    row = mergeStartRow
    For sheetIndex = 1 To mocCnt
        Dim lnkStr As String
        Dim dspStr As String
        Dim dspStr1 As String
        lnkStr = "'" + coverSheet.Cells(row, "D").value + "'"
        'coverSheet.Cells(row, "C").Formula = ""   'C D列为临时数据
        dspStr = coverSheet.Cells(row, "B").value
        With coverSheet.Cells(row, "B")
            .value = dspStr
            .Hyperlinks.Add Anchor:=coverSheet.Cells(row, "B"), address:="", SubAddress:=lnkStr + "!A1", TextToDisplay:=dspStr
         End With
         '大颗粒场景
         If subFeatureSize > 0 Then
            With Worksheets(coverSheet.Cells(row, "D").value).Cells(1, "A")
                .value = "Back"
                .Hyperlinks.Add Anchor:=ThisWorkbook.Worksheets(coverSheet.Cells(row, "D").value).Cells(1, "A"), address:="", SubAddress:=coverSheet.name + "!B" + CStr(row), TextToDisplay:=coverSheet.name
            End With
         End If
         
         row = row + 1
    Next sheetIndex
    
    '大颗粒场景增加Template页签的链接，放在MO下面，合并成一个单元格
    If subFeatureSize > 0 Then
        Dim tempRow As Integer
        tempRow = row - 1
        Call setFrameFroMO(tempRow, tempRow)
    
        Dim tempRange As Range
        Set tempRange = Range("A" + CStr(tempRow) + ":" + "B" + CStr(tempRow) + ":" + "C" + CStr(tempRow))
        tempRange.Merge
    End If


    'CoverSHeet Must Give
    Dim mustGiveStartRow As Integer
    mustGiveStartRow = mergeStartRow
    Dim mustGive As String
    Do
        mustGive = coverSheet.Cells(mustGiveStartRow, "F").value
        If mustGive <> "" Then
            coverSheet.Cells(mustGiveStartRow, "B").Interior.ColorIndex = 46
            coverSheet.Cells(mustGiveStartRow, "C").Interior.ColorIndex = 46
            coverSheet.Cells(mustGiveStartRow, "F").Formula = ""
        End If
        mustGiveStartRow = mustGiveStartRow + 1
    Loop Until coverSheet.Cells(mustGiveStartRow, "E").value = ""
    Call mustGivesheet
    If subFeatureSize > 0 Then
        Call executeTemplateBranchControlAll
    End If
    coverSheet.Activate

    '首列NeType信息Merge
    'Dim mergeStartRow As Integer,
    Dim mergeRowCnt As Integer, mergeEndRow As Integer, clearRow As Integer
    Dim startCell As String, endCell As String

    'mergeStartRow = 8
    Call setFrameFroMO(mergeStartRow - 1, mergeStartRow - 1) 'MO信息加边框
    On Error Resume Next  '规避报错
    Do
        'Stop
        mergeRowCnt = CInt(coverSheet.Cells(mergeStartRow, "E").value)
        mergeEndRow = mergeStartRow + mergeRowCnt - 1

        For clearRow = mergeStartRow To mergeEndRow
            Range("D" + CStr(clearRow)).Formula = ""   'C D列为临时数据
            Range("E" + CStr(clearRow)).Formula = ""   'C D列为临时数据
        Next clearRow

        'SetFrame Begin
        startCell = "A" + CStr(mergeStartRow)
        endCell = "C" + CStr(mergeEndRow)
        'Range(startCell + ":" + endCell).Select
        Call setFrameFroMO(mergeStartRow, mergeEndRow)
        'SetFrame End

        endCell = "A" + CStr(mergeEndRow)
        Range(startCell + ":" + endCell).Merge
        mergeStartRow = mergeStartRow + mergeRowCnt

    Loop Until coverSheet.Cells(mergeStartRow, "E").value = ""


    '各个Sheet到Cover页的链接,单特性场景添加链接
    Dim validCol As Integer
    For sheetIndex = 1 To sheetCount
        If Worksheets(sheetIndex).Visible = xlSheetVisible And Worksheets(sheetIndex).name <> getResByKey("Home") And Worksheets(sheetIndex).name <> getResByKey("Help") Then
            If subFeatureSize <= 0 Then
                With Worksheets(sheetIndex).Cells(1, "A")
                    .value = "Back"
                    .Hyperlinks.Add Anchor:=Worksheets(sheetIndex).Cells(1, "A"), address:="", SubAddress:=coverSheet.name + "!A1", TextToDisplay:=coverSheet.name
                End With
            End If

            'Green color for suggested value
            validCol = 0
            Do
                validCol = validCol + 1
            Loop Until Worksheets(sheetIndex).Cells(3, validCol).value = ""
            
            If Worksheets(sheetIndex).name <> getResByKey("PackageCustomTemplate") Then
                Dim suggestRange As Range
                Set suggestRange = Worksheets(sheetIndex).Range(Worksheets(sheetIndex).Cells(5, 1), Worksheets(sheetIndex).Cells(5, validCol - 1))
                suggestRange.Interior.ColorIndex = 43
                Call setComment(suggestRange)
                Call mergeSpecialColumns(sheetIndex, validCol - 1)
                Dim commentRange As Range
                Set commentRange = Worksheets(sheetIndex).Range(Worksheets(sheetIndex).Cells(4, 1), Worksheets(sheetIndex).Cells(4, validCol - 1))
                Call refreshComment(commentRange)
            End If
        End If

    Next sheetIndex
    
    If subFeatureSize > 0 Then
        Dim templateSheet As Worksheet
        Set templateSheet = ThisWorkbook.Worksheets(getResByKey("PackageCustomTemplate"))
        templateSheet.Move Before:=ActiveWorkbook.Sheets(1)
    End If
    coverSheet.Move Before:=ActiveWorkbook.Sheets(1)
    helpSheet.Move Before:=ActiveWorkbook.Sheets(1)
    helpSheet.Activate
    
End Sub


Private Sub mustGivesheet()
'Sheet Must Give
    Dim startIndex As Integer
    Dim mustGiveAttr As String
    startIndex = 6
    Dim sheetInfo1 As Variant
    Dim sheetInfo2 As Variant
    Dim tmpSheet As Worksheet
    Dim mustGiveCol As Integer
    Dim endRow As Long
    Dim cellStr As String
    Dim firstMocPage As Integer
    firstMocPage = 3
    Do
        mustGiveAttr = mustGiveArr(startIndex)
        If mustGiveAttr <> "" Then
            sheetInfo1 = Split(mustGiveAttr, ":")
            sheetInfo2 = Split(sheetInfo1(2), ";")
            Set tmpSheet = ThisWorkbook.Worksheets(sheetInfo1(0))
            endRow = CLng(sheetInfo1(1))
            Dim needHilightSheet As Boolean
            needHilightSheet = False
            
            
            'set the color of must give attribute
            If sheetInfo2(0) <> "" Then
                Dim emptyColCnt As Integer
                emptyColCnt = 0
                needHilightSheet = True
                Do
                    emptyColCnt = emptyColCnt + 1
                Loop Until tmpSheet.Cells(2, emptyColCnt) <> ""

                tmpSheet.Range(tmpSheet.Cells(attrStartRow + 2, 1), tmpSheet.Cells(endRow + 3, emptyColCnt - 1)).Interior.ColorIndex = 46
            Else 'when only hava sitename col
                tmpSheet.Range(tmpSheet.Cells(attrStartRow + 2, 1), tmpSheet.Cells(endRow + 3, 1)).Interior.ColorIndex = 46
            End If

            Dim mocId As Integer
            mocId = 0
            Do
            Dim tmpCol As Integer
            tmpCol = 0
                Do
                    Do
                        tmpCol = tmpCol + 1
                    Loop Until tmpSheet.Cells(attrStartRow, tmpCol).value = sheetInfo2(mocId) Or tmpSheet.Cells(attrStartRow, tmpCol).value = ""
    
                    If tmpSheet.Cells(attrStartRow, tmpCol).value <> "" Then
                        tmpSheet.Range(tmpSheet.Cells(attrStartRow + 2, tmpCol), tmpSheet.Cells(endRow + 3, tmpCol)).Interior.ColorIndex = 46
                        Call refreshMappingDef(sheetInfo1(0), sheetInfo2(mocId))
                    End If
                Loop Until tmpSheet.Cells(attrStartRow, tmpCol).value = ""

                mocId = mocId + 1
            Loop Until sheetInfo2(mocId) = ""

            startIndex = startIndex + 1
            If needHilightSheet Then
                tmpSheet.Tab.ColorIndex = 46
            End If
            tmpSheet.Move Before:=ActiveWorkbook.Sheets(firstMocPage)
        End If
    Loop Until mustGiveArr(startIndex) = ""
End Sub

Private Sub refreshMappingDef(ByVal sheetName As String, ByVal attrName As String)
        Dim index As Long
        Dim sheetDef As Worksheet
        Set sheetDef = ThisWorkbook.Worksheets("MAPPING DEF")
        For index = 2 To sheetDef.Range("a65536").End(xlUp).row
            If sheetDef.Cells(index, 1).value = sheetName And sheetDef.Cells(index, 5).value = attrName Then
                sheetDef.Cells(index, 7) = "true"
                Exit For
            End If
        Next
End Sub
Private Sub SortWorksheets()
'以升序排列工作表
    Dim sCount As Integer, i As Integer, j As Integer
    Application.ScreenUpdating = False
    sCount = Worksheets.count
    If sCount = 1 Then Exit Sub
    For i = 1 To sCount - 1
        If Worksheets(i).Visible = xlSheetVisible Then
            If Worksheets(i).name <> getResByKey("Home") And Worksheets(i).name <> getResByKey("Help") And Worksheets(i).name <> FeatureListSheetName And Worksheets(i).name <> ControllSheetName And Worksheets(i).name <> SheetDefName And Worksheets(i).name <> MappingSheetName And Worksheets(i).name <> getResByKey("PackageCustomTemplate") Then
                With Worksheets(i).Cells(1, "A")
                    .value = "Back"
                    .Hyperlinks.Add Anchor:=Worksheets(i).Cells(1, "A"), address:="", SubAddress:=getResByKey("Home") + "!A1", TextToDisplay:=getResByKey("Home")
                End With
            End If
            For j = i + 1 To sCount
                If Worksheets(j).Visible = xlSheetVisible Then
                    If Worksheets(j).name < Worksheets(i).name Then
                        Worksheets(j).Move Before:=Worksheets(i)
                    End If
                End If
            Next j
        End If
    Next i
End Sub

Private Sub insertRowForTemplateSheet()
    Dim tempSheet As Worksheet
    If containsASheet(ThisWorkbook, getResByKey("PackageCustomTemplate"), tempSheet) Then
        tempSheet.Rows(1).Insert
    End If
End Sub

Public Function containsASheet(ByRef wb As Workbook, ByRef sheetName As String, Optional ByRef ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    containsASheet = True
    Set ws = wb.Worksheets(sheetName)
    Exit Function
ErrorHandler:
    containsASheet = False
End Function

Private Sub NextSheet()
    If ActiveSheet.index <> Worksheets.count Then
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

Private Sub setSubFeatureFrame(mergeStartRow As Integer, mergeEndRow As Integer, colCnt As Integer)
    Dim colStr As Variant
    colStr = Split(Trim("A*B*C*D*E*F*G*H*I*J*K*L*M*N*O*P*Q*R*S*T*U*V*W*X*Y*Z"), "*")
    Range("A" + CStr(mergeStartRow) + ":" + colStr(colCnt) + CStr(mergeEndRow)).Select
    Selection.BorderAround Weight:=xlMedium
    With Selection.Borders(xlInsideVertical)
        .Weight = xlMedium
    End With
    Range("A" + CStr(mergeEndRow + 3)).Select
    
End Sub

Private Sub setFrameFroMO(mergeStartRow As Integer, mergeEndRow As Integer)
    Range("A" + CStr(mergeStartRow) + ":" + "B" + CStr(mergeEndRow) + ":" + "C" + CStr(mergeEndRow)).Select
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

Private Sub setFrameNone4FeaturePkg(mergeStartRow As Integer, mergeEndRow As Integer, startCol As Integer, endCol As Integer)
    Dim arr()
    arr = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    Range(arr(startCol) + CStr(mergeStartRow) + ":" + arr(endCol) + CStr(mergeEndRow)).Select
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
    
    CfgArray = Split(Trim(readUTF8File(ThisWorkbook.path + "\Cfg.ini")), "*")
    mustGiveArr = CfgArray
    If (CfgArray(0) = "en") Then
        language = ENG
    Else
        language = CHN
    End If
    
    versionEndRow = CInt(CfgArray(1))
    mergeStartRow = CInt(CfgArray(2))
    mocCnt = CInt(CfgArray(3))
    attrStartRow = CInt(CfgArray(4))
    subFeatureSize = CInt(CfgArray(5))
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
    MsgBox Err.Description, , "错误 - " & Err.Number
    readUTF8File = ""
End Function

'设置批注自动调整大小
Private Sub refreshComment(ByRef myRange As Range)
    On Error Resume Next
    Dim cell As Range
    For Each cell In myRange
        If Not cell.Comment Is Nothing Then
            cell.Comment.Shape.TextFrame.AutoSize = True
        End If
    Next
End Sub

'set commont for suggestion line
Private Sub setComment(ByRef myRange As Range)
On Error Resume Next

    Dim recommendValTip As String
    recommendValTip = getResByKey("RecommendValueTip")
    Dim cell As Range
    For Each cell In myRange
        If cell.Comment Is Nothing Then
            cell.AddComment Text:=recommendValTip
            cell.Comment.Visible = False
        End If
    Next
    
End Sub

'Merge NENAME and CELLID column, these two columns display in the first and second column
Private Sub mergeSpecialColumns(sheetIndex As Long, validCol As Integer)
On Error Resume Next
    Application.DisplayAlerts = False
    
    Call unMergeRanges(sheetIndex, validCol)
    Dim curSheet As Worksheet
    Set curSheet = Worksheets(sheetIndex)
    Dim shortName As String
    Dim i As Integer
    For i = 1 To validCol
        shortName = curSheet.Cells(SHORTNAMEROW, i).value
        If curSheet.Range(curSheet.Cells(MERGETROW, i), curSheet.Cells(MERGETROW, i)).MergeArea.Cells(1, 1).value = "" Then
            curSheet.Cells(DISPLAYNAMEROW, i).Copy curSheet.Cells(MERGETROW, i)
            
            Dim mergeRange As Range
            Set mergeRange = curSheet.Range(curSheet.Cells(2, i), curSheet.Cells(5, i))
            mergeRange.Merge
            
            mergeRange.HorizontalAlignment = xlCenter
            mergeRange.VerticalAlignment = xlCenter
            
            With mergeRange.Borders(xlEdgeRight) ' set right border
                    .LineStyle = xlContinuous
                    .ColorIndex = 1
                    .Weight = xlThin
            End With
        End If
    Next i
    
End Sub

Private Sub unMergeRanges(sheetIndex As Long, validCol As Integer)
    Dim curSheet As Worksheet
    Set curSheet = Worksheets(sheetIndex)
    Dim i As Integer
    For i = 1 To validCol
        If curSheet.Range(curSheet.Cells(MERGETROW, i), curSheet.Cells(MERGETROW, i)).MergeArea.Cells(1, 1).value = "" Then
            curSheet.Range(curSheet.Cells(MERGETROW, i), curSheet.Cells(MERGETROW, i)).MergeArea.UnMerge
        End If
    Next i
        
End Sub

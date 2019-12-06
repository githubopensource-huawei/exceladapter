Attribute VB_Name = "Util"
Option Explicit

Private Const SW_SHOWNORMAL = 1
Public HasHistoryData As Boolean

Private Const listShtTitleRow = 2

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

Sub DisplayMessageOnStatusbar()
    Application.DisplayStatusBar = True '显示状态栏
    Application.StatusBar = "Running,please wait......" '状态栏显示信息

End Sub

Public Sub ReturnStatusbaring()
    Application.StatusBar = "Ready." '状态栏恢复正常
End Sub


'装载用于添加「Base Station Transport Data」页「*Site Template」列侯选值的窗体
Sub addTemplate()
    Load TemplateForm
    TemplateForm.Show
End Sub
Sub addIPRoute()
    Load IPRouteForm
    IPRouteForm.Show
End Sub
Sub addHyperlinks()
    Load HyperlinksForm
    HyperlinksForm.Show
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
    For m_rowNum = 2 To mappingDef.range("a65536").End(xlUp).row
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
        For m_colNum = 1 To ws.range("IV" + CStr(recordRow)).End(xlToLeft).column
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
    For rowNum = 1 To sheetDef.range("a65536").End(xlUp).row
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
    
    For rowNum = 1 To sheetDef.range("a65536").End(xlUp).row
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
        If Not isEmpty(ws.Cells(1, index).value) And ws.Cells(1, index).value <> "" Then
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
    For index = 2 To sheetDef.range("a65536").End(xlUp).row
        Set worksh = ThisWorkbook.Worksheets(sheetDef.Cells(index, 1).value)
        If sheetDef.Cells(index, 2) = "COMMON" Then
            For commIndex = 1 To worksh.range("a65536").End(xlUp).row
                For commCloumIndex = 1 To worksh.range("IV" + CStr(commIndex)).End(xlToLeft).column
                    If worksh.Cells(commIndex, commCloumIndex).Interior.colorIndex = 16 And _
                        worksh.Cells(commIndex, commCloumIndex).Interior.Pattern = xlGray16 Then
                            worksh.Cells(commIndex, commCloumIndex).Interior.colorIndex = xlNone
                            worksh.Cells(commIndex, commCloumIndex).Interior.Pattern = xlNone
                    End If
                Next
            Next
        ElseIf "Pattern" = sheetDef.Cells(index, 2) Then
            
        Else
            For cloumIndex = 1 To worksh.range("IV" + CStr(3)).End(xlToLeft).column
                If worksh.Cells(3, cloumIndex).Interior.colorIndex = 16 And _
                    worksh.Cells(3, cloumIndex).Interior.Pattern = xlGray16 Then
                        worksh.Cells(3, cloumIndex).Interior.colorIndex = xlNone
                        worksh.Cells(3, cloumIndex).Interior.Pattern = xlNone
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
    For m_rowNum = 2 To sheetDef.range("a65536").End(xlUp).row
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
    'maxRowNumber = ws.UsedRange.rows.count
    maxRowNumber = getUsedRowCount(ws)
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

Public Sub changeAlert(ByRef flag As Boolean)
    Application.DisplayAlerts = flag
    Application.ScreenUpdating = flag
End Sub

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


Public Function getUsedColumnCount(ByRef sheet As Worksheet, Optional ByRef attrNamerowIndex As Long = 2)
    getUsedColumnCount = sheet.range("IV2").End(xlToLeft).column
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


Public Function special() As Boolean
On Error GoTo ErrorHandler
    special = False
    Dim cover As String
    Dim key As String
    cover = getResByKey("Cover")
    key = ThisWorkbook.Worksheets(cover).Cells(2, 2).value
    If InStr(key, "RANCU_ENB") > 0 Then
       special = True
    End If
ErrorHandler:
    Debug.Print "some exception in special, " & Err.Description
End Function

'某行是否为空行
Public Function rowIsBlank(ByRef ws As Worksheet, ByRef rowNumber As Long) As Boolean
    If Application.WorksheetFunction.CountBlank(ws.range("A" & rowNumber & ":IV" & rowNumber)) = 256 Then
        rowIsBlank = True
    Else
        rowIsBlank = False
    End If
End Function

'[Common]find next group name row from empty row
Public Function findNextGrpNameRow(sht As Worksheet, ByVal startRow As Long) As Long
    findNextGrpNameRow = -1
    Dim rowIdx As Long
    For rowIdx = startRow To sht.range("a65535").End(xlUp).row
        If Not rowIsBlank(sht, rowIdx) Then
            findNextGrpNameRow = rowIdx
            Exit Function
        End If
    Next
End Function

'[Common]find next group name row from data row
Public Function findNextGrpNameRowEx(sht As Worksheet, ByVal startRow As Long) As Long
    findNextGrpNameRowEx = -1
    Dim rowIdx As Long
    For rowIdx = startRow To sht.range("a65535").End(xlUp).row
        If rowIsBlank(sht, rowIdx) Then
            findNextGrpNameRowEx = findNextGrpNameRow(sht, rowIdx)
            Exit Function
        End If
    Next
End Function

Public Sub clearValidation(target As range)
On Error GoTo ErrorHandler
    With target.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .inputTitle = ""
        .ErrorTitle = ""
        .inputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    target.value = ""
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in clearValidation, " & Err.Description
End Sub

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

Public Function collectionJoin(coll As Collection, Optional delimiter As String = ",") As String
On Error GoTo ErrorHandler
    collectionJoin = ""
    If coll.count = 0 Then Exit Function
    
    Dim deli As String
    deli = ""
    
    Dim item
    For Each item In coll
        collectionJoin = collectionJoin & deli & CStr(item)
        deli = delimiter
    Next
    Exit Function
ErrorHandler:
    Debug.Print "some exception in collectionJoin, " & Err.Description
End Function

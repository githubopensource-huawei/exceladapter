Attribute VB_Name = "AddLinkModule"
Option Explicit

Public Const SheetType_List As String = "LIST"
Public Const SheetType_Pattern As String = "PATTERN"

Public Const StandardRowHeight As Double = 13.5

Public Const HyperLinkColorIndex As Integer = 6
Public Const BluePrintSheetColor As Integer = 5
Public Const refColInModDiffSht As Integer = 1



Public Sub AddLink()
    Application.ScreenUpdating = False
    
    Call addLink4NormalShts
    
    If existsASheet(getResByKey("ModelDiffSht")) Then
        Call addLink4MultiVer
    End If
    
    Application.ScreenUpdating = True
End Sub

Public Sub addLink4NormalShts()
    Application.ScreenUpdating = False
    On Error Resume Next

    Dim m_rowNum As Long
    Dim n_RowNum As Long
    Dim textValue As String
    Dim shtNameColNum As Long
    Dim grpNameColNum As Long
    Dim colNameColNum As Long
    Dim isRefColNum As Long

    Dim srcShtName As String
    Dim srcGrpName As String
    Dim srcColName As String
    Dim srcColNum As Long
    Dim shtType As String

    shtNameColNum = shtNameColNumInMappingDef
    grpNameColNum = grpNameColNumInMappingDef
    colNameColNum = colNameColNumInMappingDef
    isRefColNum = isRefColNumInMappingDef
    
    Dim mappingDef As Worksheet
    Set mappingDef = ThisWorkbook.Worksheets("MAPPING DEF")
    Dim srcSht As Worksheet
    
    '遍历『MAPPING DEF』，获取需要添加Reference的列:Is Reference = true
    With mappingDef
        For m_rowNum = 2 To .Range("a65536").End(xlUp).row
            If UCase(Trim(.Cells(m_rowNum, isRefColNum).value)) = "TRUE" Then
                srcShtName = .Cells(m_rowNum, shtNameColNum).value
                shtType = getSheetType(srcShtName)
                If shtType = SheetType_List Or shtType = SheetType_Pattern Then
                    srcGrpName = .Cells(m_rowNum, grpNameColNum).value
                    srcColName = .Cells(m_rowNum, colNameColNum).value
                    srcColNum = Get_RefCol(srcShtName, listShtAttrRow, srcGrpName, srcColName)
                    
                    Set srcSht = ThisWorkbook.Worksheets(srcShtName)
                    With srcSht
                        For n_RowNum = 3 To .UsedRange.Rows.count
                            textValue = .Cells(n_RowNum, srcColNum).value
                            Dim refArray As Variant
                            If isValidReference(textValue, refArray) Then
                                Call addNormalShtLink(srcSht, .Cells(n_RowNum, srcColNum), CStr(refArray(0)), CStr(refArray(1)), CStr(refArray(2)), textValue)
                            Else
                                .Cells(n_RowNum, srcColNum).Hyperlinks.Delete
                            End If
                        Next
                    End With
                End If
            End If
        Next
    End With
    
    Application.ScreenUpdating = True
End Sub

Public Sub addLink4MultiVer()
    Dim modelDiffSht As Worksheet
    Set modelDiffSht = ThisWorkbook.Worksheets(getResByKey("ModelDiffSht"))
    

    Dim rowIdx As Integer
    Dim refValue As String
    Dim refArray As Variant
    With modelDiffSht
        For rowIdx = 2 To .Range("a65535").End(xlUp).row
            refValue = .Cells(rowIdx, refColInModDiffSht)
            If Not isValidReference(refValue, refArray) Then GoTo NextLoop
            
            Dim shtName As String
            Dim grpName As String
            Dim attrName As String
            shtName = refArray(0)
            grpName = refArray(1)
            attrName = refArray(2)
            attrName = Replace(attrName, "*", "~*")
            
            Call addModDiffHyperLink(shtName, grpName, attrName, modelDiffSht, rowIdx)
NextLoop:
        Next
    End With
    
    Call setModDiffShtFont(modelDiffSht)
End Sub


Private Sub addNormalShtLink(srcSht As Worksheet, srcRange As Range, dstShtName As String, dstGrpName As String, dstColName As String, textValue As String)
    Dim linkRowAdd As Long
    Dim targetRow As Long
    Dim targetColName As String
    linkRowAdd = 0
    targetRow = 2
    targetColName = dstColName
    
    If InStr(dstColName, "[") <> 0 Then
        linkRowAdd = CLng(Mid(dstColName, InStr(dstColName, "[") + 1, InStr(dstColName, "]") - InStr(dstColName, "[") - 1)) + 1
        targetColName = Mid(dstColName, 1, InStr(dstColName, "[") - 1)
    End If
    
    If dstShtName = getResByKey("Comm Data") Then
        targetRow = Get_GroupRefRow(dstShtName, dstGrpName) + 1
    End If
    
    Dim targetRangeAddr As String
    targetRangeAddr = "R" + CStr(targetRow + linkRowAdd) + "C" + CStr(Get_RefCol(dstShtName, targetRow, dstGrpName, targetColName))
    
    ThisWorkbook.ActiveSheet.Hyperlinks.Add anchor:=srcRange, Address:="", SubAddress:="'" + dstShtName + "'!" + targetRangeAddr, TextToDisplay:=textValue
End Sub

Private Sub addModDiffHyperLink(shtName As String, grpName As String, attrName As String, modelDiffSht As Worksheet, ByVal srcRow As Integer)
    If shtName = getResByKey("Board Style") Then
        Dim tmpSht As Worksheet
        For Each tmpSht In ThisWorkbook.Worksheets
            If isBoardStyleSheet(tmpSht) Then
                Call addModDiffHyperLink_i(tmpSht.name, grpName, attrName, modelDiffSht, srcRow)
            End If
        Next
    Else
        Call addModDiffHyperLink_i(shtName, grpName, attrName, modelDiffSht, srcRow)
    End If

End Sub

Public Function isBoardStyleSheet(ByRef ws As Worksheet) As Boolean
    Dim sheetName As String, boardStyleSheetName As String
    isBoardStyleSheet = False
    boardStyleSheetName = getResByKey("Board Style")
    If InStr(ws.name, boardStyleSheetName) <> 0 Then
        isBoardStyleSheet = True
    End If
End Function

Private Sub addModDiffHyperLink_i(shtName As String, grpName As String, attrName As String, modelDiffSht As Worksheet, ByVal srcRow As Integer)
    Dim linkRangeInModDiffSht As String
    linkRangeInModDiffSht = "'" & modelDiffSht.name & "'!" & "R" & srcRow & "C" & refColInModDiffSht
    
    Dim sht As Worksheet
    Dim targetRange As Range
    Dim firstAddr As String
    Dim firstGrpName As String
    
    firstGrpName = ""
    Dim visitedGrp As New Collection
    
    Set sht = ThisWorkbook.Worksheets(shtName)
    With sht
        Set targetRange = .UsedRange.Find(attrName, LookAt:=xlWhole, LookIn:=xlValues)
        If Not targetRange Is Nothing Then
            firstAddr = targetRange.Address
            Do
                firstGrpName = getGroupNameFromMappingDef(realShtName(shtName), attrName, visitedGrp)
                If grpName = firstGrpName Then
                    .Hyperlinks.Add anchor:=targetRange, Address:="", SubAddress:=linkRangeInModDiffSht
                    With targetRange.Font
                        .name = "Arial"
                        .Size = 10
                    End With
                    Exit Do
                End If
                Set targetRange = .UsedRange.FindNext(targetRange)
                If Not Contains(visitedGrp, firstGrpName) Then visitedGrp.Add Item:=firstGrpName, key:=firstGrpName
            Loop While Not targetRange Is Nothing And targetRange.Address <> firstAddr
        End If
    End With
    
    Dim linkRangeInTargetSht As String
    linkRangeInTargetSht = "'" & shtName & "'!" & "R" & targetRange.row & "C" & targetRange.Column
    With modelDiffSht
        .Hyperlinks.Add anchor:=.Range(c(refColInModDiffSht) & srcRow), Address:="", SubAddress:=linkRangeInTargetSht
    End With
End Sub

Function Contains(coll As Collection, key As String) As Boolean
        On Error GoTo NotFound
        Call coll(key)
        Contains = True
        Exit Function
NotFound:
        Contains = False
End Function

Private Function realShtName(shtName As String) As String
    realShtName = shtName
    If InStr(shtName, getResByKey("Board Style")) Then realShtName = getResByKey("Board Style")
End Function

Private Sub setModDiffShtFont(sht As Worksheet)
    Dim maxRow As Long
    Dim maxCol As Long
    Dim DataRange As Range
    Dim titleRange As Range
    Dim linkRange As Range
    Dim mocAttrRange As Range
    Dim versRange As Range

    With sht
        .Activate
        ActiveWindow.FreezePanes = False
        
        With .UsedRange
            maxRow = .Rows.count
            maxCol = .columns.count
        End With
        
        Set DataRange = .Range("A2:" & c(maxCol) & maxRow)
        With DataRange
            .Rows.EntireRow.RowHeight = 54
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
        
        Set titleRange = .Range("A1:" & c(maxCol) & "1")
        With titleRange
            .Rows.EntireRow.RowHeight = StandardRowHeight
            .Interior.colorIndex = 40
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
        End With
        
        Set linkRange = .Range("A2:A" & maxRow)
        With linkRange
            .columns.EntireColumn.ColumnWidth = 40
            .Font.colorIndex = BluePrintSheetColor
            .WrapText = False
        End With
        
        Set mocAttrRange = .Range("B2:C" & maxRow)
        With mocAttrRange
            .WrapText = False
            .columns.EntireColumn.AutoFit
        End With
        
        Set versRange = .Range("D2:" & c(maxCol) & maxRow)
        With versRange
            .columns.EntireColumn.ColumnWidth = 50
            .WrapText = True
        End With
        
        Call setBorders(.UsedRange)
    End With
End Sub

Sub setBorders(ByRef certainRange As Range)
    On Error Resume Next
    certainRange.Borders(xlEdgeLeft).LineStyle = xlContinuous
    certainRange.Borders(xlEdgeTop).LineStyle = xlContinuous
    certainRange.Borders(xlEdgeBottom).LineStyle = xlContinuous
    certainRange.Borders(xlEdgeRight).LineStyle = xlContinuous
    certainRange.Borders(xlInsideVertical).LineStyle = xlContinuous
    certainRange.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    certainRange.Borders.LineStyle = xlContinuous
End Sub

Public Function existsASheet(shtName As String) As Boolean
On Error GoTo ErrorHandler:
    existsASheet = True
    Dim sht As Worksheet
    Set sht = ThisWorkbook.Worksheets(shtName)
    Exit Function
ErrorHandler:
    existsASheet = False
End Function


'从指定sheet页的指定行，查找指定列，返回列号
Function Get_RefCol(sheetName As String, RecordRow As Long, groupName As String, ColValue As String) As Long
    On Error GoTo ErrorHandler
    Dim m_ColNum As Long
    Dim m_GroupColNum As Long
    Dim f_flag As Boolean
    f_flag = False
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    For m_ColNum = 1 To ws.Range("IV" + CStr(RecordRow)).End(xlToLeft).Column
        If Get_DesStr(ColValue) = Get_DesStr(ws.Cells(RecordRow, m_ColNum).value) Then
            If groupName = "" Then
                f_flag = True
                Exit For
            Else
                m_GroupColNum = m_ColNum
                While Get_DesStr(ws.Cells(RecordRow - 1, m_GroupColNum).value) = ""
                    m_GroupColNum = m_GroupColNum - 1
                Wend
                If Get_DesStr(groupName) = Get_DesStr(ws.Cells(RecordRow - 1, m_GroupColNum).value) Then
                    f_flag = True
                    Exit For
                End If
            End If
        End If
    Next
    If f_flag = False Then
        'MsgBox sheetName & "缺少列：" & ColValue, vbExclamation, "Warning"
    Else
        Get_RefCol = m_ColNum
    End If
    Exit Function
ErrorHandler:
    Get_RefCol = -1
End Function

'将比较字符串整形
Public Function Get_DesStr(srcStr As String) As String
    Get_DesStr = UCase(Trim(srcStr))
End Function


'检查是否为合法超链接格式
Public Function isValidReference(refAddr As String, Optional refArray As Variant, Optional delimeter As String) As Boolean
    isValidReference = False
    
    If delimeter <> "" Then
        refArray = Split(refAddr, delimeter)
        If UBound(refArray) <> 2 Then Exit Function
        If refArray(0) = "" Or refArray(1) = "" Or refArray(2) = "" Then Exit Function
        isValidReference = True
        Exit Function
    End If
    
    If isValidReference(refAddr, refArray, "\") Then
        isValidReference = True
        Exit Function
    End If
    
    If isValidReference(refAddr, refArray, ".") Then
        isValidReference = True
        Exit Function
    End If
End Function


'从指定sheet页查找group所在行
Function Get_GroupRefRow(sheetName As String, groupName As String) As Long
    Dim m_rowNum As Long
    Dim f_flag As Boolean
    f_flag = False
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    For m_rowNum = 1 To ws.Range("a65536").End(xlUp).row
        If Get_DesStr(groupName) = Get_DesStr(ws.Cells(m_rowNum, 1).value) Then
            f_flag = True
            Exit For
        End If
    Next
    If f_flag = False Then
        'MsgBox sheetName & "缺少Group：" & groupName, vbExclamation, "Warning"
    End If
    
    Get_GroupRefRow = m_rowNum
    
End Function

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
        For m_rowNum = 2 To .range("a1048576").End(xlUp).row
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
        For rowIdx = 2 To .range("a1048576").End(xlUp).row
            refValue = .Cells(rowIdx, refColInModDiffSht)
            If Not isValidReference(refValue, refArray) Then GoTo NextLoop
            
            Dim shtname As String
            Dim grpName As String
            Dim attrName As String
            shtname = refArray(0)
            grpName = refArray(1)
            attrName = refArray(2)
            attrName = Replace(attrName, "*", "~*")
            
            Call addModDiffHyperLink(shtname, grpName, attrName, modelDiffSht, rowIdx)
NextLoop:
        Next
    End With
    
    Call setModDiffShtFont(modelDiffSht)
End Sub


Private Sub addNormalShtLink(srcSht As Worksheet, srcRange As range, dstShtName As String, dstGrpName As String, dstColName As String, textValue As String)
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
    
    ThisWorkbook.ActiveSheet.Hyperlinks.Add Anchor:=srcRange, address:="", SubAddress:="'" + dstShtName + "'!" + targetRangeAddr, TextToDisplay:=textValue
End Sub

Private Sub addModDiffHyperLink(shtname As String, grpName As String, attrName As String, modelDiffSht As Worksheet, ByVal srcRow As Integer)
    If shtname = getResByKey("Board Style") Then
        Dim tmpSht As Worksheet
        For Each tmpSht In ThisWorkbook.Worksheets
            If isBoardStyleSheet(tmpSht) Then
                Call addModDiffHyperLink_i(tmpSht.name, grpName, attrName, modelDiffSht, srcRow)
            End If
        Next
    Else
        Call addModDiffHyperLink_i(shtname, grpName, attrName, modelDiffSht, srcRow)
    End If

End Sub

Private Sub addModDiffHyperLink_i(shtname As String, grpName As String, attrName As String, modelDiffSht As Worksheet, ByVal srcRow As Integer)
    Dim linkRangeInModDiffSht As String
    linkRangeInModDiffSht = "'" & modelDiffSht.name & "'!" & "R" & srcRow & "C" & refColInModDiffSht
    
    Dim sht As Worksheet
    Dim targetRange As range
    Dim firstAddr As String
    Dim firstGrpName As String
    
    firstGrpName = ""
    Dim visitedGrp As New Collection
    
    Set sht = ThisWorkbook.Worksheets(shtname)
    With sht
        Set targetRange = .UsedRange.Find(attrName, lookat:=xlWhole, LookIn:=xlValues)
        If Not targetRange Is Nothing Then
            firstAddr = targetRange.address
            Do
                firstGrpName = getGroupNameFromMappingDef(realShtName(shtname), attrName, visitedGrp)
                If grpName = firstGrpName Then
                    .Hyperlinks.Add Anchor:=targetRange, address:="", SubAddress:=linkRangeInModDiffSht
                    With targetRange.Font
                        .name = "Arial"
                        .Size = 10
                    End With
                    Exit Do
                End If
                Set targetRange = .UsedRange.FindNext(targetRange)
                If Not Contains(visitedGrp, firstGrpName) Then visitedGrp.Add Item:=firstGrpName, key:=firstGrpName
            Loop While Not targetRange Is Nothing And targetRange.address <> firstAddr
        End If
    End With
    
    Dim linkRangeInTargetSht As String
    linkRangeInTargetSht = "'" & shtname & "'!" & "R" & targetRange.row & "C" & targetRange.column
    With modelDiffSht
        .Hyperlinks.Add Anchor:=.range(C(refColInModDiffSht) & srcRow), address:="", SubAddress:=linkRangeInTargetSht
    End With
End Sub

Private Function realShtName(shtname As String) As String
    realShtName = shtname
    If InStr(shtname, getResByKey("Board Style")) Then realShtName = getResByKey("Board Style")
End Function

Private Sub setModDiffShtFont(sht As Worksheet)
    Dim maxRow As Long
    Dim maxCol As Long
    Dim DataRange As range
    Dim titleRange As range
    Dim linkRange As range
    Dim mocAttrRange As range
    Dim versRange As range

    With sht
        .Activate
        ActiveWindow.FreezePanes = False
        
        With .UsedRange
            maxRow = .Rows.count
            maxCol = .Columns.count
        End With
        
        Set DataRange = .range("A2:" & C(maxCol) & maxRow)
        With DataRange
            .Rows.EntireRow.RowHeight = 54
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
        
        Set titleRange = .range("A1:" & C(maxCol) & "1")
        With titleRange
            .Rows.EntireRow.RowHeight = StandardRowHeight
            .Interior.colorIndex = 40
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
        End With
        
        Set linkRange = .range("A2:A" & maxRow)
        With linkRange
            .Columns.EntireColumn.ColumnWidth = 40
            .Font.colorIndex = BluePrintSheetColor
            .WrapText = False
        End With
        
        Set mocAttrRange = .range("B2:C" & maxRow)
        With mocAttrRange
            .WrapText = False
            .Columns.EntireColumn.AutoFit
        End With
        
        Set versRange = .range("D2:" & C(maxCol) & maxRow)
        With versRange
            .Columns.EntireColumn.ColumnWidth = 50
            .WrapText = True
        End With
        
        Call setBorders(.UsedRange)
    End With
End Sub


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IPRouteForm 
   Caption         =   "IPRoute auto-compute"
   ClientHeight    =   10170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13905
   OleObjectBlob   =   "IPRouteForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "IPRouteForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'激活事件
Private Sub UserForm_Activate()
    IPRouteForm.Caption = getResByKey("Bar_IPRoute")
    SetDstSheet
    SetListSetting
    SetColWidth
    
    IPRouteForm.CommandButton3.SetFocus
End Sub

'从「SHEET DEF」页获取Sheet页侯选值
Private Sub SetDstSheet()
    For m_rowNum = 2 To Worksheets("SHEET DEF").range("a1048576").End(xlUp).row
        If Worksheets("SHEET DEF").Cells(m_rowNum, 1).value <> "eNodeB Radio Data" And Worksheets("SHEET DEF").Cells(m_rowNum, 2).value <> "Pattern" _
            And Check_Sheet(Worksheets("MAPPING DEF").Cells(m_rowNum, 1).value) _
            And Worksheets(Worksheets("SHEET DEF").Cells(m_rowNum, 1).value).Visible = xlSheetVisible Then
            IPRouteForm.DstIPSheet.AddItem (Worksheets("SHEET DEF").Cells(m_rowNum, 1).value)
            IPRouteForm.DstMaskSheet.AddItem (Worksheets("SHEET DEF").Cells(m_rowNum, 1).value)
        End If
    Next
End Sub

'从「MAPPING DEF」页获取DstIP_Group
Private Sub DstIPSheet_Change()
    Dim addFlag As Boolean
    
    IPRouteForm.DstIPGroup.Clear
    IPRouteForm.IPGroupList.Clear
    
    For m_rowNum = 2 To Worksheets("MAPPING DEF").range("a1048576").End(xlUp).row
        If IPRouteForm.DstIPSheet.value = Worksheets("MAPPING DEF").Cells(m_rowNum, 1).value Then
            addFlag = True
            For i = 0 To IPRouteForm.DstIPGroup.ListCount - 1
                If IPRouteForm.DstIPGroup.List(i) = Worksheets("MAPPING DEF").Cells(m_rowNum, 2).value Then
                    addFlag = False
                    Exit For
                End If
            Next
            If addFlag Then
                addFlag = Check_Group(Worksheets("MAPPING DEF").Cells(m_rowNum, 1).value + "." + Worksheets("MAPPING DEF").Cells(m_rowNum, 2).value)
            End If
            If addFlag Then
                IPRouteForm.DstIPGroup.AddItem (Worksheets("MAPPING DEF").Cells(m_rowNum, 2).value)
                IPRouteForm.IPGroupList.AddItem (Worksheets("MAPPING DEF").Cells(m_rowNum, 2).value)
            End If
        End If
    Next
End Sub

'从「MAPPING DEF」页获取DstMask_Group
Private Sub DstMaskSheet_Change()
    Dim addFlag As Boolean
    
    IPRouteForm.DstMaskGroup.Clear
    IPRouteForm.MaskGroupList.Clear
    
    For m_rowNum = 2 To Worksheets("MAPPING DEF").range("a1048576").End(xlUp).row
        If IPRouteForm.DstMaskSheet.value = Worksheets("MAPPING DEF").Cells(m_rowNum, 1).value Then
            addFlag = True
            For i = 0 To IPRouteForm.DstMaskGroup.ListCount - 1
                If IPRouteForm.DstMaskGroup.List(i) = Worksheets("MAPPING DEF").Cells(m_rowNum, 2).value Then
                    addFlag = False
                    Exit For
                End If
            Next
            If addFlag Then
                addFlag = Check_Group(Worksheets("MAPPING DEF").Cells(m_rowNum, 1).value + "." + Worksheets("MAPPING DEF").Cells(m_rowNum, 2).value)
            End If
            If addFlag Then
                IPRouteForm.DstMaskGroup.AddItem (Worksheets("MAPPING DEF").Cells(m_rowNum, 2).value)
                IPRouteForm.MaskGroupList.AddItem (Worksheets("MAPPING DEF").Cells(m_rowNum, 2).value)
            End If
        End If
    Next
End Sub

'从「MAPPING DEF」页获取DstIP_Col
Private Sub DstIPGroup_Change()
    Dim addFlag As Boolean
    
    IPRouteForm.DstIPCol.Clear
    For m_rowNum = 2 To Worksheets("MAPPING DEF").range("a1048576").End(xlUp).row
        If IPRouteForm.DstIPSheet.value = Worksheets("MAPPING DEF").Cells(m_rowNum, 1).value _
            And IPRouteForm.DstIPGroup.value = Worksheets("MAPPING DEF").Cells(m_rowNum, 2).value Then
            addFlag = True
            For i = 0 To IPRouteForm.DstIPCol.ListCount - 1
                If IPRouteForm.DstIPCol.List(i) = Worksheets("MAPPING DEF").Cells(m_rowNum, 3).value Then
                    addFlag = False
                    Exit For
                End If
            Next
            If addFlag Then
                addFlag = Check_Col(Worksheets("MAPPING DEF").Cells(m_rowNum, 1).value + "." + Worksheets("MAPPING DEF").Cells(m_rowNum, 2).value + "." + Worksheets("MAPPING DEF").Cells(m_rowNum, 3).value)
            End If
            If addFlag Then
                IPRouteForm.DstIPCol.AddItem (Worksheets("MAPPING DEF").Cells(m_rowNum, 3).value)
            End If
        End If
    Next
        
    If IPRouteForm.DstIPSheet.value = "Common Data" Then
        IPRouteForm.IPGroupList.value = IPRouteForm.DstIPGroup.value
        IPRouteForm.IPGroupText.value = IPRouteForm.DstIPGroup.value
        IPRouteForm.IPGroupList.Locked = True
        IPRouteForm.IPGroupText.Locked = True
        IPRouteForm.IPGroupList.BackColor = &H8000000F
        IPRouteForm.IPGroupText.BackColor = &H8000000F
    Else
        IPRouteForm.IPGroupList.Locked = False
        IPRouteForm.IPGroupText.Locked = False
        IPRouteForm.IPGroupList.BackColor = &H80000005
        IPRouteForm.IPGroupText.BackColor = &H80000005
        IPRouteForm.IPGroupText.value = ""
        IPRouteForm.IPColText.value = ""
    End If
End Sub

'从「MAPPING DEF」页获取DstMask_Col
Private Sub DstMaskGroup_Change()
    Dim addFlag As Boolean
 
    IPRouteForm.DstMaskCol.Clear
    For m_rowNum = 2 To Worksheets("MAPPING DEF").range("a1048576").End(xlUp).row
        If IPRouteForm.DstMaskSheet.value = Worksheets("MAPPING DEF").Cells(m_rowNum, 1).value _
            And IPRouteForm.DstMaskGroup.value = Worksheets("MAPPING DEF").Cells(m_rowNum, 2).value Then
            addFlag = True
            For i = 0 To IPRouteForm.DstMaskCol.ListCount - 1
                If IPRouteForm.DstMaskCol.List(i) = Worksheets("MAPPING DEF").Cells(m_rowNum, 3).value Then
                    addFlag = False
                    Exit For
                End If
            Next
            If addFlag Then
                addFlag = Check_Col(Worksheets("MAPPING DEF").Cells(m_rowNum, 1).value + "." + Worksheets("MAPPING DEF").Cells(m_rowNum, 2).value + "." + Worksheets("MAPPING DEF").Cells(m_rowNum, 3).value)
            End If
            If addFlag Then
                IPRouteForm.DstMaskCol.AddItem (Worksheets("MAPPING DEF").Cells(m_rowNum, 3).value)
            End If
        End If
    Next
    If IPRouteForm.DstMaskSheet.value = "Common Data" Then
        IPRouteForm.MaskGroupList.value = IPRouteForm.DstMaskGroup.value
        IPRouteForm.MaskGroupText.value = IPRouteForm.DstMaskGroup.value
        IPRouteForm.MaskGroupList.Locked = True
        IPRouteForm.MaskGroupText.Locked = True
        IPRouteForm.MaskGroupList.BackColor = &H8000000F
        IPRouteForm.MaskGroupText.BackColor = &H8000000F
    Else
        IPRouteForm.MaskGroupList.Locked = False
        IPRouteForm.MaskGroupText.Locked = False
        IPRouteForm.MaskGroupList.BackColor = &H80000005
        IPRouteForm.MaskGroupText.BackColor = &H80000005
        IPRouteForm.MaskGroupText.value = ""
        IPRouteForm.MaskColText.value = ""
    End If
End Sub

'从「MAPPING DEF」页获取IPRouteIP_Col
Private Sub IPGroupList_Change()
    Dim addFlag As Boolean
 
    IPRouteForm.IPColList.Clear
    For m_rowNum = 2 To Worksheets("MAPPING DEF").range("a1048576").End(xlUp).row
        If IPRouteForm.DstIPSheet.value = Worksheets("MAPPING DEF").Cells(m_rowNum, 1).value _
            And IPRouteForm.IPGroupList.value = Worksheets("MAPPING DEF").Cells(m_rowNum, 2).value Then
            addFlag = True
            For i = 0 To IPRouteForm.IPColList.ListCount - 1
                If IPRouteForm.IPColList.List(i) = Worksheets("MAPPING DEF").Cells(m_rowNum, 3).value Then
                    addFlag = False
                    Exit For
                End If
            Next
            If addFlag Then
                addFlag = Check_Col(Worksheets("MAPPING DEF").Cells(m_rowNum, 1).value + "." + Worksheets("MAPPING DEF").Cells(m_rowNum, 2).value + "." + Worksheets("MAPPING DEF").Cells(m_rowNum, 3).value)
            End If
            If addFlag Then
                IPRouteForm.IPColList.AddItem (Worksheets("MAPPING DEF").Cells(m_rowNum, 3).value)
            End If
        End If
    Next
End Sub

'从「MAPPING DEF」页获取IPRouteMask_Col
Private Sub MaskGroupList_Change()
    Dim addFlag As Boolean
 
    IPRouteForm.MaskColList.Clear
    For m_rowNum = 2 To Worksheets("MAPPING DEF").range("a1048576").End(xlUp).row
        If IPRouteForm.DstMaskSheet.value = Worksheets("MAPPING DEF").Cells(m_rowNum, 1).value _
            And IPRouteForm.MaskGroupList.value = Worksheets("MAPPING DEF").Cells(m_rowNum, 2).value Then
            addFlag = True
            For i = 0 To IPRouteForm.MaskColList.ListCount - 1
                If IPRouteForm.MaskColList.List(i) = Worksheets("MAPPING DEF").Cells(m_rowNum, 3).value Then
                    addFlag = False
                    Exit For
                End If
            Next
            If addFlag Then
                addFlag = Check_Col(Worksheets("MAPPING DEF").Cells(m_rowNum, 1).value + "." + Worksheets("MAPPING DEF").Cells(m_rowNum, 2).value + "." + Worksheets("MAPPING DEF").Cells(m_rowNum, 3).value)
            End If
            If addFlag Then
                IPRouteForm.MaskColList.AddItem (Worksheets("MAPPING DEF").Cells(m_rowNum, 3).value)
            End If
        End If
    Next
End Sub

Private Sub IPAdd_Click()
    IPRouteForm.IPGroupList.Visible = False
    IPRouteForm.IPColList.Visible = False
    IPRouteForm.IPGroupText.Visible = True
    IPRouteForm.IPColText.Visible = True
End Sub

Private Sub IPSelect_Click()
    IPRouteForm.IPGroupList.Visible = True
    IPRouteForm.IPColList.Visible = True
    IPRouteForm.IPGroupText.Visible = False
    IPRouteForm.IPColText.Visible = False
End Sub

Private Sub MaskAdd_Click()
    IPRouteForm.MaskGroupList.Visible = False
    IPRouteForm.MaskColList.Visible = False
    IPRouteForm.MaskGroupText.Visible = True
    IPRouteForm.MaskColText.Visible = True
End Sub

Private Sub MaskSelect_Click()
    IPRouteForm.MaskGroupList.Visible = True
    IPRouteForm.MaskColList.Visible = True
    IPRouteForm.MaskGroupText.Visible = False
    IPRouteForm.MaskColText.Visible = False
End Sub

'「增加」按钮事件
Private Sub CommandButton1_Click()
    Dim dstIP As String
    Dim dstMask As String
    Dim iPRouteIP As String
    Dim iPRouteMask As String
    
    Dim iPRouteIPGroup As String
    Dim iPRouteMaskGroup As String
    Dim iPRouteIpCol As String
    Dim iPRouteMaskCol As String
    
    Dim rowCount As Long
    Dim i As Long
    
    Dim flag As Boolean
    flag = True
    
    If IPRouteForm.IPGroupList.Visible Then
        iPRouteIPGroup = IPRouteForm.IPGroupList.value
    Else
        iPRouteIPGroup = IPRouteForm.IPGroupText.value
    End If
    
    If IPRouteForm.IPColList.Visible Then
        iPRouteIpCol = IPRouteForm.IPColList.value
    Else
        iPRouteIpCol = IPRouteForm.IPColText.value
    End If
    
    If IPRouteForm.MaskGroupList.Visible Then
        iPRouteMaskGroup = IPRouteForm.MaskGroupList.value
    Else
        iPRouteMaskGroup = IPRouteForm.MaskGroupText.value
    End If
    
    If IPRouteForm.MaskColList.Visible Then
        iPRouteMaskCol = IPRouteForm.MaskColList.value
    Else
        iPRouteMaskCol = IPRouteForm.MaskColText.value
    End If
    
    dstIP = IPRouteForm.DstIPSheet.value + "." + IPRouteForm.DstIPGroup.value + "." + IPRouteForm.DstIPCol.value
    dstMask = IPRouteForm.DstMaskSheet.value + "." + IPRouteForm.DstMaskGroup.value + "." + IPRouteForm.DstMaskCol.value
    iPRouteIP = IPRouteForm.DstIPSheet.value + "." + iPRouteIPGroup + "." + iPRouteIpCol
    iPRouteMask = IPRouteForm.DstMaskSheet.value + "." + iPRouteMaskGroup + "." + iPRouteMaskCol
    
    Dim v
    v = Split(dstIP + "." + dstMask + "." + iPRouteIP + "." + iPRouteMask, ".")
    For i = 0 To UBound(v)
        If v(i) = "" Then
            MsgBox "please input setting data!", vbExclamation, "Warning"
            flag = False
            Exit For
        End If
    Next
    
    If flag = True Then
        rowCount = Worksheets("IPRouteMap").range("a1048576").End(xlUp).row
        Worksheets("IPRouteMap").Cells(rowCount + 1, 1).value = IPRouteForm.DstIPSheet.value
        Worksheets("IPRouteMap").Cells(rowCount + 1, 2).value = IPRouteForm.DstIPGroup.value
        Worksheets("IPRouteMap").Cells(rowCount + 1, 3).value = IPRouteForm.DstIPCol.value
        Worksheets("IPRouteMap").Cells(rowCount + 1, 4).value = iPRouteIPGroup
        Worksheets("IPRouteMap").Cells(rowCount + 1, 5).value = iPRouteIpCol
        
        Worksheets("IPRouteMap").Cells(rowCount + 1, 6).value = IPRouteForm.DstMaskSheet.value
        Worksheets("IPRouteMap").Cells(rowCount + 1, 7).value = IPRouteForm.DstMaskGroup.value
        Worksheets("IPRouteMap").Cells(rowCount + 1, 8).value = IPRouteForm.DstMaskCol.value
        Worksheets("IPRouteMap").Cells(rowCount + 1, 9).value = iPRouteMaskGroup
        Worksheets("IPRouteMap").Cells(rowCount + 1, 10).value = iPRouteMaskCol
        SetListSetting
        SetColWidth
    End If
End Sub

'「删除」按钮事件
Private Sub CommandButton2_Click()
    If IPRouteForm.DstIPList.ListIndex = -1 Then
        MsgBox "please select setting data!", vbExclamation, "Warning"
    Else
        curIdx = IPRouteForm.DstIPList.ListIndex
        Worksheets("IPRouteMap").rows(CStr(curIdx + 3) & ":" & CStr(curIdx + 3)).Delete
        
        SetListSetting
        SetColWidth
    End If
End Sub

'「确定」按钮事件
Private Sub CommandButton3_Click()
    Dim i As Long
    
    Dim dstIpRow As Long
    Dim dstIpRowEnd As Long
    Dim sourceIpRowEnd As Long
    Dim sourceMaskRowEnd As Long
    Dim DstIPCol As Long
    Dim dstIPValue As String
    Dim iPRouteIpCol As Long
    
    Dim dstMaskRow As Long     '与dstIpRow不在同一行时的处理
    Dim dstMaskRowEnd As Long
    Dim DstMaskCol As Long
    Dim dstMaskValue As String
    Dim iPRouteMaskCol As Long

    For i = 3 To Worksheets("IPRouteMap").range("a1048576").End(xlUp).row
        If Worksheets("IPRouteMap").Cells(i, 1).value = "Common Data" Then
            dstIpRow = Get_GroupRowOfCommon(Worksheets("IPRouteMap").Cells(i, 1).value, Worksheets("IPRouteMap").Cells(i, 2).value) + 1
            DstIPCol = Get_ColOfCommon(Worksheets("IPRouteMap").Cells(i, 1).value, dstIpRow, Worksheets("IPRouteMap").Cells(i, 3).value)
            iPRouteIpCol = Get_ColOfCommon(Worksheets("IPRouteMap").Cells(i, 1).value, dstIpRow, Worksheets("IPRouteMap").Cells(i, 5).value)
        Else
            dstIpRow = 2
            DstIPCol = Get_ColOfNormal(Worksheets("IPRouteMap").Cells(i, 1).value, Worksheets("IPRouteMap").Cells(i, 2).value, Worksheets("IPRouteMap").Cells(i, 3).value)
            iPRouteIpCol = Get_ColOfNormal(Worksheets("IPRouteMap").Cells(i, 1).value, Worksheets("IPRouteMap").Cells(i, 4).value, Worksheets("IPRouteMap").Cells(i, 5).value)
        End If
        
        If Worksheets("IPRouteMap").Cells(i, 6).value = "Common Data" Then
            dstMaskRow = Get_GroupRowOfCommon(Worksheets("IPRouteMap").Cells(i, 6).value, Worksheets("IPRouteMap").Cells(i, 7).value) + 1
            DstMaskCol = Get_ColOfCommon(Worksheets("IPRouteMap").Cells(i, 6).value, dstMaskRow, Worksheets("IPRouteMap").Cells(i, 8).value)
            iPRouteMaskCol = Get_ColOfCommon(Worksheets("IPRouteMap").Cells(i, 6).value, dstMaskRow, Worksheets("IPRouteMap").Cells(i, 10).value)
        Else
            dstMaskRow = 2
            DstMaskCol = Get_ColOfNormal(Worksheets("IPRouteMap").Cells(i, 6).value, Worksheets("IPRouteMap").Cells(i, 7).value, Worksheets("IPRouteMap").Cells(i, 8).value)
            iPRouteMaskCol = Get_ColOfNormal(Worksheets("IPRouteMap").Cells(i, 6).value, Worksheets("IPRouteMap").Cells(i, 9).value, Worksheets("IPRouteMap").Cells(i, 10).value)
        End If

        If Worksheets("IPRouteMap").Cells(i, 1).value = "Common Data" Then
            For dstIpRowEnd = 1 To Worksheets("Common Data").range("a1048576").End(xlUp).row
                If Worksheets("Common Data").Cells(dstIpRow + dstIpRowEnd, 1).Borders(xlEdgeRight).LineStyle = xlNone Then
                    Exit For
                End If
            Next
            dstIpRowEnd = dstIpRow + dstIpRowEnd - 1
        Else
            dstIpRowEnd = Worksheets(Worksheets("IPRouteMap").Cells(i, 1).value).range("a1048576").End(xlUp).row
        End If
        
        
         If Worksheets("IPRouteMap").Cells(i, 1).value = "Common Data" Then
            For sourceMaskRowEnd = 1 To Worksheets("Common Data").range("a1048576").End(xlUp).row
                If Worksheets("Common Data").Cells(dstMaskRow + sourceMaskRowEnd, iPRouteMaskCol).Borders(xlEdgeRight).LineStyle = xlNone Then
                    Exit For
                End If
            Next
            sourceMaskRowEnd = dstMaskRow + sourceMaskRowEnd - 1
        Else
            sourceMaskRowEnd = Worksheets(Worksheets("IPRouteMap").Cells(i, 6).value).range("a1048576").End(xlUp).row
        End If
        
        
        
        
        '遍历dstIpRow
        For dstIpRow = dstIpRow + 1 To dstIpRowEnd
            If Worksheets("IPRouteMap").Cells(i, 1).value = Worksheets("IPRouteMap").Cells(i, 6).value Then
            dstMaskRow = dstMaskRow + 1
            Else
                If dstMaskRow < sourceMaskRowEnd Then
                dstMaskRow = dstMaskRow + 1
                End If
            End If
            dstIPValue = Trim(Worksheets(Worksheets("IPRouteMap").Cells(i, 1).value).Cells(dstIpRow, DstIPCol).value)
            dstMaskValue = Trim(Worksheets(Worksheets("IPRouteMap").Cells(i, 6).value).Cells(dstMaskRow, DstMaskCol).value)
            If dstIPValue <> "" And dstMaskValue <> "" Then
                If IPRouteForm.OptionButton1.value = True Then
                    Worksheets(Worksheets("IPRouteMap").Cells(i, 1).value).Cells(dstIpRow, iPRouteIpCol).value = Get_IPRoute(dstIPValue, dstMaskValue)
                    Worksheets(Worksheets("IPRouteMap").Cells(i, 6).value).Cells(dstMaskRow, iPRouteMaskCol).value = dstMaskValue
                Else
                    Worksheets(Worksheets("IPRouteMap").Cells(i, 1).value).Cells(dstIpRow, iPRouteIpCol).value = dstIPValue
                    Worksheets(Worksheets("IPRouteMap").Cells(i, 6).value).Cells(dstMaskRow, iPRouteMaskCol).value = "255.255.255.255"
                End If
            End If
        Next
    Next
    
    Unload Me
End Sub

'「取消」按钮事件
Private Sub CommandButton4_Click()
    Unload Me
End Sub

Private Sub SetListSetting()
    Dim maxLeft
    Dim leftValue
    Dim rightValue
    Dim listValue
    Dim i As Long
    
    maxLeft = 0
    For i = 3 To Worksheets("IPRouteMap").range("a1048576").End(xlUp).row
        leftValue = "{" + Get_dstIP(i) + "[" + Worksheets("IPRouteMap").Cells(i, 4).value + "." + Worksheets("IPRouteMap").Cells(i, 5).value + "]}"
        rightValue = "{" + Get_dstMask(i) + "[" + Worksheets("IPRouteMap").Cells(i, 9).value + "." + Worksheets("IPRouteMap").Cells(i, 10).value + "]}"
        
        If Len(leftValue) > maxLeft Then
            maxLeft = Len(leftValue)
        End If
    Next
    
    IPRouteForm.DstIPList.Clear
    
    For i = 3 To Worksheets("IPRouteMap").range("a1048576").End(xlUp).row
        leftValue = "{" + Get_dstIP(i) + "[" + Worksheets("IPRouteMap").Cells(i, 4).value + "." + Worksheets("IPRouteMap").Cells(i, 5).value + "]}"
        rightValue = "{" + Get_dstMask(i) + "[" + Worksheets("IPRouteMap").Cells(i, 9).value + "." + Worksheets("IPRouteMap").Cells(i, 10).value + "]}"
        
        If maxLeft > Len(leftValue) Then
            For j = 1 To maxLeft - Len(leftValue)
                leftValue = " " + leftValue
            Next
        End If
        listValue = leftValue + rightValue
        IPRouteForm.DstIPList.AddItem (listValue)
    Next
End Sub

'设置列表宽度
Private Sub SetColWidth()
    Dim ipMaxLenth

    ipMaxLenth = 120
    For i = 0 To IPRouteForm.DstIPList.ListCount - 1
        If Len(IPRouteForm.DstIPList.List(i)) > ipMaxLenth Then
            ipMaxLenth = Len(IPRouteForm.DstIPList.List(i))
        End If
    Next
    IPRouteForm.DstIPList.ColumnWidths = 60 + ipMaxLenth * 4.5

End Sub

'从指定sheet页的指定行，查找指定列，返回列号
Function Get_ColOfCommon(sheetName As String, recordRow As Long, colName As String) As Long
    Dim m_colNum As Long
    Dim f_flag As Boolean
    Dim m_ColEnd As Long
    Dim dstIpRowEnd As Long
    f_flag = False
    
    m_ColEnd = Worksheets(sheetName).range("XFD" + CStr(recordRow)).End(xlToLeft).column
    For m_colNum = 1 To m_ColEnd
        If colName = Worksheets(sheetName).Cells(recordRow, m_colNum).value Then
            f_flag = True
            Exit For
        End If
    Next
    
    If f_flag = False Then
        Worksheets(sheetName).Select
        Worksheets(sheetName).Cells(recordRow, m_ColEnd + 1).value = colName

        For dstIpRowEnd = 1 To Worksheets(sheetName).range("a1048576").End(xlUp).row
            If Worksheets("Common Data").Cells(recordRow + dstIpRowEnd, 1).Borders(xlEdgeRight).LineStyle = xlNone Then
                Exit For
            End If
        Next
        range(Cells(recordRow - 1, 1), Cells(recordRow - 1, m_ColEnd)).Select
        Selection.UnMerge
        range(Cells(recordRow - 1, 1), Cells(recordRow - 1, m_ColEnd + 1)).Select
        Selection.Merge
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .colorIndex = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .colorIndex = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .colorIndex = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .colorIndex = 0
            .Weight = xlThin
        End With
        
        range(Cells(recordRow, m_ColEnd), Cells(recordRow + dstIpRowEnd, m_ColEnd)).Select
        Selection.Copy
        range(Cells(recordRow, m_ColEnd + 1), Cells(recordRow + dstIpRowEnd, m_ColEnd + 1)).Select
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        
        m_colNum = m_ColEnd + 1
        
        Dim mappingRow
        mappingRow = Worksheets("MAPPING DEF").range("a1048576").End(xlUp).row + 1
        Worksheets("MAPPING DEF").Cells(mappingRow, 1) = sheetName
        Worksheets("MAPPING DEF").Cells(mappingRow, 2) = Worksheets(sheetName).Cells(recordRow - 1, 1).value
        Worksheets("MAPPING DEF").Cells(mappingRow, 3) = colName
    End If
    
    Get_ColOfCommon = m_colNum
End Function

'从指定sheet页的指定行，查找指定列，返回列号
Function Get_ColOfNormal(sheetName As String, groupName As String, colName As String) As Long
    Dim m_colNum As Long
    Dim f_flag As Boolean
    
    Dim m_ColStart As Long
    Dim m_ColEnd As Long
    
    Dim dstIpRowEnd As Long
    f_flag = False
    
    m_ColStart = Worksheets(sheetName).range("XFD2").End(xlToLeft).column
    m_ColEnd = m_ColStart
    For m_colNum = 1 To Worksheets(sheetName).range("XFD1").End(xlToLeft).column
        If groupName = Worksheets(sheetName).Cells(1, m_colNum).value Then
            m_ColStart = m_colNum
            
            For m_ColEnd = m_colNum + 1 To Worksheets(sheetName).range("XFD2").End(xlToLeft).column
                If Worksheets(sheetName).Cells(1, m_ColEnd).value <> "" Then
                    Exit For
                End If
            Next
            
            Exit For
        End If
    Next
    
    For m_colNum = m_ColStart To m_ColEnd
        If colName = Worksheets(sheetName).Cells(2, m_colNum).value Then
            f_flag = True
            Exit For
        End If
    Next
    
    If f_flag = False Then
        If m_ColEnd <> m_ColStart Then
            m_colNum = m_colNum - 1
            Worksheets(sheetName).Select
            Worksheets(sheetName).Cells(2, m_ColEnd).Select
            Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
            
            Worksheets(sheetName).Cells(2, m_ColEnd).value = colName
            
            range(Cells(1, m_ColStart), Cells(1, m_ColEnd - 1)).Select
            Selection.UnMerge
            range(Cells(1, m_ColStart), Cells(1, m_ColEnd)).Select
            Selection.Merge
        Else
            m_ColEnd = m_ColEnd + 1
            Worksheets(sheetName).Select
            Worksheets(sheetName).Cells(1, m_ColEnd).value = groupName
            Worksheets(sheetName).Cells(1, m_ColEnd).Select
            Selection.Font.Bold = True
            Selection.HorizontalAlignment = xlCenter
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 16777164
            End With
            
            Worksheets(sheetName).Cells(2, m_ColEnd).value = colName
            Worksheets(sheetName).Cells(2, m_ColEnd).Select
            Selection.HorizontalAlignment = xlCenter
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 10079487
            End With
        End If
                     
        Dim rowCount
        For rowCount = 2 To Worksheets(sheetName).range("a1048576").End(xlUp).row
            If Worksheets(sheetName).Cells(rowCount, 1).Borders(xlEdgeRight).LineStyle = xlNone Then
                Exit For
            End If
        Next
        
        Worksheets(sheetName).range(Cells(1, m_ColEnd), Cells(rowCount, m_ColEnd)).Select
        Selection.Font.name = "Arial"
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .colorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .colorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .colorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .colorIndex = xlAutomatic
        End With
        With Selection.Borders(xlInsideVertical)
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .colorIndex = xlAutomatic
        End With
        
        Dim mappingRow
        mappingRow = Worksheets("MAPPING DEF").range("a1048576").End(xlUp).row + 1
        Worksheets("MAPPING DEF").Cells(mappingRow, 1) = sheetName
        Worksheets("MAPPING DEF").Cells(mappingRow, 2) = groupName
        Worksheets("MAPPING DEF").Cells(mappingRow, 3) = colName
    
    End If
    
    Get_ColOfNormal = m_colNum
End Function

'从指定sheet页查找group所在行
Function Get_GroupRowOfCommon(sheetName As String, groupName As String) As Long
    Dim m_rowNum As Long
    Dim m_colNum As Long
    Dim f_flag As Boolean
    f_flag = False
    
    For m_rowNum = 1 To Worksheets(sheetName).range("a1048576").End(xlUp).row
        If GetDesStr(groupName) = GetDesStr(Worksheets(sheetName).Cells(m_rowNum, 1).value) Then
            f_flag = True
            Exit For
        End If
    Next
    If f_flag = False Then
        'MsgBox sheetName & "缺少Group：" & groupName, vbExclamation, "Warning"
    End If
    
    Get_GroupRowOfCommon = m_rowNum
    
End Function

Function Get_dstIP(row As Long) As String
    Get_dstIP = Worksheets("IPRouteMap").Cells(row, 1).value + "." + Worksheets("IPRouteMap").Cells(row, 2).value + "." + Worksheets("IPRouteMap").Cells(row, 3).value
End Function

Function Get_iPRouteIP(row As Long) As String
    Get_iPRouteIP = Worksheets("IPRouteMap").Cells(row, 1).value + "." + Worksheets("IPRouteMap").Cells(row, 4).value + "." + Worksheets("IPRouteMap").Cells(row, 5).value
End Function

Function Get_dstMask(row As Long) As String
    Get_dstMask = Worksheets("IPRouteMap").Cells(row, 6).value + "." + Worksheets("IPRouteMap").Cells(row, 7).value + "." + Worksheets("IPRouteMap").Cells(row, 8).value
End Function

Function Get_iPRouteMask(row As Long) As String
    Get_iPRouteMask = Worksheets("IPRouteMap").Cells(row, 6).value + "." + Worksheets("IPRouteMap").Cells(row, 9).value + "." + Worksheets("IPRouteMap").Cells(row, 10).value
End Function

Function Get_IPRoute(dstIPValue As String, dstMaskValue As String) As String
    Dim regIp As New VBScript_RegExp_55.regExp
    Dim result As String
    Dim vIP
    Dim vMask
    
    regIp.Pattern = "^(\d{1,2}|1\d\d|2[0-4]\d|25[0-5]).(\d{1,2}|1\d\d|2[0-4]\d|25[0-5]).(\d{1,2}|1\d\d|2[0-4]\d|25[0-5]).(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])$"
    
    If regIp.test(dstIPValue) And regIp.test(dstMaskValue) Then
        vIP = Split(dstIPValue, ".")
        vMask = Split(dstMaskValue, ".")
        result = CStr(vIP(0) And vMask(0)) + "." + CStr(vIP(1) And vMask(1)) + "." + CStr(vIP(2) And vMask(2)) + "." + CStr(vIP(3) And vMask(3))
        'MsgBox result, vbExclamation, "Warning"
    ElseIf regIp.test(dstIPValue) = False Then
        MsgBox dstIPValue & " is invalid ip address", vbExclamation, "Warning"
    ElseIf regIp.test(dstMaskValue) = False Then
        MsgBox dstMaskValue & " is invalid ip address", vbExclamation, "Warning"
    End If

    Get_IPRoute = result
End Function

Function Check_Col(ColValue As String) As Boolean
    Dim flag As Boolean
    Dim i As Long
    
    flag = True
    
    If IPRouteForm.IPGroupList.Visible Then
        iPRouteIPGroup = IPRouteForm.IPGroupList.value
    Else
        iPRouteIPGroup = IPRouteForm.IPGroupText.value
    End If
    
    If IPRouteForm.IPColList.Visible Then
        iPRouteIpCol = IPRouteForm.IPColList.value
    Else
        iPRouteIpCol = IPRouteForm.IPColText.value
    End If
    
    If IPRouteForm.MaskGroupList.Visible Then
        iPRouteMaskGroup = IPRouteForm.MaskGroupList.value
    Else
        iPRouteMaskGroup = IPRouteForm.MaskGroupText.value
    End If
    
    If IPRouteForm.MaskColList.Visible Then
        iPRouteMaskCol = IPRouteForm.MaskColList.value
    Else
        iPRouteMaskCol = IPRouteForm.MaskColText.value
    End If
    
    dstIP = IPRouteForm.DstIPSheet.value + "." + IPRouteForm.DstIPGroup.value + "." + IPRouteForm.DstIPCol.value
    dstMask = IPRouteForm.DstMaskSheet.value + "." + IPRouteForm.DstMaskGroup.value + "." + IPRouteForm.DstMaskCol.value
    iPRouteIP = IPRouteForm.DstIPSheet.value + "." + iPRouteIPGroup + "." + iPRouteIpCol
    iPRouteMask = IPRouteForm.DstMaskSheet.value + "." + iPRouteMaskGroup + "." + iPRouteMaskCol
    
    If ColValue = dstIP Or ColValue = dstMask Or ColValue = iPRouteIP Or ColValue = iPRouteMask Then
        'MsgBox dstIP & " exist!", vbExclamation, "Warning"
        flag = False
    End If

    If flag = True Then
        For i = 3 To Worksheets("IPRouteMap").range("a1048576").End(xlUp).row
            If ColValue = Get_dstIP(i) Or ColValue = Get_iPRouteIP(i) Or ColValue = Get_dstMask(i) Or ColValue = Get_iPRouteMask(i) Then
                'MsgBox dstIP & " exist!", vbExclamation, "Warning"
                flag = False
                Exit For
            End If
        Next
    End If
    
    Check_Col = flag
End Function

Function Check_Group(groupValue As String) As Boolean
    Dim flag As Boolean
    flag = False

    For m_rowNum = 2 To Worksheets("MAPPING DEF").range("a1048576").End(xlUp).row
        If groupValue = Worksheets("MAPPING DEF").Cells(m_rowNum, 1).value + "." + Worksheets("MAPPING DEF").Cells(m_rowNum, 2).value Then
            flag = Check_Col(groupValue + "." + Worksheets("MAPPING DEF").Cells(m_rowNum, 3).value)
            If flag Then
                Exit For
            End If
        End If
    Next
    
    Check_Group = flag
End Function

Function Check_Sheet(sheetValue As String) As Boolean
    Dim flag As Boolean
    Dim i As Long
    
    flag = False

    For m_rowNum = 2 To Worksheets("MAPPING DEF").range("a1048576").End(xlUp).row
        If sheetValue = Worksheets("MAPPING DEF").Cells(m_rowNum, 1).value Then
            flag = Check_Group(sheetValue + "." + Worksheets("MAPPING DEF").Cells(m_rowNum, 2).value)
            If flag Then
                Exit For
            End If
        End If
    Next
    
    Check_Sheet = flag
End Function

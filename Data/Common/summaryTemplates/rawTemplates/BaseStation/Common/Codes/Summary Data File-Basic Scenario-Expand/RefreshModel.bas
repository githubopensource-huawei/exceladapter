Attribute VB_Name = "RefreshModel"
Const MapDef_ShetName = "MAPPING DEF"
Const ShetDef_ShetName = "SHEET DEF"
Const MapDef_StartRow = 2
Const ShetTypp_Column = 2
Const DataSheet_StartRow = 3
Const FieldName_Row = 2
Const ShetName_Column = 1
Const ColumnName_Column = 3
Dim TitleMaxColumn As String
Dim titleRange As String
Dim flag As Long



Sub refreshSummary()
    Dim tmpShetName As String
    Dim index As Long
    index = 0
    Application.ScreenUpdating = False
    DisplayMessageOnStatusbar
    If innerPositionMgr Is Nothing Then loadInnerPositions
    
    For index = 2 To Worksheets(ShetDef_ShetName).range("a65536").End(xlUp).row
         tmpShetName = Worksheets(ShetDef_ShetName).Cells(index, innerPositionMgr.sheetDef_shtNameColNo)
         If UCase(Worksheets(ShetDef_ShetName).Cells(index, innerPositionMgr.sheetDef_shtTypeColNo)) = "LIST" And Sheets(tmpShetName).Visible < 0 And Not isTrasnPortSheet(tmpShetName) Then
            
             flag = 0 '��ʼ���FlagΪ0
            
             '1.ѭ����Ҫ������Sheetҳ����ճ���һ��NodeB��¼֮�������������
             Call ClearData_NotFirstNodeB(tmpShetName)
             
             If flag = 1 Then
                 '2.ѭ����Ҫ������Sheetҳ��ɾ����ЩSheetҳ������Ϊ��*NodeB Name������
                 Call DelColumn_ByName(tmpShetName)
      
                 '3.����Sheetҳ����+*NodeB Name���ڡ�MAPPING DEF�����ҵ���һ�����ݣ�ɾ����һ��
                 Call DelRow_ByName(tmpShetName)
             
                 '4.����Sheetҳ���ƣ�ˢ�¡�SHET DEF���а���ShetTypeˢ�³�Pattern
                 Call ModShetType_ByName(tmpShetName)
            End If
         End If
    Next
    '������ɾ���У���Ҫinit ���³�ʼ��Ref�Ļ���
    Call initAddRef
    'Sheets("Base Station Transport Data").Select 'ˢ����ɺ�ص���������ҳ
    EndDisplayMessageOnStatusbar
    Application.ScreenUpdating = True
    
    MsgBox "Refresh Summary Finished."

End Sub

Function is_DelColumnName(columnName As String) As Boolean
    is_DelColumnName = False
    
    If columnName = getResByKey("*NODEB_NAME") Then
        is_DelColumnName = True
    ElseIf columnName = getResByKey("*BTS_NAME") Then
        is_DelColumnName = True
    ElseIf columnName = getResByKey("*BASESTATION_NAME") Then
        is_DelColumnName = True
    ElseIf columnName = getResByKey("*ENODEB_NAME") Then
        is_DelColumnName = True
    ElseIf columnName = getResByKey("*NBBSName") Then
        is_DelColumnName = True
    ElseIf columnName = getResByKey("*ICSNAME") Then
        is_DelColumnName = True
    ElseIf columnName = getResByKey("*RFANAME") Then
        is_DelColumnName = True
    ElseIf columnName = getResByKey("*ENODEBEQMNAME") Then
        is_DelColumnName = True
     ElseIf columnName = getResByKey("*NRNAME") Then
        is_DelColumnName = True
    End If
    
End Function

Function get_NodeBNameColumn(tmpShet As Worksheet) As Long
    
    Dim Column_NodeBName As Long
    Dim iColumn As Long
    Column_NodeBName = 0
    iColumn = 1
    
    Do While (Trim(tmpShet.Cells(FieldName_Row, iColumn)) <> "")
        If is_DelColumnName(Trim(tmpShet.Cells(FieldName_Row, iColumn))) Then
            Column_NodeBName = iColumn
            GoTo Mark
        End If
        iColumn = iColumn + 1
    Loop
    
Mark:
    get_NodeBNameColumn = Column_NodeBName

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

Sub UnMergeTitle(tmpShet As Worksheet)
    TitleMaxColumn = tmpShet.range("IV" + CStr(2)).End(xlToLeft).column
    titleRange = "A1:" + getColStr(TitleMaxColumn) + "1"
    
    
    With tmpShet.range(titleRange)
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    tmpShet.range(titleRange).UnMerge
    
End Sub

Sub MergeTitle(tmpShet As Worksheet)
    If (TitleMaxColumn > 1) Then
        TitleMaxColumn = TitleMaxColumn - 1 'Mergeʱ�Ѿ�ɾ����һ�У���Ҫ��1
        titleRange = "A1:" + getColStr(TitleMaxColumn) + "1"
        
        With tmpShet.range(titleRange)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        tmpShet.range(titleRange).Merge
        
    End If
    
End Sub

Sub DelColumn_ByName(tmpShetName As String)
    Dim tmpSheet As Worksheet
    Set tmpSheet = Sheets(tmpShetName)
    
    Dim Column_NodeBName As Long
    Column_NodeBName = get_NodeBNameColumn(tmpSheet)

    If get_NodeBNameColumn(tmpSheet) > 0 Then
        'Call UnMergeTitle(tmpSheet)
        'ɾ����ǰ�ȱ������
        'tmpSheet.Cells(1, Column_NodeBName + 1) = tmpSheet.Cells(1, Column_NodeBName)
        tmpSheet.columns(Column_NodeBName).Delete Shift:=xlLeft
        ''columns(Column_NodeBName).Select
        ''Selection.Delete Shift:=xlLeft
        'Call MergeTitle(tmpSheet)
    End If
    
End Sub

Sub ClearData_NotFirstNodeB(tmpShetName As String)
    Dim tmpSheet As Worksheet
    Set tmpSheet = Sheets(tmpShetName)
    
    Dim iRow As Long
    iRow = DataSheet_StartRow
    
    Dim firstNodeBName As String
    
    Dim NodeBColumn As Long
    NodeBColumn = get_NodeBNameColumn(tmpSheet)
    
    If NodeBColumn > 0 Then
         flag = 1
         firstNodeBName = GetBluePrintSheetName()
         Do While (tmpSheet.Cells(iRow, NodeBColumn) <> "")
             If (tmpSheet.Cells(iRow, NodeBColumn) <> firstNodeBName) Then
                 tmpSheet.rows(iRow).Delete Shift:=xlUp
                 ''rows(iRow).Select
                 ''Selection.Clear
             Else
                 iRow = iRow + 1
             End If
         Loop
         
     End If

End Sub

Sub DelRow_ByName(tmpShetName As String)
    Dim MapDef_Shet As Worksheet
    Set MapDef_Shet = Sheets(MapDef_ShetName)
    'MapDef_Shet.Visible = True
    
    Dim iRow As Long
    iRow = MapDef_StartRow
    
    Do While (MapDef_Shet.Cells(iRow, ShetName_Column) <> "")
        If ((MapDef_Shet.Cells(iRow, ShetName_Column) = tmpShetName) And is_DelColumnName(Trim(MapDef_Shet.Cells(iRow, ColumnName_Column)))) Then
            MapDef_Shet.rows(iRow).Delete Shift:=xlUp
            'rows(iRow).Select
            'Selection.Delete Shift:=xlUp
        End If
        iRow = iRow + 1
    Loop
    
    'MapDef_Shet.Visible = False
    
End Sub

Sub ModShetType_ByName(tmpShetName As String)
    Dim ShetDef_Shet As Worksheet
    Set ShetDef_Shet = Sheets(ShetDef_ShetName)
    If innerPositionMgr Is Nothing Then loadInnerPositions
    
    Dim iRow As Long
    iRow = MapDef_StartRow
    
    Do While (ShetDef_Shet.Cells(iRow, innerPositionMgr.sheetDef_shtNameColNo) <> "")
        If (ShetDef_Shet.Cells(iRow, innerPositionMgr.sheetDef_shtNameColNo) = tmpShetName) Then
            ShetDef_Shet.Cells(iRow, innerPositionMgr.sheetDef_shtTypeColNo) = "Pattern"
            Exit Sub
        End If
        iRow = iRow + 1
    Loop

End Sub






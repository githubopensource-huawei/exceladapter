Attribute VB_Name = "DF_FreshTemplate"
'*****************************************************************************************************
'模块名称：DF_FreshTemplate
'模块作用：依据TableDef等页面数据刷新DoubleFrequencyCellSetting页面
'*****************************************************************************************************
Option Explicit

'Set template for DoubleFrequencyCellSetting Sheet
Public Sub SetTemplate_DF()
    Call PrepareMocInfos_DF

    Dim sht As Worksheet, r As Range
    Set sht = Sheets(SHT_DOUBLE_FREQ_CELL_SETTING)

    Dim iEnd As Integer
    Dim i As Integer
    iEnd = GetLastRowIndex(sht)
    'For i = 1 To iEnd - DF_ROW_MOC + 1
        'sht.Rows(DF_ROW_MOC).Delete
    'Next i
    
    Dim j As Integer
    Dim Moc As MocClass, Attr As AttrClass
    Dim iRow As Integer, iRow2 As Integer
    iRow = DF_ROW_MOC

    For i = 1 To g_df_colMocNames.Count
        Set Moc = g_c_clsMocInfos.Moc(g_df_colMocNames.Item(i))
        sht.Cells(iRow, DF_COL_MOC).Value = Moc.Name
        iRow2 = iRow 'MOC开始行
        For j = 1 To Moc.Count
            Set Attr = Moc.Attr(j)
            Application.StatusBar = FormatStr(RSC_STR_FRESHING_MOC_ATTR, Moc.Name, Attr.Name)

            Set r = sht.Cells(iRow, DF_COL_ATTR)
            With r
                .Value = Attr.Name
                .ClearComments
                .AddComment "Description Name: " & Chr(10) & Attr.caption
            End With
            r.Interior.ColorIndex = 36

            SetBorderWeight Range:=r, Top:=xlThin, Bottom:=xlThin, Left:=xlThin, Right:=xlThin
           
            Set r = sht.Cells(iRow, DF_COL_ATTR_DFT_VALUE)
            SetBorderWeight Range:=r, Top:=xlThin, Bottom:=xlThin, Left:=xlThin, Right:=xlMedium
            Call SetRangeValidation(r, Attr)
            
            iRow = iRow + 1
        Next j
        
        Set r = sht.Range(sht.Cells(iRow2, DF_COL_MOC), sht.Cells(iRow - 1, DF_COL_MOC))
        r.Merge
        r.Interior.ColorIndex = 36
        SetBorderWeight Range:=r, Top:=xlThin, Bottom:=xlThin, Left:=xlMedium, Right:=xlThin

        'MOC之间留一个空白行
        iRow = iRow + 1
        If i < g_df_colMocNames.Count Then
            Set r = sht.Range(sht.Cells(iRow - 1, DF_COL_MOC), sht.Cells(iRow - 1, DF_COL_ATTR_DFT_VALUE))
            SetBorderWeight Range:=r, Top:=xlThin, Bottom:=xlThin, Left:=xlMedium, Right:=xlMedium
        End If
    Next i
    
     Set r = sht.Range(sht.Cells(iRow - 2, DF_COL_MOC), sht.Cells(iRow - 2, DF_COL_ATTR_DFT_VALUE))
     SetBorderWeight Range:=r, Top:=xlThin, Bottom:=xlMedium, Left:=xlMedium, Right:=xlMedium

     Application.StatusBar = RSC_STR_FINISHED
End Sub

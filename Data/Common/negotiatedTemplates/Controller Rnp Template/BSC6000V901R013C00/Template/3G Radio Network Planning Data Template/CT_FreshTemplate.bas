Attribute VB_Name = "CT_FreshTemplate"
'*****************************************************************************************************
'模块名称：CvtTemplateHelpModule
'模块作用：依据TableDef等页面数据刷新ConvertTemplate页面
'*****************************************************************************************************
Option Explicit

'Set template for ConvertTemplate Sheet
Public Sub SetTemplate_CT()
    Call PrepareMocInfos_CT(Application.ActiveWorkbook)

    g_c_clsMocInfos.Moc(MOC_INTRAFREQNCELL).Remove (ATTR_BIDIRECTION)
    g_c_clsMocInfos.Moc(MOC_INTERFREQNCELL).Remove (ATTR_BIDIRECTION)
    g_c_clsMocInfos.Moc(MOC_GSMNCELL).Remove (ATTR_BIDIRECTION)
    
    Dim sht As Worksheet, r As Range
    Set sht = Sheets(SHT_CONVERT_TEMPLATE)

    Dim i As Integer, j As Integer
    Dim Moc As MocClass, Attr As AttrClass
    Dim iRow As Integer, iMocRowIndex As Integer
    iRow = CT_ROW_MOC

    For i = 1 To g_ct_colMocNames.Count
        Set Moc = g_c_clsMocInfos.Moc(g_ct_colMocNames.Item(i))
        sht.Cells(iRow, CT_COL_MOC_HW).Value = Moc.Name
        iMocRowIndex = iRow 'MOC开始行
        For j = 1 To Moc.CountWithVirtualAttr
            Set Attr = Moc.Attr(j)
            Application.StatusBar = FormatStr(RSC_STR_FRESHING_MOC_ATTR, Moc.Name, Attr.Name)

            Set r = sht.Cells(iRow, CT_COL_ATTR_HW)
            With r
                .Value = Attr.Name
                .ClearComments
                .AddComment "Description Name: " & Chr(10) & Attr.caption
            End With
            If Attr.IsVirtualAttr Then '虚拟属性
                r.Interior.ColorIndex = 6
            Else
                r.Interior.ColorIndex = 36
            End If

            SetBorderWeight Range:=r, Top:=xlThin, Bottom:=xlThin, Left:=xlThin, Right:=xlThin
            
            Set r = sht.Cells(iRow, CT_COL_ATTR_VDF)
            r.Interior.ColorIndex = 35
            SetBorderWeight Range:=r, Top:=xlThin, Bottom:=xlThin, Left:=xlThin, Right:=xlThin
            
            Set r = sht.Cells(iRow, CT_COL_ATTR_DFT_VALUE)
            SetBorderWeight Range:=r, Top:=xlThin, Bottom:=xlThin, Left:=xlThin, Right:=xlMedium
            Call SetRangeValidation(r, Attr)
            
            iRow = iRow + 1
        Next j
        
        Set r = sht.Range(sht.Cells(iMocRowIndex, CT_COL_MOC_HW), sht.Cells(iRow - 1, CT_COL_MOC_HW))
        r.Merge
        r.Interior.ColorIndex = 36
        SetBorderWeight Range:=r, Top:=xlThin, Bottom:=xlThin, Left:=xlMedium, Right:=xlThin
        
        Set r = sht.Range(sht.Cells(iMocRowIndex, CT_COL_MOC_VDF), sht.Cells(iRow - 1, CT_COL_MOC_VDF))
        r.Merge
        r.Interior.ColorIndex = 35
        SetBorderWeight Range:=r, Top:=xlThin, Bottom:=xlThin, Left:=xlThin, Right:=xlThin
        
        'MOC之间留一个空白行
        iRow = iRow + 1
        If i < g_ct_colMocNames.Count Then
            Set r = sht.Range(sht.Cells(iRow - 1, CT_COL_MOC_HW), sht.Cells(iRow - 1, CT_COL_ATTR_DFT_VALUE))
            sht.Cells(iRow - 1, CT_COL_ATTR_HW).Clear
            SetBorderWeight Range:=r, Top:=xlThin, Bottom:=xlThin, Left:=xlMedium, Right:=xlMedium
            r.Interior.ColorIndex = xlNone
            r.Borders(xlInsideVertical).LineStyle = xlNone
        End If
    Next i
    
     Set r = sht.Range(sht.Cells(iRow - 2, CT_COL_MOC_HW), sht.Cells(iRow - 2, CT_COL_ATTR_DFT_VALUE))
     SetBorderWeight Range:=r, Top:=xlThin, Bottom:=xlMedium, Left:=xlMedium, Right:=xlMedium
     
    sht.Range(RNG_FILE_VDF).Value = ""
    sht.Range(RNG_FILE_HW).Value = ""

    Application.StatusBar = RSC_STR_FINISHED
End Sub

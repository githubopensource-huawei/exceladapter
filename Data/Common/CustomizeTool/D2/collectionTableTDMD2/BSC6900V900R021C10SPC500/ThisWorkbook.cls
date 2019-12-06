VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' 激活时添加自定义工具栏
Private Sub Workbook_Activate()
  On Error Resume Next
  init ThisWorkbook
  BuildToolBar
  'ThisWorkbook.Protect Password:=GetSheetsPass, Structure:=True, Windows:=False
  'init ThisWorkbook
End Sub
' 去激活时删除工具栏
Private Sub Workbook_Deactivate()
  On Error Resume Next
  DelToolBar
  'ThisWorkbook.Protect Password:=GetSheetsPass, Structure:=True, Windows:=False
End Sub

Function isTransportSheet(sheetName As String) As Boolean
    If sheetName = "BTS Transport Layer" Then
        isTransportSheet = True
        Exit Function
    End If
    isTransportSheet = False
End Function


'切换站型，若输入的模板名称在该站型的模板列表中没有，则将模板名称列置空
Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)
    Dim sheet As New Worksheet
    Dim sheetName As String
    Dim m_bIsMatch As Boolean
    Dim m_TemplateList As String
    Dim constTypeCol As Integer
    Dim constPattenCol As Integer
    Dim m_Str_Template As String
    
    init ThisWorkbook
    
    Set sheet = sh
    sheetName = sheet.name
    
        Dim copyPasteGuard As New CCopyPasteGuard
    copyPasteGuard.init
        
    If ThisWorkbook.name <> ActiveWorkbook.name Or sheet.Visible <> xlSheetVisible Then
        Exit Sub
    End If
    
    If isTransportSheet(sheetName) = True Then
    
    
        constTypeCol = Get_Col(gShtName_bts, constRecordRow, gColName_btsType)
    
        constPattenCol = Get_Col(gShtName_bts, constRecordRow, gColName_btsTpltName)
    


        If Target.Column = constTypeCol And Target.count = 1 And Target.row > constRecordRow Then
            m_TemplateList = Get_Template_Related(Target.Offset(0, 0).value)
         
                If Target.Offset(0, 0).value <> "" And m_TemplateList <> "" Then
                    m_bIsMatch = IsTypeMatchName(Target.Offset(0, constPattenCol - constTypeCol).value, Target.Offset(0, 0).value)

                    If Not m_bIsMatch Then
                        Target.Offset(0, constPattenCol - constTypeCol).value = ""
                    End If
                End If
      
        End If
    End If

End Sub

Private Sub Workbook_SheetSelectionChange(ByVal sh As Object, ByVal Target As Range)
    Dim sheet As New Worksheet
    Dim sheetName As String
    Dim m_Str_Cabinet As String
    Dim m_Str As String
    Dim constTypeCol As Integer
    Dim constPattenCol As Integer
    Dim m_Str_Template As String
    Dim cur_value As String
    Dim tmp_str_val As String
    
    init ThisWorkbook
    
    Set sheet = sh
    sheetName = sheet.name
    
    If isTransportSheet(sheetName) = True Then
        constTypeCol = Get_Col(gShtName_bts, constRecordRow, gColName_btsType)
    
        constPattenCol = Get_Col(gShtName_bts, constRecordRow, gColName_btsTpltName)

        If constTypeCol = -1 Or constPattenCol = -1 Then
            Exit Sub
        End If
        
        If Not isCustomizedTpl And Target.Column = constPattenCol And Target.count = 1 And Target.row > 2 Then
            sheet.Cells(Target.row, Target.Column).value = ""
        End If
        If Not isCustomizedTpl Then
            Exit Sub
        End If
        
        If Target.Column = constPattenCol And Target.count = 1 And Target.row > constRecordRow Then
            m_Str_Template = Get_Template_Related(Target.Offset(0, constTypeCol - constPattenCol).value)
            
            '判断的当前单元格的值是否在范围内，不在则清空
            cur_value = sheet.Cells(Target.row, Target.Column).value
            tmp_str_val = Get_Template_Related_str(Target.Offset(0, constTypeCol - constPattenCol).value)
            If InStr(tmp_str_val, cur_value) = 0 Then
                sheet.Cells(Target.row, Target.Column).value = ""
            End If
            
            If m_Str_Template <> "" Then
                With Target.Validation
                   .Delete
                   .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, Formula1:=m_Str_Template
                   .IgnoreBlank = True
                   .InCellDropdown = True
                   .InputTitle = ""
                   .ErrorTitle = ""
                   .InputMessage = ""
                   .ErrorMessage = ""
                   .IMEMode = xlIMEModeNoControl
                   .ShowInput = True
                   .ShowError = False
                End With
            Else
                With Target.Validation
                    .Delete
                    .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .InputTitle = ""
                    .ErrorTitle = ""
                    .InputMessage = ""
                    .ErrorMessage = ""
                    .IMEMode = xlIMEModeNoControl
                    .ShowInput = True
                    .ShowError = False
                End With
            End If
        End If
    End If
End Sub

'从「MappingSiteTypeCabinetType」页获取「BTS Type」列侯选值
'Function GetSiteType() As String
'    Dim m_rowNum As Integer
'    Dim m_RowNum_Inner As Integer
'    Dim m_Str As String
    
'    m_Str = MappingSiteTypeCabinetType.Cells(2, 1).Value
'    For m_rowNum = 3 To MappingSiteTypeCabinetType.Range("a65536").End(xlUp).row
'        For m_RowNum_Inner = 2 To m_rowNum - 1
'            If MappingSiteTypeCabinetType.Cells(m_rowNum, 1) <> MappingSiteTypeCabinetType.Cells(m_RowNum_Inner, 1) Then
'                flag = False
'            Else
'                flag = True
'                Exit For
'            End If
'        Next
'        If flag = False Then
'            m_Str = m_Str & "," & MappingSiteTypeCabinetType.Cells(m_rowNum, 1).Value
'        End If
'    Next
'    GetSiteType = m_Str
    
'End Function


'从「MappingSiteTemplate」页获取「BTS Template」列侯选值
Function Get_Template_Related(SiteType As String) As String
    Dim m_rowNum As Integer
    Dim m_Str As String
    Dim m_start As Integer
    Dim m_end As Integer
    m_start = 0

    For m_rowNum = 2 To MappingSiteTemplate.Range("a65536").End(xlUp).row
        If SiteType = MappingSiteTemplate.Cells(m_rowNum, 1).value And Trim(MappingSiteTemplate.Cells(m_rowNum, 2).value) <> "" Then
            If m_start = 0 Then
                m_start = m_rowNum
            End If
            m_end = m_rowNum
        End If
    Next
    
    If m_start = 0 Then
        m_Str = ""
    Else
        m_Str = "=INDIRECT(""MappingSiteTemplate!B" & CStr(m_start) & ":B" & CStr(m_end) & """)"
    End If
    
    Get_Template_Related = m_Str
    
End Function

'从「MappingSiteTemplate」页获取「BTS Template」列侯选值，返回以逗号分隔的列表值字符串
Function Get_Template_Related_str(SiteType As String) As String
    Dim m_rowNum As Integer
    Dim m_Str As String
    Dim m_start As Integer
    Dim m_end As Integer
    Dim m_site_name As String
    Dim temp_val As String
    m_site_name = ""

    For m_rowNum = 2 To MappingSiteTemplate.Range("a65536").End(xlUp).row
        If SiteType = MappingSiteTemplate.Cells(m_rowNum, 1).value And Trim(MappingSiteTemplate.Cells(m_rowNum, 2).value) <> "" Then
            temp_val = MappingSiteTemplate.Cells(m_rowNum, 2).value
            If m_site_name = "" Then
                m_site_name = Trim(temp_val)
            Else
                m_site_name = m_site_name + "," + Trim(temp_val)
            End If
        End If
    Next
    
    Get_Template_Related_str = m_site_name
    
End Function

'从指定sheet页的指定行，查找指定列，返回列号
Function Get_Col(sheetName As String, RecordRow As Integer, ColValue As String) As Integer
    Dim m_ColNum As Integer
    Dim f_flag As Boolean
    f_flag = False
    
    For m_ColNum = 1 To ThisWorkbook.Worksheets(sheetName).Range("IV1").End(xlToLeft).Column
        If ColValue = ThisWorkbook.Worksheets(sheetName).Cells(RecordRow, m_ColNum).value Then
            f_flag = True
            Exit For
        End If
    Next
    
    If f_flag = False Then
      m_ColNum = -1
    End If
    
    Get_Col = m_ColNum
    
End Function
'遍历模板名称是否与站型匹配
Function IsTypeMatchName(TemplateName As String, BTSType As String) As Boolean

    Dim rowscount As Integer
    Dim m_rowNum As Integer
    Dim IsMatch As Boolean
    
    
    IsMatch = False
    
     '当前数据行数
    rowscount = MappingSiteTemplate.Range("a65536").End(xlUp).row
    
    For m_rowNum = 2 To rowscount
        If MappingSiteTemplate.Cells(m_rowNum, 2) = TemplateName _
                And MappingSiteTemplate.Cells(m_rowNum, 1) = BTSType Then
            
            IsMatch = True
            IsTypeMatchName = IsMatch
            
            Exit Function
        End If
    Next

End Function




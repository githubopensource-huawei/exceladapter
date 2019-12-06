Attribute VB_Name = "TransportLayerSub"
Option Explicit


'�л�վ�ͣ��������ģ�������ڸ�վ�͵�ģ���б���û�У���ģ���������ÿ�
Public Sub transportLayer_SheetChange(ByVal Target As Range)
    
    Dim m_bIsMatch As Boolean
    Dim m_TemplateList As String
    
    init ThisWorkbook
    
    Dim constTypeCol As Integer, constPattenCol As Integer
    constTypeCol = Get_Col(gShtName_bts, constRecordRow, gColName_btsType)
    
    constPattenCol = Get_Col(gShtName_bts, constRecordRow, gColName_btsTpltName)
    


    If Target.Column = constTypeCol And Target.Count = 1 And Target.row > constRecordRow Then
         m_TemplateList = Get_Template_Related(Target.Offset(0, 0).Value)
         
         If Target.Offset(0, 0).Value <> "" And m_TemplateList <> "" Then
            m_bIsMatch = IsTypeMatchName(Target.Offset(0, constPattenCol - constTypeCol).Value, Target.Offset(0, 0).Value)

            If Not m_bIsMatch Then
              Target.Offset(0, constPattenCol - constTypeCol).Value = ""
            End If
         End If
      
    End If

End Sub

Public Sub transportLayer_SelectionChange(ByVal Target As Range)

    Dim m_Str_Cabinet As String
    Dim m_Str As String
    
    init ThisWorkbook
    
    Dim constTypeCol As Integer, constPattenCol As Integer
    constTypeCol = Get_Col(gShtName_bts, constRecordRow, gColName_btsType)
    
    constPattenCol = Get_Col(gShtName_bts, constRecordRow, gColName_btsTpltName)

    If constTypeCol = -1 Or constPattenCol = -1 Or Not isCustomizedTpl Then
       Exit Sub
    End If
    
    If Target.Column = constPattenCol And Target.Count = 1 And Target.row > constRecordRow Then
        Dim m_Str_Template As String
        m_Str_Template = Get_Template_Related(Target.Offset(0, constTypeCol - constPattenCol).Value)
        
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
    
End Sub


'�ӡ�MappingSiteTemplate��ҳ��ȡ��BTS Template���к�ѡֵ
Function Get_Template_Related(SiteType As String) As String
    Dim m_rowNum As Integer
    Dim m_Str As String
    Dim m_start As Integer
    Dim m_end As Integer
    m_start = 0
    
    Dim mappingSiteTemplateSht As Worksheet
    Set mappingSiteTemplateSht = ThisWorkbook.Worksheets("MappingSiteTemplate")
    
    For m_rowNum = 2 To mappingSiteTemplateSht.Range("a65536").End(xlUp).row
        If SiteType = mappingSiteTemplateSht.Cells(m_rowNum, 1).Value And Trim(mappingSiteTemplateSht.Cells(m_rowNum, 2).Value) <> "" Then
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

'��ָ��sheetҳ��ָ���У�����ָ���У������к�
Function Get_Col(sheetName As String, RecordRow As Integer, ColValue As String) As Integer
    Dim m_ColNum As Integer
    Dim f_flag As Boolean
    f_flag = False
    
    For m_ColNum = 1 To ThisWorkbook.Worksheets(sheetName).Range("IV1").End(xlToLeft).Column
        If ColValue = ThisWorkbook.Worksheets(sheetName).Cells(RecordRow, m_ColNum).Value Then
            f_flag = True
            Exit For
        End If
    Next
    
    If f_flag = False Then
      m_ColNum = -1
    End If
    
    Get_Col = m_ColNum
    
End Function

'����ģ�������Ƿ���վ��ƥ��
Function IsTypeMatchName(TemplateName As String, BTSType As String) As Boolean

    Dim rowscount As Integer
    Dim m_rowNum As Integer
    Dim IsMatch As Boolean
    
    
    IsMatch = False
    
     '��ǰ��������
    Dim mappingSiteTemplateSht As Worksheet
    Set mappingSiteTemplateSht = ThisWorkbook.Worksheets("MappingSiteTemplate")
    
    rowscount = mappingSiteTemplateSht.Range("a65536").End(xlUp).row
    
    For m_rowNum = 2 To rowscount
        If mappingSiteTemplateSht.Cells(m_rowNum, 2) = TemplateName _
                And mappingSiteTemplateSht.Cells(m_rowNum, 1) = BTSType Then
            
            IsMatch = True
            IsTypeMatchName = IsMatch
            
            Exit Function
        End If
    Next

End Function









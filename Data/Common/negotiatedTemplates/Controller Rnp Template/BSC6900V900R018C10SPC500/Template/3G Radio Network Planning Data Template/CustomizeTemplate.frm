VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomizeTemplate 
   Caption         =   "Customize Template"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6210
   OleObjectBlob   =   "CustomizeTemplate.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "CustomizeTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btCancel_Click()
    CustomizeTemplate.Hide
End Sub

Private Sub btOK_Click()
    Worksheets("NODEB").Visible = cbNODEB.Value
    Worksheets("CELL").Visible = cbCELL.Value
    
    Worksheets("NRNCCELL").Visible = cbUMTS.Value
    Worksheets("INTRAFREQNCELL").Visible = cbUMTS.Value
    Worksheets("INTERFREQNCELL").Visible = cbUMTS.Value
    
    Worksheets("GSMCELL").Visible = cbGSM.Value
    Worksheets("GSMNCELL").Visible = cbGSM.Value
    
    Worksheets("LTECELL").Visible = cbLTE.Value
    Worksheets("LTENCELL").Visible = cbLTE.Value
        
    Worksheets("USMLCEXT3GCELL").Visible = cbUSMLCEXT3GCELL.Value
    Worksheets("PhyNBRadio").Visible = cbPhyNBRadio.Value
    
    CustomizeTemplate.Hide
    
End Sub
Public Sub initCheckBox(iLanguageType As Integer)
    
    'if the value is assigned dirtectly, the color of v in the checkbox is grey. The following code is a worakaruond for the unkown problem
    If Not Worksheets("NODEB").Visible Then
        cbNODEB.Value = False
    End If
    
    If Not Worksheets("CELL").Visible Then
        cbCELL.Value = False
    End If
    
    If Not (Worksheets("NRNCCELL").Visible Or Worksheets("INTRAFREQNCELL").Visible Or Worksheets("INTERFREQNCELL").Visible) Then
        cbUMTS.Value = False
    End If
    
    
    If Not (Worksheets("GSMCELL").Visible Or Worksheets("GSMNCELL").Visible) Then
        cbGSM.Value = False
    End If
    
    If Not (Worksheets("LTECELL").Visible Or Worksheets("LTENCELL").Visible) Then
        cbLTE.Value = False
    End If
    
    If Not Worksheets("USMLCEXT3GCELL").Visible Then
        cbUSMLCEXT3GCELL.Value = False
    End If
    
    If Not Worksheets("PhyNBRadio").Visible Then
        cbPhyNBRadio.Value = False
    End If
    
    If iLanguageType = 0 Then
    
        cbGSM.caption = "GSM Neighboring Cell Realted"
        cbUMTS.caption = "UMTS Neighboring Cell Realted"
        cbLTE.caption = "LTE Neighboring Cell Realted"
        
        CustomizeTemplate.caption = "Customize Template"
        CustomizeTemplate.Summary.caption = "Category"
        
        btOK.caption = "OK"
        btCancel.caption = "Cancel"
    Else
        
        cbGSM.caption = "GSM相邻小区相关对象"
        cbUMTS.caption = "UMTS相邻小区相关对象"
        cbLTE.caption = "LTE相邻小区相关对象"
        
        CustomizeTemplate.caption = "定制模板"
        CustomizeTemplate.Summary.caption = "类别"
        
        btOK.caption = "确定"
        btCancel.caption = "取消"
    
    End If
    
    
End Sub



Private Sub UserForm_Click()

End Sub

Attribute VB_Name = "Cover_Sheet"
Private Sub cmBts_Click()
    TemplateCMForm.Show vbModeless
End Sub



Public Sub ShowSheet()
  ThisWorkbook.Unprotect
  
  Call HideAllSheets
  Call ShowTDMSheets(cbTDM.value)
  Call ShowcbTDMDXXSheets(cbTDMDXX.value)
  Call ShowIPOESheets(cbIPOE.value)
  Call ShowIPFESheets(cbIPFE.value)
  Call ShowIPFEandE1T1(cbIPOEandFE.value)
  
  ThisWorkbook.Protect Structure:=True, Windows:=False
End Sub


Public Sub cbIPFE_Click()
    Call ShowSheet
End Sub

Public Sub cbIPOE_Click()
    Call ShowSheet
End Sub

Public Sub cbIPOEandFE_Click()
 Call ShowSheet
End Sub

Public Sub cbTDM_Click()
    Call ShowSheet
End Sub

Public Sub cbTDMDXX_Click()
    Call ShowSheet
End Sub



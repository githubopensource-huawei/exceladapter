VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TemplateCMForm 
   Caption         =   " IuIur Template Configuration"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   OleObjectBlob   =   "TemplateCMForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "TemplateCMForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public CheckReport As String

Private Sub OKBtn_Click()
    If CheckBeforeHide Then
    
        ProcessSheetName = "IUCS"
        Set TemplateSheet = Sheets(ProcessSheetName)
        Call ResetTemplate
        Call HideIUCSCommonMOLine
        Call HideIUCSATMMOLine
        Call HideIUCSIPMOLine
        Call ReProtectTemplate
        
        ProcessSheetName = "IUPS"
        Set TemplateSheet = Sheets(ProcessSheetName)
        Call ResetTemplate
        Call HideIUPSCommonMOLine
        Call HideIUPSATMMOLine
        Call HideIUPSIPMOLine
        Call ReProtectTemplate
        
        ProcessSheetName = "IUR"
        Set TemplateSheet = Sheets(ProcessSheetName)
        Call ResetTemplate
        Call HideIURCommonMOLine
        Call HideIURATMMOLine
        Call HideIURIPMOLine
        Call ReProtectTemplate
        
        TemplateCMForm.Hide
        Sheets(CurrentSheetName).Select
    End If
End Sub

Private Function CheckBeforeHide()
  CheckBeforeHide = True
  CheckReport = ""
  
  ProcessSheetName = "IUCS"
  Set TemplateSheet = Sheets(ProcessSheetName)
  CheckIUCSCommonMOLine
  CheckIUCSATMMOLine
  CheckIUCSIPMOLine
  
  ProcessSheetName = "IUPS"
  Set TemplateSheet = Sheets(ProcessSheetName)
  CheckIUPSCommonMOLine
  CheckIUPSATMMOLine
  CheckIUPSIPMOLine
  
  ProcessSheetName = "IUR"
  Set TemplateSheet = Sheets(ProcessSheetName)
  CheckIURCommonMOLine
  CheckIURATMMOLine
  CheckIURIPMOLine
  If CheckReport <> "" Then
     TemplateSheet.Select
     MsgBox "Can not change template because of " & vbCrLf + CheckReport
     CheckBeforeHide = False
  End If
End Function

Private Sub ResetBtn_Click()
    ProcessSheetName = "IUCS"
    Set TemplateSheet = Sheets(ProcessSheetName)
    Call ResetTemplate
     
    ProcessSheetName = "IUPS"
    Set TemplateSheet = Sheets(ProcessSheetName)
    Call ResetTemplate
     
    ProcessSheetName = "IUR"
    Set TemplateSheet = Sheets(ProcessSheetName)
    Call ResetTemplate
    Sheets(CurrentSheetName).Select
End Sub
Private Sub CancelBtn_Click()
    TemplateCMForm.Hide
    Sheets(CurrentSheetName).Select
End Sub

Private Sub HideIUCSCommonMOLine()
    Call HideLine("CNDOMAIN", IUCSCNDOMAINComboBox.Text)
    Call HideLine("N7DPC", IUCSN7DPCComboBox.Text)
    Call HideLine("CNNODE", IUCSCNNODEComboBox.Text)
    Call HideLine("ADJNODE", IUCSAdjNodeComboBox.Text)
    Call HideLine("ADJMAP", IUCSAdjMapComboBox.Text)
End Sub
 
Private Sub HideIUCSATMMOLine()
    If IUCSATMOptionButton.Value Then
        Call HideLine("MTP3LKS", IUCSMTP3LKSComboBox.Text)
        Call HideLine("MTP3LNK", IUCSMTP3LNKComboBox.Text)
        Call HideLine("MTP3RT", IUCSMTP3RTComboBox.Text)
        Call HideLine("AAL2RT", IUCSAAL2RTComboBox.Text)
        Call HideLine("AAL2PATH", IUCSAAL2PathComboBox.Text)
    Else
        Call HideLine("MTP3LKS", 0)
        Call HideLine("MTP3LNK", 0)
        Call HideLine("MTP3RT", 0)
        Call HideLine("AAL2RT", 0)
        Call HideLine("AAL2PATH", 0)
    End If
End Sub
 
Private Sub HideIUCSIPMOLine()
    If IUCSIPOptionButton.Value Then
        Call HideLine("M3LKS", IUCSM3LKSComboBox.Text)
        Call HideLine("M3LNK", IUCSM3LNKComboBox.Text)
        Call HideLine("M3RT", IUCSM3RTComboBox.Text)
        Call HideLine("IPPATH", IUCSIPPathComBoBox.Text)
    Else
        Call HideLine("M3LKS", 0)
        Call HideLine("M3LNK", 0)
        Call HideLine("M3RT", 0)
        Call HideLine("IPPATH", 0)
    End If
End Sub

Private Sub HideIUPSCommonMOLine()
    Call HideLine("CNDOMAIN", IUPSCNDOMAINComboBox.Text)
    Call HideLine("N7DPC", IUPSN7DPCComboBox.Text)
    Call HideLine("CNNODE", IUPSCNNODEComboBox.Text)
    Call HideLine("ADJNODE", IUPSADJNODEComboBox.Text)
    Call HideLine("ADJMAP", IUPSADJMAPComboBox.Text)
    Call HideLine("IPPATH", IUPSIPPATHComboBox.Text)
End Sub
 
Private Sub HideIUPSATMMOLine()
    If IUPSATMOptionButton.Value Then
        Call HideLine("MTP3LKS", IUPSMTP3LKSComboBox.Text)
        Call HideLine("MTP3LNK", IUPSMTP3LNKComboBox.Text)
        Call HideLine("MTP3RT", IUPSMTP3RTComboBox.Text)
        Call HideLine("IPOAPVC", IUPSIPOAPVCComboBox.Text)
    Else
        Call HideLine("MTP3LKS", 0)
        Call HideLine("MTP3LNK", 0)
        Call HideLine("MTP3RT", 0)
        Call HideLine("IPOAPVC", 0)
    End If
End Sub
 
Private Sub HideIUPSIPMOLine()
    If IUPSIPOptionButton.Value Then
        Call HideLine("M3LKS", IUPSM3LKSComboBox.Text)
        Call HideLine("M3LNK", IUPSM3LNKComboBox.Text)
        Call HideLine("M3RT", IUPSM3RTComboBox.Text)
        'Call HideLine("IPPATH", IUPSIPPATHComboBox.Text)
    Else
        Call HideLine("M3LKS", 0)
        Call HideLine("M3LNK", 0)
        Call HideLine("M3RT", 0)
        'Call HideLine("IPPATH", 0)
    End If
End Sub

 Private Sub HideIURCommonMOLine()
    Call HideLine("N7DPC", IURN7DPCComboBox.Text)
    Call HideLine("NRNC", IURNRNCComboBox.Text)
    Call HideLine("ADJNODE", IURADJNODEComboBox.Text)
    Call HideLine("ADJMAP", IURADJMAPComboBox.Text)
End Sub
 
Private Sub HideIURATMMOLine()
    If IURATMOptionButton.Value Then
        Call HideLine("MTP3LKS", IURMTP3LKSComboBox.Text)
        Call HideLine("MTP3LNK", IURMTP3LNKComboBox.Text)
        Call HideLine("MTP3RT", IURMTP3RTComboBox.Text)
        Call HideLine("AAL2RT", IURAAL2RTComboBox.Text)
        Call HideLine("AAL2PATH", IURAAL2PATHComboBox.Text)
    Else
        Call HideLine("MTP3LKS", 0)
        Call HideLine("MTP3LNK", 0)
        Call HideLine("MTP3RT", 0)
        Call HideLine("AAL2RT", 0)
        Call HideLine("AAL2PATH", 0)
    End If
End Sub
 
Private Sub HideIURIPMOLine()
    If IURIPOptionButton.Value Then
        Call HideLine("M3LKS", IURM3LKSComboBox.Text)
        Call HideLine("M3LNK", IURM3LNKComboBox.Text)
        Call HideLine("M3RT", IURM3RTComboBox.Text)
        Call HideLine("IPPATH", IURIPPATHComboBox.Text)
    Else
        Call HideLine("M3LKS", 0)
        Call HideLine("M3LNK", 0)
        Call HideLine("M3RT", 0)
        Call HideLine("IPPATH", 0)
    End If
End Sub
Private Sub IUCSATMOptionButton_Click()
    If IUCSATMOptionButton.Value Then
       Call EnableIUCSATMItem
       Call DisableIUCSIPItem
    End If
End Sub

Private Sub IUCSIPOptionButton_Click()
    If IUCSIPOptionButton.Value Then
       Call EnableIUCSIPItem
       Call DisableIUCSATMItem
    End If
End Sub

Private Sub IUPSATMOptionButton_Click()
    If IUPSATMOptionButton.Value Then
       Call EnableIUPSATMItem
       Call DisableIUPSIPItem
    End If
End Sub

Private Sub IUPSIPOptionButton_Click()
    If IUPSIPOptionButton.Value Then
       Call EnableIUPSIPItem
       Call DisableIUPSATMItem
    End If
End Sub

Private Sub IURATMOptionButton_Click()
    If IURATMOptionButton.Value Then
       Call EnableIURATMItem
       Call DisableIURIPItem
    End If
End Sub

Private Sub IURIPOptionButton_Click()
    If IURIPOptionButton.Value Then
       Call EnableIURIPItem
       Call DisableIURATMItem
    End If
End Sub

Private Sub EnableIUCSATMItem()
    IUCSMTP3LKSComboBox.Enabled = True
    IUCSMTP3LNKComboBox.Enabled = True
    IUCSMTP3RTComboBox.Enabled = True
    IUCSAAL2RTComboBox.Enabled = True
    IUCSAAL2PathComboBox.Enabled = True
End Sub
 
Private Sub DisableIUCSATMItem()
    IUCSMTP3LKSComboBox.Enabled = False
    IUCSMTP3LNKComboBox.Enabled = False
    IUCSMTP3RTComboBox.Enabled = False
    IUCSAAL2RTComboBox.Enabled = False
    IUCSAAL2PathComboBox.Enabled = False
End Sub
 
Private Sub EnableIUCSIPItem()
    IUCSM3LKSComboBox.Enabled = True
    IUCSM3LNKComboBox.Enabled = True
    IUCSM3RTComboBox.Enabled = True
    IUCSIPPathComBoBox.Enabled = True
End Sub
 
Private Sub DisableIUCSIPItem()
    IUCSM3LKSComboBox.Enabled = False
    IUCSM3LNKComboBox.Enabled = False
    IUCSM3RTComboBox.Enabled = False
    IUCSIPPathComBoBox.Enabled = False
 End Sub
 
 Private Sub EnableIUPSATMItem()
    IUPSMTP3LKSComboBox.Enabled = True
    IUPSMTP3LNKComboBox.Enabled = True
    IUPSMTP3RTComboBox.Enabled = True
    IUPSIPOAPVCComboBox.Enabled = True
End Sub
 
Private Sub DisableIUPSATMItem()
    IUPSMTP3LKSComboBox.Enabled = False
    IUPSMTP3LNKComboBox.Enabled = False
    IUPSMTP3RTComboBox.Enabled = False
    IUPSIPOAPVCComboBox.Enabled = False
End Sub

Private Sub EnableIUPSIPItem()
    IUPSM3LKSComboBox.Enabled = True
    IUPSM3LNKComboBox.Enabled = True
    IUPSM3RTComboBox.Enabled = True
    'IUPSIPPATHComboBox.Enabled = True
End Sub
 
Private Sub DisableIUPSIPItem()
    IUPSM3LKSComboBox.Enabled = False
    IUPSM3LNKComboBox.Enabled = False
    IUPSM3RTComboBox.Enabled = False
    'IUPSIPPATHComboBox.Enabled = False
 End Sub

Private Sub EnableIURATMItem()
    IURMTP3LKSComboBox.Enabled = True
    IURMTP3LNKComboBox.Enabled = True
    IURMTP3RTComboBox.Enabled = True
    IURAAL2RTComboBox.Enabled = True
    IURAAL2PATHComboBox.Enabled = True
End Sub
 
Private Sub DisableIURATMItem()
    IURMTP3LKSComboBox.Enabled = False
    IURMTP3LNKComboBox.Enabled = False
    IURMTP3RTComboBox.Enabled = False
    IURAAL2RTComboBox.Enabled = False
    IURAAL2PATHComboBox.Enabled = False
End Sub
 
Private Sub EnableIURIPItem()
    IURM3LKSComboBox.Enabled = True
    IURM3LNKComboBox.Enabled = True
    IURM3RTComboBox.Enabled = True
    IURIPPATHComboBox.Enabled = True
End Sub
 
Private Sub DisableIURIPItem()
    IURM3LKSComboBox.Enabled = False
    IURM3LNKComboBox.Enabled = False
    IURM3RTComboBox.Enabled = False
    IURIPPATHComboBox.Enabled = False
End Sub

 Private Sub CheckIUCSCommonMOLine()
    Call CheckIfExistData("CNDOMAIN", IUCSCNDOMAINComboBox.Text, CheckReport)
    Call CheckIfExistData("N7DPC", IUCSN7DPCComboBox.Text, CheckReport)
    Call CheckIfExistData("CNNODE", IUCSCNNODEComboBox.Text, CheckReport)
    Call CheckIfExistData("ADJNODE", IUCSAdjNodeComboBox.Text, CheckReport)
    Call CheckIfExistData("ADJMAP", IUCSAdjMapComboBox.Text, CheckReport)
 End Sub
 
 Private Sub CheckIUCSATMMOLine()
   If IUCSIPOptionButton.Value Then
      Call CheckIfExistData("MTP3LKS", IUCSMTP3LKSComboBox.Text, CheckReport)
      Call CheckIfExistData("MTP3LNK", IUCSMTP3LNKComboBox.Text, CheckReport)
      Call CheckIfExistData("AAL2PATH", IUCSAAL2PathComboBox.Text, CheckReport)
      Call CheckIfExistData("AAL2RT", IUCSAAL2RTComboBox.Text, CheckReport)
      Call CheckIfExistData("MTP3RT", IUCSMTP3RTComboBox.Text, CheckReport)
   Else
      Call CheckIfExistData("MTP3LKS", 0, CheckReport)
      Call CheckIfExistData("MTP3LNK", 0, CheckReport)
      Call CheckIfExistData("AAL2PATH", 0, CheckReport)
      Call CheckIfExistData("AAL2RT", 0, CheckReport)
      Call CheckIfExistData("MTP3RT", 0, CheckReport)
   End If
 End Sub
Private Sub CheckIUCSIPMOLine()
   If IUCSIPOptionButton.Value Then
      Call CheckIfExistData("M3LKS", IUCSM3LKSComboBox.Text, CheckReport)
      Call CheckIfExistData("M3LNK", IUCSM3LNKComboBox.Text, CheckReport)
      Call CheckIfExistData("M3RT", IUCSM3RTComboBox.Text, CheckReport)
      Call CheckIfExistData("IPPATH", IUCSIPPathComBoBox.Text, CheckReport)
   Else
      Call CheckIfExistData("M3LKS", 0, CheckReport)
      Call CheckIfExistData("M3LNK", 0, CheckReport)
      Call CheckIfExistData("M3RT", 0, CheckReport)
      Call CheckIfExistData("IPPATH", 0, CheckReport)
   End If
 End Sub
 
 Private Sub CheckIUPSCommonMOLine()
    Call CheckIfExistData("CNDOMAIN", IUPSCNDOMAINComboBox.Text, CheckReport)
    Call CheckIfExistData("N7DPC", IUPSN7DPCComboBox.Text, CheckReport)
    Call CheckIfExistData("CNNODE", IUPSCNNODEComboBox.Text, CheckReport)
    Call CheckIfExistData("ADJNODE", IUPSADJNODEComboBox.Text, CheckReport)
    Call CheckIfExistData("ADJMAP", IUPSADJMAPComboBox.Text, CheckReport)
    Call CheckIfExistData("IPPATH", IUPSIPPATHComboBox.Text, CheckReport)
 End Sub
 
 Private Sub CheckIUPSATMMOLine()
   If IUPSIPOptionButton.Value Then
      Call CheckIfExistData("MTP3LKS", IUPSMTP3LKSComboBox.Text, CheckReport)
      Call CheckIfExistData("MTP3LNK", IUPSMTP3LNKComboBox.Text, CheckReport)
      Call CheckIfExistData("IPOAPVC", IUPSIPOAPVCComboBox.Text, CheckReport)
      Call CheckIfExistData("MTP3RT", IUPSMTP3RTComboBox.Text, CheckReport)
   Else
      Call CheckIfExistData("MTP3LKS", 0, CheckReport)
      Call CheckIfExistData("MTP3LNK", 0, CheckReport)
      Call CheckIfExistData("IPOAPVC", 0, CheckReport)
      Call CheckIfExistData("MTP3RT", 0, CheckReport)
   End If
 End Sub

Private Sub CheckIUPSIPMOLine()
   If IUPSIPOptionButton.Value Then
      Call CheckIfExistData("M3LKS", IUPSM3LKSComboBox.Text, CheckReport)
      Call CheckIfExistData("M3LNK", IUPSM3LNKComboBox.Text, CheckReport)
      Call CheckIfExistData("M3RT", IUPSM3RTComboBox.Text, CheckReport)
      'Call CheckIfExistData("IPPATH", IUPSIPPATHComboBox.Text, CheckReport)
   Else
      Call CheckIfExistData("M3LKS", 0, CheckReport)
      Call CheckIfExistData("M3LNK", 0, CheckReport)
      Call CheckIfExistData("M3RT", 0, CheckReport)
      'Call CheckIfExistData("IPPATH", 0, CheckReport)
   End If
 End Sub

Private Sub CheckIURCommonMOLine()
    Call CheckIfExistData("N7DPC", IURN7DPCComboBox.Text, CheckReport)
    Call CheckIfExistData("NRNC", IURNRNCComboBox.Text, CheckReport)
    Call CheckIfExistData("ADJNODE", IURADJNODEComboBox.Text, CheckReport)
    Call CheckIfExistData("ADJMAP", IURADJMAPComboBox.Text, CheckReport)
 End Sub
 
 Private Sub CheckIURATMMOLine()
   If IURIPOptionButton.Value Then
      Call CheckIfExistData("MTP3LKS", IURMTP3LKSComboBox.Text, CheckReport)
      Call CheckIfExistData("MTP3LNK", IURMTP3LNKComboBox.Text, CheckReport)
      Call CheckIfExistData("AAL2PATH", IURAAL2PATHComboBox.Text, CheckReport)
      Call CheckIfExistData("AAL2RT", IURAAL2RTComboBox.Text, CheckReport)
      Call CheckIfExistData("MTP3RT", IURMTP3RTComboBox.Text, CheckReport)
   Else
      Call CheckIfExistData("MTP3LKS", 0, CheckReport)
      Call CheckIfExistData("MTP3LNK", 0, CheckReport)
      Call CheckIfExistData("AAL2PATH", 0, CheckReport)
      Call CheckIfExistData("AAL2RT", 0, CheckReport)
      Call CheckIfExistData("MTP3RT", 0, CheckReport)
   End If
 End Sub

Private Sub CheckIURIPMOLine()
   If IURIPOptionButton.Value Then
      Call CheckIfExistData("M3LKS", IURM3LKSComboBox.Text, CheckReport)
      Call CheckIfExistData("M3LNK", IURM3LNKComboBox.Text, CheckReport)
      Call CheckIfExistData("M3RT", IURM3RTComboBox.Text, CheckReport)
      Call CheckIfExistData("IPPATH", IURIPPATHComboBox.Text, CheckReport)
   Else
      Call CheckIfExistData("M3LKS", 0, CheckReport)
      Call CheckIfExistData("M3LNK", 0, CheckReport)
      Call CheckIfExistData("M3RT", 0, CheckReport)
      Call CheckIfExistData("IPPATH", 0, CheckReport)
   End If
 End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BatchDeleteFreqForm 
   Caption         =   "Batch Delete Frequency"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7950
   OleObjectBlob   =   "BatchDeleteFreqForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "BatchDeleteFreqForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub CommitButton_Click()
    BatchDelTrxMain (FileNamesList)
    Unload Me
End Sub

Private Sub SelectFileButton_Click()
    If ImportDataFile = False Then
       MsgBox getResByKey("LoadfileFail")
       Exit Sub
    Else
       Src_Path.text = FileNamesList
    End If
End Sub

Private Sub UserForm_Initialize()
    Call Upt_Desc
    FileNamesList = ""
End Sub

Private Sub Upt_Desc()
    BatchDeleteFreqForm.Caption = getResByKey("BatchDeleteFreqForm.Caption")
    LabelFilePath.Caption = getResByKey("LabelFilePath.Caption")
    LabelSelectFile.Caption = getResByKey("LabelSelectFile.Caption")
    CommitButton.Caption = getResByKey("ImportButton.Caption")
    SelectFileButton.Caption = getResByKey("SelectFileButton.Caption")
End Sub

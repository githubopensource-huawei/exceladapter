Attribute VB_Name = "T_CovertTemplate"
Option Explicit

Private Const COVERT_DATA = "CovertDataBar"

Public Sub cmdSelectFileNameForVDF_Click()
    If SelectFileName(0) Then 'VDF
          Call createConvertDataBar
    End If
End Sub

Public Sub cmdSelectFileNameForHW_Click()
    If SelectFileName(1) Then  'HW
          Call createConvertDataBar
    End If
End Sub

Public Sub cmdConvertData_Click()
    Call ConvertData
End Sub

Sub createConvertDataBar()
    If existToolBar(COVERT_DATA) Then
        Application.CommandBars(COVERT_DATA).Delete
    End If
    
    Dim baseStationChooseBar As CommandBar
    Dim baseStationStyle As CommandBarButton
    
    Set baseStationChooseBar = Application.CommandBars.Add(COVERT_DATA, msoBarTop)
    With baseStationChooseBar
        .Protection = msoBarNoResize
        .Visible = True
        Set baseStationStyle = .Controls.Add(Type:=msoControlButton)
        With baseStationStyle
            .Style = msoButtonIconAndCaption
            .caption = "Covert Data"
            .TooltipText = "CovertData"
            .OnAction = "cmdConvertData_Click"
            .FaceId = 50
            .Enabled = True
        End With
    End With
End Sub

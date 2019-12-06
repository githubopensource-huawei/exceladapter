Attribute VB_Name = "CreateBar"
Option Explicit

Public Const CustomTemplateBarName As String = "CustomTemplate Bar"


Public Sub InsertUserToolBar()
    Dim customTemplateBar As CommandBar
    
    If Not containsAToolBar(CustomTemplateBarName) Then
        Set customTemplateBar = Application.CommandBars.Add(CustomTemplateBarName, msoBarTop)
        With customTemplateBar
            .Protection = msoBarNoResize
            .Visible = True
            With .Controls.Add(Type:=msoControlButton)
                .Style = msoButtonIconAndCaption
                .Caption = getResByKey("Bar_CustomTemplate")
                .TooltipText = getResByKey("Bar_CustomTemplate")
                .OnAction = "customTemplate"
                .FaceId = 186
            End With
        End With
    End If
    
End Sub

Sub customTemplate()
    Dim language As String
    language = getResByKey("Language")
    If language = "EN" Then
        Call CustomizeXls(CONST_LAN_EN)
    Else
        Call CustomizeXls(CONST_LAN_ZH)
    End If
End Sub


Public Sub DeleteUserToolBar()
    If containsAToolBar(CustomTemplateBarName) Then
        Application.CommandBars(CustomTemplateBarName).Delete
    End If
End Sub


Public Function containsAToolBar(ByRef barName As String) As Boolean
    On Error GoTo ErrorHandler
    containsAToolBar = True
    Dim bar As CommandBar
    Set bar = CommandBars(barName)
    Exit Function
ErrorHandler:
    containsAToolBar = False
End Function


Sub DestroyMenuStatus()
    With Application
        .CommandBars("Row").Reset
        .CommandBars("Column").Reset
        .CommandBars("Cell").Reset
        .CommandBars("Ply").Reset
    End With
End Sub

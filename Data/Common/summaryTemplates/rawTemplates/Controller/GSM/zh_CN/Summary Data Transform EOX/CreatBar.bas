Attribute VB_Name = "CreatBar"
Option Explicit

Public Const OperationBarName As String = "Operation Bar"


Public Sub InsertUserToolBar()
    Dim toolbar As CommandBar
    
    If Not containsAToolBar(OperationBarName) Then
        Set toolbar = Application.CommandBars.Add(OperationBarName, msoBarTop)
        With toolbar
            .Protection = msoBarNoResize
            .Visible = True
            With .Controls.Add(Type:=msoControlButton)
                .Caption = getResByKey("Bar_Template")
                .OnAction = "addTemplate" '增删小区模板
                .Style = msoButtonIconAndCaption
                .Enabled = True
                .FaceId = 186
            End With
        End With
    End If
End Sub

Public Sub DeleteUserToolBar()
    If containsAToolBar(OperationBarName) Then
        Application.CommandBars(OperationBarName).Delete
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








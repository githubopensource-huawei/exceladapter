Attribute VB_Name = "CreatBarModule"
Option Explicit

Private Const UniversalBar = "Universal Bar"

Private Type UserBar
    toolBarName As String
    captionName As String
    actionMacroName As String
End Type

Private userBarArr(100) As UserBar
Private numberUserBar As Integer

Public Sub InsertUserToolBar()
    numberUserBar = 0
    Dim newBar As UserBar
    Dim toolBarName As String, captionName As String, actionMacroName As String
    
    '修改toolBarName、captionName、actionMacroName为需要的名字，关闭再重新打开表格即可看到加载项宏按钮
    newBar.toolBarName = UniversalBar '宏按钮定义名称
    newBar.captionName = getResByKey("Bar_AddComments") '宏按钮显示名称
    newBar.actionMacroName = "addAllComments" '宏按钮绑定的函数名
    Call addNewBarToArr(newBar)
    
    Call addAllToolBars
End Sub

Private Function containsAToolBar(ByRef barName As String) As Boolean
    On Error GoTo ErrorHandler
    containsAToolBar = True
    Dim bar As CommandBar
    Set bar = CommandBars(barName)
    Exit Function
ErrorHandler:
    containsAToolBar = False
End Function

Private Sub addToolBar(ByRef toolBarName As String, ByRef captionName As String, ByRef actionMacroName As String)
    Dim toolbar As CommandBar
    If Not containsAToolBar(toolBarName) Then
        Set toolbar = Application.CommandBars.Add(toolBarName, msoBarTop)
        With toolbar
            .Protection = msoBarNoResize
            .Visible = True
            'Add User Define Template
            With .Controls.Add(Type:=msoControlButton)
                .Caption = captionName
                .TooltipText = captionName
                .OnAction = actionMacroName
                .Style = msoButtonIconAndCaption
                .Enabled = True
                .FaceId = 186
            End With
        End With
    End If
End Sub

Private Sub addNewBarToArr(ByRef newBar As UserBar)
    userBarArr(numberUserBar) = newBar
    numberUserBar = numberUserBar + 1
End Sub

Private Sub addAllToolBars()
    Dim k As Integer
    Dim eachUserBar As UserBar
    For k = 0 To numberUserBar - 1
        eachUserBar = userBarArr(k)
        Call addToolBar(eachUserBar.toolBarName, eachUserBar.captionName, eachUserBar.actionMacroName)
    Next
End Sub

Public Sub HideUserToolBar()
    Dim k As Integer
    Dim eachUserBar As UserBar
    Dim toolBarName As String
    For k = 0 To numberUserBar - 1
        eachUserBar = userBarArr(k)
        toolBarName = eachUserBar.toolBarName
        If containsAToolBar(toolBarName) Then
            Application.CommandBars(toolBarName).Protection = msoBarNoResize
            Application.CommandBars(toolBarName).Visible = False
        End If
    Next
End Sub

Public Sub DeleteUserToolBar()
    Dim k As Integer
    Dim eachUserBar As UserBar
    Dim toolBarName As String
    For k = 0 To numberUserBar - 1
        eachUserBar = userBarArr(k)
        toolBarName = eachUserBar.toolBarName
        If containsAToolBar(toolBarName) Then
            Application.CommandBars(toolBarName).Delete
        End If
    Next
End Sub

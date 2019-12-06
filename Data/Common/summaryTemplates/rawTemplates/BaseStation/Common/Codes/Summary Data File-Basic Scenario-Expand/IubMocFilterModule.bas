Attribute VB_Name = "IubMocFilterModule"
Option Explicit

Public iubMocManager As CIubMocManager '管理所有IUB页签MOC的管理类
Private nextStepFlag As Boolean '在选定MOC窗体界面是否进行下一步的标志

Public Sub initIubMocManager()
    Set iubMocManager = New CIubMocManager
    Call iubMocManager.init
End Sub

'展示窗体供客户选择，进行MOC是否选定的操作
Public Function getIubUnselectMocCollection(ByRef iubUnselectMocCollection As Collection) As Boolean
    Call displayMocFilterForm
    Set iubUnselectMocCollection = iubMocManager.getUnselectedMocCollection
    '如果选择了取消，则不进行后续转换
    getIubUnselectMocCollection = nextStepFlag
End Function

Public Sub setNextStepFlag(ByRef flag As Boolean)
    nextStepFlag = flag
End Sub

Public Function getNextStepFlag()
    getNextStepFlag = nextStepFlag
End Function

Public Sub displayMocFilterForm()
    MocFilterForm.Show
End Sub
'
'Private Sub test()
'    Call displayMocFilterForm
'End Sub



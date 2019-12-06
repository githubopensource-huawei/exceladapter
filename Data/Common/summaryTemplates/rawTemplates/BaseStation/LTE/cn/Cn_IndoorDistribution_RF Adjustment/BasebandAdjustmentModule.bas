Attribute VB_Name = "BasebandAdjustmentModule"
Option Explicit
Private Const BasebandEqmMocName As String = "BASEBANDEQM" '基带设备MOC名称
Public Const BasebandReferenceBoardNoDelimeter As String = ";" '处理单板编号的分隔符

Private basebandBoardStyleMappingDefData As CBoardStyleMappingDefData '基带设备的类对象
Public basebandEqmIdManager As CBaseBandEqmIdManager '基带设备编号和单编号列表的管理类

'调整基带处理单板编号窗体
Public Sub AdjustBasebandEqmBoardNo()
    BasebandEqmAdjustmentForm.Show
End Sub

Public Function initCurrentSheetBaseBandEqmIdManager() As Boolean
    Dim flag As Boolean
    flag = True
    '得到BASEBANDEQM的组名
    Dim baseBandEqmGroupName As String
    baseBandEqmGroupName = boardStyleData.getGroupNamebyMocName(BasebandEqmMocName)
    
    '如果没有基带设备对象，则提示并退出
    If baseBandEqmGroupName = "" Then
        Call MsgBox(getResByKey("NoBasebandEqmObject"), vbInformation + vbOKOnly, getResByKey("ErrorInfo"))
        flag = False
        Exit Function
    End If
    
    If boardStyleMappingDefMap Is Nothing Then
        Set boardStyleMappingDefMap = boardStyleData.getBoardStyleMappingDefMap
    End If
    
    '获取BaseBand类对象
    Set basebandBoardStyleMappingDefData = boardStyleMappingDefMap.GetAt(baseBandEqmGroupName)
    
    '得到BaseBand对象的起始行和结束行
    Dim startRow As Long, endRow As Long
    Call getGroupNameStartAndEndRowNumber(baseBandEqmGroupName, startRow, endRow)
    
    '初始化编号管理类
    Set basebandEqmIdManager = New CBaseBandEqmIdManager
    '如果EqmId列或处理单板编号没有找到，则退出
    If basebandEqmIdManager.init(startRow, endRow, basebandBoardStyleMappingDefData) = False Then
        flag = False
        Exit Function
    End If
    
    initCurrentSheetBaseBandEqmIdManager = flag
End Function

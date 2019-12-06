VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBaseBandEqmIdManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Delimeter_ As String = ";"
Private Const EqmIdTypeDelimter_ As String = "_"

Private Const BasebandEqmIdAttributeName_ As String = "BASEBANDEQMID" '基带设备编号属性名称，后面用来查找ColumnName
Private basebandEqmIdColummName_ As String 'BASEBANDEQMID对应的组名
Private basebandEqmIdColumnLetter_ As String 'BASEBANDEQMID所在的列号

Private Const BasebandEqmTypeAttributeName_ As String = "BASEBANDEQMTYPE" '基带设备类型短名称，后面用来查找ColumnName
Private basebandEqmTypeColumnName_ As String 'BASEBANDEQMTYPE对应的组名，后面要用它的值和EqmId拼接起来作为唯一的字段
Private basebandEqmTypeColumnLetter_ As String 'BASEBANDEQMTYPE对应的列号

Private Const BasebandEqmBoardNoAttributeName_ As String = "BasebandeqmBoardNo" '基带处理单板编号 属性名称，后面用来查找ColumnName
Private basebandEqmBoardNoColummName_ As String 'BasebandeqmBoardNo对应的组名
Private basebandEqmBoardNoColumnLetter_ As String 'BasebandeqmBoardNo所在的列号


Private basebandEqmIdToBoardNoColMap_ As CMapValueObject '基带设备编号值与基带设备类型拼接的字符串 到 引用单板编号字符串 的映射，如0_UL-0_0_1_0;0_0_1_2

Private basebandBoardStyleMappingDefData_ As CBoardStyleMappingDefData '基带设备对象类

Private startRowNumber_ As Long 'BASEBANDEQM对象起始行
Private endRowNumber_ As Long 'BASEBANDEQM对象结束行

Private allReferenceBoardNoCol_ As Collection '所有可引用的单板编号容器，key-value都是单板编号值


Public Function init(ByRef startRowNumber As Long, ByRef endRowNumber As Long, Optional ByRef basebandBoardStyleMappingDefData As CBoardStyleMappingDefData) As Boolean
    init = True
    startRowNumber_ = startRowNumber
    endRowNumber_ = endRowNumber
    
    If Not basebandBoardStyleMappingDefData Is Nothing Then
        Set basebandBoardStyleMappingDefData_ = basebandBoardStyleMappingDefData
        '得到BASEBANDEQMID和BasebandeqmBoardNo的列名和列号
        Call getColumnNameAndColumnLetterByAttributeName(basebandBoardStyleMappingDefData_, BasebandEqmIdAttributeName_, basebandEqmIdColummName_, basebandEqmIdColumnLetter_)
        Call getColumnNameAndColumnLetterByAttributeName(basebandBoardStyleMappingDefData_, BasebandEqmBoardNoAttributeName_, basebandEqmBoardNoColummName_, basebandEqmBoardNoColumnLetter_)
        Call getColumnNameAndColumnLetterByAttributeName(basebandBoardStyleMappingDefData_, BasebandEqmTypeAttributeName_, basebandEqmTypeColumnName_, basebandEqmTypeColumnLetter_)
    End If
    
    '如果没有找到EqmId或处理单板编号列，则退出
    If basebandEqmIdColummName_ = "" Or basebandEqmBoardNoColummName_ = "" Or basebandEqmBoardNoColummName_ = "" Then
        Call MsgBox(getResByKey("EqmIdOrBoardNoNotFound"), vbInformation + vbOKOnly, getResByKey("ErrorInfo"))
        init = False
        Exit Function
    End If
    '生成EqmId到BoardNoCol的列表
    If makeBasebandEqmIdToBoardNoColMap = False Then
        init = False
        Exit Function
    End If
    
    '把所有可引用的的单板编号生成在allReferenceBoardNoCol_
    Call makeAllReferencedCol
End Function

'把所有可引用的单板编号字符串获取
Private Sub makeAllReferencedCol()
    Dim referencedString As String
    referencedString = basebandBoardStyleMappingDefData_.getMultiListReferenceValue(basebandEqmBoardNoColummName_)
    If referencedString = "" Then
        Exit Sub
    End If
    referencedString = getReferencedStringByInputString(referencedString)
    
    Set allReferenceBoardNoCol_ = New Collection
    '把得到所有可引用的单板编号值分隔生成容器
    Set allReferenceBoardNoCol_ = getKeyValueCollectionFromString(referencedString, ",")
End Sub

Private Function makeBasebandEqmIdToBoardNoColMap() As Boolean
    makeBasebandEqmIdToBoardNoColMap = True
    Dim rowIndex As Long
    Dim eachBasebandEqmId As String, eachReferenceBoardNoString As String, eachBasebandEqmType As String, key As String
    Set basebandEqmIdToBoardNoColMap_ = New CMapValueObject
    
    Dim referenceBoardNoCol As Collection
    '分别得到每一行的BASEBANDEQMID和BasebandeqmBoardNo
    For rowIndex = startRowNumber_ + 2 To endRowNumber_
        eachBasebandEqmId = currentSheet.range(basebandEqmIdColumnLetter_ & rowIndex).value
        eachBasebandEqmType = currentSheet.range(basebandEqmTypeColumnLetter_ & rowIndex).value
        key = getKey(eachBasebandEqmId, eachBasebandEqmType) '将EqmId和EqmType拼接出来的字符串作为Key值
        If basebandEqmIdToBoardNoColMap_.hasKey(key) Then
            Call MsgBox(getResByKey("SameEqmId"), vbInformation + vbOKOnly, getResByKey("ErrorInfo"))
            makeBasebandEqmIdToBoardNoColMap = False
            Exit Function
        End If
        eachReferenceBoardNoString = currentSheet.range(basebandEqmBoardNoColumnLetter_ & rowIndex).value
        If eachBasebandEqmId <> "" Then
            '把BASEBANDEQMID和BasebandeqmBoardNo的映射加入
            Set referenceBoardNoCol = getKeyValueCollectionFromString(eachReferenceBoardNoString, Delimeter_)
            Call basebandEqmIdToBoardNoColMap_.SetAt(key, referenceBoardNoCol)
        End If
    Next rowIndex
End Function

Private Function getKey(ByRef eqmId As String, ByRef eqmType As String) As String
    If eqmType <> "" Then
        getKey = eqmId & EqmIdTypeDelimter_ & eqmType
    Else
        getKey = eqmId
    End If
End Function

'得到基带设备编号容器
Public Function getBasebandEqmIdCol() As Collection
    Set getBasebandEqmIdCol = basebandEqmIdToBoardNoColMap_.KeyCollection
End Function

'根据一个属性名称得到这个参数的显示名称和列号
Private Sub getColumnNameAndColumnLetterByAttributeName(ByRef boardStyleMappingDefData As CBoardStyleMappingDefData, _
    ByRef attributeName As String, ByRef columnName As String, ByRef columnLetter As String)
    columnName = boardStyleMappingDefData.getColumnNameByAttributeName(attributeName)
    If columnName = "" Then
        columnLetter = ""
        Exit Sub
    End If
    columnLetter = boardStyleMappingDefData.getColumnNamePositionLetter(columnName)
End Sub

'增删一个EqmId对应的引用单板编号容器
Public Sub modifyReferencedBoardNoOfEqmId(ByRef eqmId As String, ByRef referencedBoardNo As String, ByRef addOrDelete As String)
    Dim referencedBoardNoCol As Collection
    
    If Not basebandEqmIdToBoardNoColMap_.hasKey(eqmId) Then Exit Sub '如果没有这个eqmId，则退出
    Set referencedBoardNoCol = basebandEqmIdToBoardNoColMap_.GetAt(eqmId)
    
    If addOrDelete = "+" Then
        '新增一个BoardNo
        If Not Contains(referencedBoardNoCol, referencedBoardNo) Then
           referencedBoardNoCol.Add Item:=referencedBoardNo, key:=referencedBoardNo
        End If
    ElseIf addOrDelete = "-" Then
        '删除一个BoardNo
        If Contains(referencedBoardNoCol, referencedBoardNo) Then
            referencedBoardNoCol.Remove (referencedBoardNo)
        End If
    End If
    
    '把增删过的容器重新设置给对应的eqmId
    Call basebandEqmIdToBoardNoColMap_.SetAt(eqmId, referencedBoardNoCol)
End Sub

'传入一个EqmId，得到它剩余的可选引用单板编号容器和当前的引用单板编号容器
Public Sub getOptionalAndCurrentBoardNoCols(ByRef eqmId As String, ByRef optionalCol As Collection, ByRef currentBoardNoCol As Collection)
    If Not basebandEqmIdToBoardNoColMap_.hasKey(eqmId) Then Exit Sub '没有这个编号的EqmId，则退出
    
    Set currentBoardNoCol = basebandEqmIdToBoardNoColMap_.GetAt(eqmId)
    
    '遍历所有可引用的单板编号，如果不在当前已引用的boardNo容器中，则加入optionalCol中，供客户选择
    Dim eachBoardNo As Variant
    For Each eachBoardNo In allReferenceBoardNoCol_
        If Not Contains(currentBoardNoCol, CStr(eachBoardNo)) Then
            optionalCol.Add Item:=eachBoardNo, key:=eachBoardNo
        End If
    Next eachBoardNo
End Sub

Public Sub writeNewBasebandEqmBoardNo()
    Dim rowIndex As Long
    Dim eachEqmId As String, eachEqmType As String, key As String
    Dim eachReferenceBoardNoCol As Collection
    Dim newBoardNoString As String
    
    For rowIndex = startRowNumber_ + 2 To endRowNumber_
        '得到每一行的EqmId值
        eachEqmId = currentSheet.range(basebandEqmIdColumnLetter_ & rowIndex).value
        eachEqmType = currentSheet.range(basebandEqmTypeColumnLetter_ & rowIndex).value
        key = getKey(eachEqmId, eachEqmType)
        If eachEqmId = "" Or (Not basebandEqmIdToBoardNoColMap_.hasKey(key)) Then GoTo NextLoop
        
        '得到对应的BoardNo列表容器
        Set eachReferenceBoardNoCol = basebandEqmIdToBoardNoColMap_.GetAt(key)
        
        '得到拼接的BoardNo字符串，写入基带处理单板编号列中
        newBoardNoString = getConnectedStringFromCol(eachReferenceBoardNoCol, Delimeter_)
        currentSheet.range(basebandEqmBoardNoColumnLetter_ & rowIndex).value = newBoardNoString
NextLoop:
    Next rowIndex
    
    '调整行高
    currentSheet.rows(startRowNumber_ + 2 & ":" & endRowNumber_).AutoFit
End Sub

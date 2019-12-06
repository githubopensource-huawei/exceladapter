VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBTSNameManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private IubMocSelectedColumn_ As String '在SheetDef中选定的MOC列号
'Private Const IubMocSelectedColumnName_ As String = "Selected" '在SheetDef中选定的MOC列号
'
'Private iubMocMap_ As CMapValueObject '管理所有IUB页签对象的Map，key为mocName，value为CIubMoc
'Private iubSelectedMocCollection_ As Collection '记录所有选定的MOC，key/value都为mocName;未选定的MOC每次需要时再获取
Private iubUnSelectedMocCollection_ As Collection '记录所有未选定的MOC，key/value都为mocName;选定的MOC每次需要时再获取
'Private sheetDefSheet_ As Worksheet 'SHEET DEF页签

Private baseStationCollection As Collection '所有基站名称
'



Public Sub init()
    Dim RowNum As Long
    Dim maxRow As String
    Dim nowSelection As range
    Dim index As Long

    Dim baseStationName As Variant
    Dim selectBtsName As String

    
    Set nowSelection = Selection
    Dim btsNameColIndex As Long
    Dim mocName As String
    Dim attrName As String
    Call getBaseStationMocNameAndAttrName(mocName, attrName)
    btsNameColIndex = getColNum(ActiveSheet.name, 2, attrName, mocName)
    selectBtsName = ActiveSheet.Cells(nowSelection.row, btsNameColIndex).value
    
    maxRow = ActiveSheet.range("b1048576").End(xlUp).row
    For RowNum = 3 To maxRow
        baseStationName = ActiveSheet.Cells(RowNum, btsNameColIndex).value

        If existInCollection(baseStationName, baseStationCollection) = False And Trim(baseStationName) <> "" Then
            baseStationCollection.Add CStr(baseStationName)
        End If
    Next
    
    For Each baseStationName In baseStationCollection
        If Trim(baseStationName) <> "" Then
            iubUnSelectedMocCollection_.Add CStr(baseStationName)
        End If
    Next
End Sub


Public Sub getCollections(ByRef selectedMocCollection As Collection, ByRef unselectedMocCollection As Collection)
    Set selectedMocCollection = getSelectedMocCollection
    Set unselectedMocCollection = iubUnSelectedMocCollection_
End Sub

Public Function getUnselectedMocCollection() As Collection
    Set getUnselectedMocCollection = iubUnSelectedMocCollection_
End Function

Public Function getSelectedMocCollection() As Collection
    Dim iubSelectedMocCollection As New Collection
    Dim mocName As Variant
    For Each mocName In baseStationCollection
        If Not existInCollection(CStr(mocName), iubUnSelectedMocCollection_) Then
            iubSelectedMocCollection.Add Item:=mocName, key:=mocName
        End If
    Next mocName
    Set getSelectedMocCollection = iubSelectedMocCollection
End Function

'将一个MOC从未选定容器中加入到选定容器中
Public Function addMocToSelected(ByRef mocName As String)
    Dim index As Long
    
    If existInCollection(mocName, iubUnSelectedMocCollection_) Then
        index = getIndexInCollection(mocName, iubUnSelectedMocCollection_)
        iubUnSelectedMocCollection_.Remove (index + 1)
    End If

End Function

Private Function getIndexInCollection(strValue As Variant, strCollection As Collection) As Long
    Dim sItem As Variant
    Dim index As Long
    index = 0
    For Each sItem In strCollection
        If sItem = strValue Then
            getIndexInCollection = index
            Exit Function
        End If
        index = index + 1
    Next
    getIndexInCollection = 0
End Function

'将一个MOC从选定容器中加入到未选定容器中
Public Function addMocToUnselected(ByRef mocName As String)
    If Not existInCollection(mocName, iubUnSelectedMocCollection_) Then
        iubUnSelectedMocCollection_.Add Item:=mocName, key:=mocName
    End If
End Function

Private Function existInCollection(strValue As Variant, strCollection As Collection) As Boolean
    Dim sItem As Variant
    If Trim(CStr(strValue)) = "" Then
        existInCollection = True
        Exit Function
    End If
    For Each sItem In strCollection
        If sItem = strValue Then
            existInCollection = True
            Exit Function
        End If
    Next
    existInCollection = False
End Function

'把所有未选定的MOC的标志"N"设置到Sheet Def页签中
'Public Sub setUnselectedMocFlag()
'    Dim mocName As Variant
'    Dim iubMoc As CIubMoc
'    Dim dstRowNumber As Long
'    For Each mocName In iubMocMap_.KeyCollection
'        Set iubMoc = iubMocMap_.GetAt(mocName)
'        dstRowNumber = iubMoc.getRowNumber
'
'        If Not Contains(iubUnSelectedMocCollection_, CStr(mocName)) Then
'            '选定则清空该标志
'            sheetDefSheet_.range(IubMocSelectedColumn_ & dstRowNumber).ClearContents
'        Else
'            '未选定则将该标志位置为N
'            sheetDefSheet_.range(IubMocSelectedColumn_ & dstRowNumber).value = "N"
'        End If
'    Next mocName
'End Sub
'
''得到一个moc的起始行和结束行
'Public Sub getMocRows(ByRef mocName As String, ByRef startRow As Long, ByRef endRow As Long)
'    startRow = -1
'    endRow = -1
'    If Not iubMocMap_.haskey(mocName) Then Exit Sub '如果没有此moc，直接退出，一般不会出现此情况
'    Dim iubMoc As CIubMoc
'    Set iubMoc = iubMocMap_.GetAt(mocName)
'    startRow = iubMoc.startRow
'    endRow = iubMoc.endRow
'End Sub
'
''将F列的名称标为Selected
'Private Sub initSelectedColumnName()
'    With sheetDefSheet_.range(IubMocSelectedColumn_ & "1")
'        If .value = "" Then
'            .value = IubMocSelectedColumnName_
'        End If
'    End With
'End Sub
'
'Private Sub initIubMoc(ByRef mocName As String, ByRef selectedString As String, ByRef rowNumber As Long, ByRef startRow As Long, ByRef endRow As Long)
'    Dim iubMoc As New CIubMoc '声明一个IubMoc的类对象，记录该对象名称，是否选定，在SheetDef中所有行等信息
'    Call iubMoc.init(mocName, selectedString, rowNumber, startRow, endRow)
'
'    '将该MOC放入Map中，用mocName为key
'    Call iubMocMap_.SetAt(mocName, iubMoc)
'
'    '根据该Moc是否选择，放入相应的选择和未选择的容器中
'    If Not iubMoc.isMocSelected Then
'        iubUnSelectedMocCollection_.Add Item:=mocName, key:=mocName
'    End If
'End Sub


Private Sub Class_Initialize()
'    Set iubMocMap_ = New CMapValueObject
    Set iubUnSelectedMocCollection_ = New Collection
'    Set sheetDefSheet_ = ThisWorkbook.Worksheets("SHEET DEF")
'    IubMocSelectedColumn_ = getColStr(sheetDefSheet_.range("XFD2").End(xlToLeft).column + 1)
    Set baseStationCollection = New Collection
End Sub

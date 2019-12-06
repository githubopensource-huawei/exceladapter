VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIubMocManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private IubMocSelectedColumn_ As String  '在SheetDef中选定的MOC列号
Private Const IubMocSelectedColumnName_ As String = "Selected" '在SheetDef中选定的MOC列号

Private iubMocMap_ As CMapValueObject '管理所有IUB页签对象的Map，key为mocName，value为CIubMoc
'Private iubSelectedMocCollection_ As Collection '记录所有选定的MOC，key/value都为mocName;未选定的MOC每次需要时再获取
Private iubUnSelectedMocCollection_ As Collection '记录所有未选定的MOC，key/value都为mocName;选定的MOC每次需要时再获取
Private sheetDefSheet_ As Worksheet 'SHEET DEF页签

Public Sub init()
    Call initSelectedColumnName
    
    Dim rowNumber As Long
    Dim mocName As String, selectedString As String, startRowValue As String, endRowValue As String
    With sheetDefSheet_
        For rowNumber = 2 To .range("A65536").End(xlUp).row
            mocName = .Cells(rowNumber, innerPositionMgr.sheetDef_shtNameColNo).value
            startRowValue = CStr(.Cells(rowNumber, innerPositionMgr.sheetDef_startRowColNo).value)
            endRowValue = CStr(.Cells(rowNumber, innerPositionMgr.sheetDef_endRowColNo).value)
            selectedString = CStr(.Cells(rowNumber, IubMocSelectedColumn_).value)
            
            If startRowValue <> "" Then
                Call initIubMoc(mocName, selectedString, rowNumber, CLng(startRowValue), CLng(endRowValue))
            End If
        Next rowNumber
    End With
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
    For Each mocName In iubMocMap_.KeyCollection
        If Not Contains(iubUnSelectedMocCollection_, CStr(mocName)) Then
            iubSelectedMocCollection.Add Item:=mocName, key:=mocName
        End If
    Next mocName
    Set getSelectedMocCollection = iubSelectedMocCollection
End Function

'将一个MOC从未选定容器中加入到选定容器中
Public Function addMocToSelected(ByRef mocName As String)
    If Contains(iubUnSelectedMocCollection_, mocName) Then
        iubUnSelectedMocCollection_.Remove (mocName)
    End If
End Function

'将一个MOC从选定容器中加入到未选定容器中
Public Function addMocToUnselected(ByRef mocName As String)
    If Not Contains(iubUnSelectedMocCollection_, mocName) Then
        iubUnSelectedMocCollection_.Add Item:=mocName, key:=mocName
    End If
End Function

'把所有未选定的MOC的标志"N"设置到Sheet Def页签中
Public Sub setUnselectedMocFlag()
    Dim mocName As Variant
    Dim iubMoc As CIubMoc
    Dim dstRowNumber As Long
    For Each mocName In iubMocMap_.KeyCollection
        Set iubMoc = iubMocMap_.GetAt(mocName)
        dstRowNumber = iubMoc.getRowNumber
        
        If Not Contains(iubUnSelectedMocCollection_, CStr(mocName)) Then
            '选定则清空该标志
            sheetDefSheet_.range(IubMocSelectedColumn_ & dstRowNumber).ClearContents
        Else
            '未选定则将该标志位置为N
            sheetDefSheet_.range(IubMocSelectedColumn_ & dstRowNumber).value = "N"
        End If
    Next mocName
End Sub

'得到一个moc的起始行和结束行
Public Sub getMocRows(ByRef mocName As String, ByRef startRow As Long, ByRef endRow As Long)
    startRow = -1
    endRow = -1
    If Not iubMocMap_.hasKey(mocName) Then Exit Sub '如果没有此moc，直接退出，一般不会出现此情况
    Dim iubMoc As CIubMoc
    Set iubMoc = iubMocMap_.GetAt(mocName)
    startRow = iubMoc.startRow
    endRow = iubMoc.endRow
End Sub

Private Sub initSelectedColumnName()
    Dim maxCol As Long
    maxCol = sheetDefSheet_.range("IV" & functionShtTitleRow).End(xlToLeft).column
    IubMocSelectedColumn_ = getColStr(maxCol + 1)
    
    sheetDefSheet_.range(IubMocSelectedColumn_ & functionShtTitleRow).value = IubMocSelectedColumnName_
    If innerPositionMgr Is Nothing Then
        loadInnerPositions
    Else
        innerPositionMgr.updateSheetDefPositions
    End If
End Sub

Private Sub initIubMoc(ByRef mocName As String, ByRef selectedString As String, ByRef rowNumber As Long, ByRef startRow As Long, ByRef endRow As Long)
    Dim iubMoc As New CIubMoc '声明一个IubMoc的类对象，记录该对象名称，是否选定，在SheetDef中所有行等信息
    Call iubMoc.init(mocName, selectedString, rowNumber, startRow, endRow)
    
    '将该MOC放入Map中，用mocName为key
    Call iubMocMap_.SetAt(mocName, iubMoc)
    
    '根据该Moc是否选择，放入相应的选择和未选择的容器中
    If Not iubMoc.isMocSelected Then
        iubUnSelectedMocCollection_.Add Item:=mocName, key:=mocName
    End If
End Sub


Private Sub Class_Initialize()
    Set iubMocMap_ = New CMapValueObject
    Set iubUnSelectedMocCollection_ = New Collection
    Set sheetDefSheet_ = ThisWorkbook.Worksheets("SHEET DEF")
End Sub

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

'Private IubMocSelectedColumn_ As String '��SheetDef��ѡ����MOC�к�
'Private Const IubMocSelectedColumnName_ As String = "Selected" '��SheetDef��ѡ����MOC�к�
'
'Private iubMocMap_ As CMapValueObject '��������IUBҳǩ�����Map��keyΪmocName��valueΪCIubMoc
'Private iubSelectedMocCollection_ As Collection '��¼����ѡ����MOC��key/value��ΪmocName;δѡ����MOCÿ����Ҫʱ�ٻ�ȡ
Private iubUnSelectedMocCollection_ As Collection '��¼����δѡ����MOC��key/value��ΪmocName;ѡ����MOCÿ����Ҫʱ�ٻ�ȡ
'Private sheetDefSheet_ As Worksheet 'SHEET DEFҳǩ

Private baseStationCollection As Collection '���л�վ����
'



Public Sub init()
'    Call initSelectedColumnName
    
'    Dim rowNumber As Long
'    Dim mocName As String, selectedString As String, startRowValue As String, endRowValue As String
'    For rowNumber = 2 To sheetDefSheet_.range("A1048576").End(xlUp).row
'        mocName = sheetDefSheet_.range("A" & rowNumber).value
'        startRowValue = CStr(sheetDefSheet_.range("D" & rowNumber).value)
'        endRowValue = CStr(sheetDefSheet_.range("E" & rowNumber).value)
'        selectedString = CStr(sheetDefSheet_.range(IubMocSelectedColumn_ & rowNumber).value)
'
'        If startRowValue <> "" Then
'            Call initIubMoc(mocName, selectedString, rowNumber, CLng(startRowValue), CLng(endRowValue))
'        End If
'    Next rowNumber

    Dim rowNum As Long
    Dim maxRow As String
    Dim nowSelection As range
    Dim index As Long
'    Dim baseStationCollection As Collection
    Dim baseStationName As Variant
    Dim selectBtsName As String
'    Set baseStationCollection = New Collection
'    BaseStationList.Clear
    
    Set nowSelection = Selection
    selectBtsName = ActiveSheet.Cells(nowSelection.row, 1).value
    If IsGBTSTemplate() Then
        If isOperationExcel Then
            selectBtsName = ActiveSheet.Cells(nowSelection.row, getGTRXBTSNameCol(ActiveSheet.name)).value
        Else
            selectBtsName = ActiveSheet.Cells(nowSelection.row, getGcellBTSNameCol(ActiveSheet.name)).value
        End If
    End If
    
    Dim startColnumLetter As Long
    startColnumLetter = 1
    If isOperationExcel Then
        startColnumLetter = 2
    End If
    
    If CELL_TYPE = 0 Or CELL_TYPE = 4 Then
        baseStationCollection.Add (selectBtsName)
    Else
        maxRow = ActiveSheet.range("b1048576").End(xlUp).row
        For rowNum = 3 To maxRow
            baseStationName = ActiveSheet.Cells(rowNum, startColnumLetter).value
            '@gbts
            If IsGBTSTemplate() Then
                If isOperationExcel Then
                    baseStationName = ActiveSheet.Cells(rowNum, getGTRXBTSNameCol(ActiveSheet.name)).value
                Else
                    baseStationName = ActiveSheet.Cells(rowNum, getGcellBTSNameCol(ActiveSheet.name)).value
                End If
            End If
    
            If existInCollection(baseStationName, baseStationCollection) = False And Trim(baseStationName) <> "" Then
                baseStationCollection.Add CStr(baseStationName)
            End If
        Next
    End If
    
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

'��һ��MOC��δѡ�������м��뵽ѡ��������
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

'��һ��MOC��ѡ�������м��뵽δѡ��������
Public Function addMocToUnselected(ByRef mocName As String)
    If Not existInCollection(mocName, iubUnSelectedMocCollection_) Then
        iubUnSelectedMocCollection_.Add Item:=mocName, key:=mocName
    End If
End Function

'������δѡ����MOC�ı�־"N"���õ�Sheet Defҳǩ��
'Public Sub setUnselectedMocFlag()
'    Dim mocName As Variant
'    Dim iubMoc As CIubMoc
'    Dim dstRowNumber As Long
'    For Each mocName In iubMocMap_.KeyCollection
'        Set iubMoc = iubMocMap_.GetAt(mocName)
'        dstRowNumber = iubMoc.getRowNumber
'
'        If Not Contains(iubUnSelectedMocCollection_, CStr(mocName)) Then
'            'ѡ������ոñ�־
'            sheetDefSheet_.range(IubMocSelectedColumn_ & dstRowNumber).ClearContents
'        Else
'            'δѡ���򽫸ñ�־λ��ΪN
'            sheetDefSheet_.range(IubMocSelectedColumn_ & dstRowNumber).value = "N"
'        End If
'    Next mocName
'End Sub
'
''�õ�һ��moc����ʼ�кͽ�����
'Public Sub getMocRows(ByRef mocName As String, ByRef startRow As Long, ByRef endRow As Long)
'    startRow = -1
'    endRow = -1
'    If Not iubMocMap_.haskey(mocName) Then Exit Sub '���û�д�moc��ֱ���˳���һ�㲻����ִ����
'    Dim iubMoc As CIubMoc
'    Set iubMoc = iubMocMap_.GetAt(mocName)
'    startRow = iubMoc.startRow
'    endRow = iubMoc.endRow
'End Sub
'
''��F�е����Ʊ�ΪSelected
'Private Sub initSelectedColumnName()
'    With sheetDefSheet_.range(IubMocSelectedColumn_ & "1")
'        If .value = "" Then
'            .value = IubMocSelectedColumnName_
'        End If
'    End With
'End Sub
'
'Private Sub initIubMoc(ByRef mocName As String, ByRef selectedString As String, ByRef rowNumber As Long, ByRef startRow As Long, ByRef endRow As Long)
'    Dim iubMoc As New CIubMoc '����һ��IubMoc������󣬼�¼�ö������ƣ��Ƿ�ѡ������SheetDef�������е���Ϣ
'    Call iubMoc.init(mocName, selectedString, rowNumber, startRow, endRow)
'
'    '����MOC����Map�У���mocNameΪkey
'    Call iubMocMap_.SetAt(mocName, iubMoc)
'
'    '���ݸ�Moc�Ƿ�ѡ�񣬷�����Ӧ��ѡ���δѡ���������
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
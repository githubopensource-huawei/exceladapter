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

Private IubMocSelectedColumn_ As String  '��SheetDef��ѡ����MOC�к�
Private Const IubMocSelectedColumnName_ As String = "Selected" '��SheetDef��ѡ����MOC�к�

Private iubMocMap_ As CMapValueObject '��������IUBҳǩ�����Map��keyΪmocName��valueΪCIubMoc
'Private iubSelectedMocCollection_ As Collection '��¼����ѡ����MOC��key/value��ΪmocName;δѡ����MOCÿ����Ҫʱ�ٻ�ȡ
Private iubUnSelectedMocCollection_ As Collection '��¼����δѡ����MOC��key/value��ΪmocName;ѡ����MOCÿ����Ҫʱ�ٻ�ȡ
Private sheetDefSheet_ As Worksheet 'SHEET DEFҳǩ

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

'��һ��MOC��δѡ�������м��뵽ѡ��������
Public Function addMocToSelected(ByRef mocName As String)
    If Contains(iubUnSelectedMocCollection_, mocName) Then
        iubUnSelectedMocCollection_.Remove (mocName)
    End If
End Function

'��һ��MOC��ѡ�������м��뵽δѡ��������
Public Function addMocToUnselected(ByRef mocName As String)
    If Not Contains(iubUnSelectedMocCollection_, mocName) Then
        iubUnSelectedMocCollection_.Add Item:=mocName, key:=mocName
    End If
End Function

'������δѡ����MOC�ı�־"N"���õ�Sheet Defҳǩ��
Public Sub setUnselectedMocFlag()
    Dim mocName As Variant
    Dim iubMoc As CIubMoc
    Dim dstRowNumber As Long
    For Each mocName In iubMocMap_.KeyCollection
        Set iubMoc = iubMocMap_.GetAt(mocName)
        dstRowNumber = iubMoc.getRowNumber
        
        If Not Contains(iubUnSelectedMocCollection_, CStr(mocName)) Then
            'ѡ������ոñ�־
            sheetDefSheet_.range(IubMocSelectedColumn_ & dstRowNumber).ClearContents
        Else
            'δѡ���򽫸ñ�־λ��ΪN
            sheetDefSheet_.range(IubMocSelectedColumn_ & dstRowNumber).value = "N"
        End If
    Next mocName
End Sub

'�õ�һ��moc����ʼ�кͽ�����
Public Sub getMocRows(ByRef mocName As String, ByRef startRow As Long, ByRef endRow As Long)
    startRow = -1
    endRow = -1
    If Not iubMocMap_.hasKey(mocName) Then Exit Sub '���û�д�moc��ֱ���˳���һ�㲻����ִ����
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
    Dim iubMoc As New CIubMoc '����һ��IubMoc������󣬼�¼�ö������ƣ��Ƿ�ѡ������SheetDef�������е���Ϣ
    Call iubMoc.init(mocName, selectedString, rowNumber, startRow, endRow)
    
    '����MOC����Map�У���mocNameΪkey
    Call iubMocMap_.SetAt(mocName, iubMoc)
    
    '���ݸ�Moc�Ƿ�ѡ�񣬷�����Ӧ��ѡ���δѡ���������
    If Not iubMoc.isMocSelected Then
        iubUnSelectedMocCollection_.Add Item:=mocName, key:=mocName
    End If
End Sub


Private Sub Class_Initialize()
    Set iubMocMap_ = New CMapValueObject
    Set iubUnSelectedMocCollection_ = New Collection
    Set sheetDefSheet_ = ThisWorkbook.Worksheets("SHEET DEF")
End Sub
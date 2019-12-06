VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MuliBtsFilterForm 
   Caption         =   "ѡ���վ"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12300
   OleObjectBlob   =   "MuliBtsFilterForm.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "MuliBtsFilterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Option Explicit

Private Sub AddCommandButton_Click()
    Dim eachListIndex As Long
    Dim mocName As String
    With Me.UnselectedBTSNameListBox
        '���û�п���ӵģ���ֱ���˳�
        If .ListCount = 0 Then Exit Sub
        For eachListIndex = 0 To .ListCount - 1
            '���α�������ListIndex�����ѡ���ˣ������
            If .Selected(eachListIndex) = True Then
                mocName = .List(eachListIndex)
                '��������MOC���뵽�������ӳ����
                Call btsNameManager.addMocToSelected(mocName)
            End If
        Next eachListIndex
    End With
    
    '������ʾ
    Call initListboxAndButtons
End Sub

Private Sub AddAllCommandButton_Click()
    Dim eachListIndex As Long
    Dim mocName As String
    With Me.UnselectedBTSNameListBox
        '���û�п���ӵģ���ֱ���˳�
        If .ListCount = 0 Then Exit Sub
        
        Dim maxLength As Integer
        maxLength = 0
        
        For eachListIndex = 0 To .ListCount - 1
            '���α�������ListIndex������
            mocName = .List(eachListIndex)
            '��������MOC���뵽�������ӳ����
            Call btsNameManager.addMocToSelected(mocName)
            
            Dim length As Integer
            length = LenB(mocName)
            
            If length > maxLength Then maxLength = length
        Next eachListIndex
        
        .ColumnWidths = maxLength + 100
    End With
    
    '������ʾ
    Call initListboxAndButtons
End Sub

Private Sub DeleteCommandButton_Click()
    Dim eachListIndex As Long
    Dim mocName As String
    With Me.SelectedBTSNameListBox
        '���û�п���ӵģ���ֱ���˳�
        If .ListCount = 0 Then Exit Sub
        For eachListIndex = 0 To .ListCount - 1
            '���α�������ListIndex�����ѡ���ˣ������
            If .Selected(eachListIndex) = True Then
                mocName = .List(eachListIndex)
                '��������MOC���뵽�������ӳ����
                Call btsNameManager.addMocToUnselected(mocName)
            End If
        Next eachListIndex
    End With
    
    '������ʾ
    Call initListboxAndButtons
End Sub

Private Sub DeleteAllCommandButton_Click()
    Dim eachListIndex As Long
    Dim mocName As String
    With Me.SelectedBTSNameListBox
        '���û�п���ӵģ���ֱ���˳�
        If .ListCount = 0 Then Exit Sub
                
        Dim maxLength As Integer
        maxLength = 0
        
        For eachListIndex = 0 To .ListCount - 1
            '���α�������ListIndex������
            mocName = .List(eachListIndex)
            '��������MOC���뵽�������ӳ����
            Call btsNameManager.addMocToUnselected(mocName)
            
            mocName = .List(eachListIndex)
            
            Dim length As Integer
            length = LenB(mocName)
            
            If length > maxLength Then maxLength = length
        Next eachListIndex
        
         .ColumnWidths = maxLength + 100
    End With
    
    '������ʾ
    Call initListboxAndButtons
End Sub

Private Sub CancelCommandButton_Click()
'    Call setNextStepFlag(False)
    Unload Me
End Sub

Private Sub NextCommandButton_Click()
    Dim baseStationName As Variant
    Dim CellSheetName As String
    Call setNextStepFlag(True)
'    Call btsNameManager.setUnselectedMocFlag
    Dim selectedMocCol As Collection, unselectedMocCol As Collection 'ѡ��MOC��δѡ��MOC����
    Call btsNameManager.getCollections(selectedMocCol, unselectedMocCol)
    
    CellSheetName = ActiveSheet.name
    
    Call AddSectorEqm(selectedMocCol, CellSheetName)
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    Call initBTSNameManager
    
    '��ʼ������������
    'Call initBaseStationScenario
    
    Call initListboxAndButtons
ErrorHandler:

End Sub

Private Sub initListboxAndButtons()
    Call initListBox
    Call controlAddDeleteButtons
    
    '�����ʼ��ʱ��ѡ��ť�һ� DTS2017042603370
    Call setSingleAddButtonsFlag(False)
    Call setSingleDeleteButtonsFlag(False)
End Sub

Private Sub initListBox()
    '�õ�ѡ����δѡ����moc����
    Dim selectedMocCol As Collection, unselectedMocCol As Collection 'ѡ��MOC��δѡ��MOC����
    Call btsNameManager.getCollections(selectedMocCol, unselectedMocCol)
    
    '�������е�ֵ���뵽����ListBox��
    Call addItemOfListBox(Me.SelectedBTSNameListBox, selectedMocCol)
    Call addItemOfListBox(Me.UnselectedBTSNameListBox, unselectedMocCol)
End Sub

'������ѡ��δѡ��moc��ListBox��������ɾ����ť�Ļһ���ʾ
Private Sub controlAddDeleteButtons()
    If Me.UnselectedBTSNameListBox.ListCount = 0 Then
        Call setAddButtonsFlag(False)
    Else
        Call setAddButtonsFlag(True)
    End If
    
    If Me.SelectedBTSNameListBox.ListCount = 0 Then
        Call setDeleteButtonsFlag(False)
        Call setNextButtonsFlag(False)
    Else
        Call setDeleteButtonsFlag(True)
        Call setNextButtonsFlag(True)
    End If
End Sub

Private Sub setAddButtonsFlag(ByRef flag As Boolean)
    Me.AddCommandButton.Enabled = flag
    Me.AddAllCommandButton.Enabled = flag
End Sub

Private Sub setDeleteButtonsFlag(ByRef flag As Boolean)
    Me.DeleteCommandButton.Enabled = flag
    Me.DeleteAllCommandButton.Enabled = flag
End Sub

Private Sub setSingleAddButtonsFlag(ByRef flag As Boolean)
    Me.AddCommandButton.Enabled = flag
End Sub

Private Sub setSingleDeleteButtonsFlag(ByRef flag As Boolean)
    Me.DeleteCommandButton.Enabled = flag
End Sub

Private Sub setNextButtonsFlag(ByRef flag As Boolean)
    Me.NextCommandButton.Enabled = flag
End Sub
'��һ��ListBox����һ��Col����Col�е�������ӵ�ListBox��
Private Sub addItemOfListBox(ByRef lb As Variant, ByRef col As Collection)
    '����б�
    lb.Clear
    
    '���û������ֵΪ�գ���ֱ���˳�
    If col.count = 0 Then Exit Sub
        
    Dim maxLength As Integer
    maxLength = 0
    
    Dim eachItem As Variant
    For Each eachItem In col
        lb.AddItem eachItem
        
        Dim length As Integer
        length = LenB(eachItem)
        
        If length > maxLength Then maxLength = length
    Next eachItem

    lb.ColumnWidths = maxLength * 7 / 3
    lb.listIndex = 0
End Sub


'
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


'Private Sub initBaseStationScenario()
'    Dim sheet As Worksheet
'    Dim maxRow As Long
'    Set sheet = ThisWorkbook.Worksheets(getResByKey("BaseTransPort"))
'    maxRow = sheet.range("a1048576").End(xlUp).row
'    '��ʼ����������
'    Set scenaioCollData = New Collection
'    '��ȡ��վ����ҳǩ��վ������
'    Dim btsColNum As Long
'    btsColNum = getTransBTSNameCol(getResByKey("BaseTransPort"))
'    '��ȡ������ʽ������
'    Dim brdColNum As Long
'    brdColNum = getColNum(sheet.name, 2, "BRDSTYLE", "BTS")
'
'    '��ȡ������
'    Dim scenarioColNumArr() As String
'    Dim scenarioColNumStr As String
'    scenarioColNumStr = getScenarioColNum(sheet.name, 2, CUSTOM_SCENARIO_ATTR_NAME, CUSTOM_SCENARIO_MOC_NAME)
'    If scenarioColNumStr = "" Then Exit Sub
'    scenarioColNumArr = Split(scenarioColNumStr, ",")
'
'    Dim rowIndex As Long
'    Dim sceIndex As Long
'    Dim temSceIndex As Long
'    Dim temBtsName As String
'    Dim temBrdShtName As String
'    Dim temGroupName As String
'    Dim temColName As String
'    Dim temScenarioStr As String
'    Dim temScenarioArr() As String
'    Dim temSecnario As String
'    '������վ����ҳ����ʼ�����������ݣ����ݽṹ��ʽΪcollection(������ʽҳǩ��collection(�飬collection(�У�collection(������))))
'    For rowIndex = 3 To maxRow
'        temBtsName = sheet.Cells(rowIndex, btsColNum)
'        temBrdShtName = sheet.Cells(rowIndex, brdColNum)
'        For sceIndex = LBound(scenarioColNumArr) To UBound(scenarioColNumArr)
'            temGroupName = sheet.Cells(1, val(scenarioColNumArr(sceIndex)))
'            temColName = sheet.Cells(2, val(scenarioColNumArr(sceIndex)))
'            temScenarioStr = sheet.Cells(rowIndex, val(scenarioColNumArr(sceIndex)))
'            temScenarioArr = Split(temScenarioStr, ",")
'            For temSceIndex = LBound(temScenarioArr) To UBound(temScenarioArr)
'                temSecnario = temScenarioArr(temSceIndex)
'                '��ʼ����������
'                Call prepareScenarioData(temBtsName, temBrdShtName, temGroupName, temColName, temSecnario)
'            Next
'        Next
'    Next
'End Sub

'���ݽṹ��ʽΪcollection(��վ����#������ʽҳǩ��collection(�飬collection(�У�collection(������))))
'Private Sub prepareScenarioData(ByRef btsName As String, ByRef BrdShtName As String, ByRef groupName As String, ByRef ColName As String, ByRef Secnario As String)
'    If btsName = "" Or BrdShtName = "" Or groupName = "" Or ColName = "" Or Secnario = "" Then
'        Exit Sub
'    End If
'    '������վ���ƺ͵�����ʽҳǩ���ƣ����������������
'    Dim keyValue As String
'    keyValue = btsName + "#" + BrdShtName
'    If Contains(scenaioCollData, keyValue) Then
'        Dim groupColl As Collection
'        Set groupColl = scenaioCollData(keyValue)
'        If Contains(groupColl, groupName) Then
'            Dim columnColl As Collection
'            Set columnColl = groupColl(groupName)
'            If Contains(columnColl, ColName) Then
'                Dim secnarioColl As Collection
'                Set secnarioColl = columnColl(ColName)
'                If Not Contains(secnarioColl, Secnario) Then
'                    secnarioColl.Add Item:=Secnario
'                End If
'            Else
'                Dim temSecnarioColl As Collection
'                Set temSecnarioColl = New Collection
'
'                temSecnarioColl.Add Item:=Secnario
'                columnColl.Add Item:=temSecnarioColl, key:=ColName
'            End If
'        Else
'            Dim temSecnarioColl_1 As Collection
'            Set temSecnarioColl_1 = New Collection
'            temSecnarioColl_1.Add Item:=Secnario
'
'            Dim temColumnColl As Collection
'            Set temColumnColl = New Collection
'            temColumnColl.Add Item:=temSecnarioColl_1, key:=ColName
'
'            groupColl.Add Item:=temColumnColl, key:=groupName
'        End If
'    Else
'        Dim temSecnarioColl_2 As Collection
'        Set temSecnarioColl_2 = New Collection
'        temSecnarioColl_2.Add Item:=Secnario
'
'        Dim temColumnColl_1 As Collection
'        Set temColumnColl_1 = New Collection
'        temColumnColl_1.Add Item:=temSecnarioColl_2, key:=ColName
'
'        Dim temGroupColl As Collection
'        Set temGroupColl = New Collection
'        temGroupColl.Add Item:=temColumnColl_1, key:=groupName
'
'        scenaioCollData.Add Item:=temGroupColl, key:=keyValue
'    End If
'End Sub
'��ߴ������ѡ������ѡ����ť����
Private Sub UnselectedBTSNameListBox_Change()
    Call setSingleAddButtonsFlag(True)
End Sub

'�ұߴ������ѡ������ѡ����ť����
Private Sub SelectedBTSNameListBox_Change()
    Call setSingleDeleteButtonsFlag(True)
End Sub


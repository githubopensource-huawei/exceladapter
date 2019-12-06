VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BoardStyleForm 
   Caption         =   "Expand/Migration to Row"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7830
   OleObjectBlob   =   "BoardStyleForm.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "BoardStyleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Option Explicit

Private Const MaxMocNumber As Long = 10
Private cancleMigrateAllRec As Boolean

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub FinishButton_Click()
    Unload Me
End Sub

Private Sub MigrateAllRec_Click()
    If Me.Migrateallrec.value = True Then
        Me.SourceBaseStationList.Enabled = True
        Me.MocNumberComboBox.Enabled = False
        Me.NextButton.Caption = getResByKey("Finish")
        Call initBaseStationList
    Else
        Me.SourceBaseStationList.Enabled = False
        Me.MocNumberComboBox.Enabled = True
        Me.NextButton.Enabled = True
        Me.NextButton.Caption = getResByKey("NextButton.Caption")
        Call initMocNameComboBox
    End If
End Sub

Private Sub NextButton_Click()
    Call makeNewRecords
    Me.Hide
    If Me.Migrateallrec.value = True Then
        Application.DisplayAlerts = False
        Application.EnableEvents = False
        
        If cancleMigrateAllRec = False Then
            Call addBoardStyleMoiFinishButton
        Else
            Unload Me
        End If
        
        Application.EnableEvents = True
        Application.DisplayAlerts = True
    End If
End Sub

Private Sub makeNewRecords()
    Dim groupName As String
    Dim moiNumber As Long
    Dim sourceNeName As String
    
    Dim sourceWs As Worksheet
    Dim sourcewsName As String
    Dim sourcegroupNameStartRowNumber As Long, sourcegroupNameEndRowNumber As Long
    
    groupName = Me.MocNameComboBox.value
    Set selectedGroupMappingDefData = boardStyleMappingDefMap.GetAt(groupName)
    
    Dim groupNameStartRowNumber As Long, groupNameEndRowNumber As Long
    Call getGroupNameStartAndEndRowNumber(groupName, groupNameStartRowNumber, groupNameEndRowNumber)
    
    If Me.Migrateallrec.value = True Then
        sourcewsName = ""
        sourceNeName = Me.SourceBaseStationList.value
        If neBoardStyleMap.hasKey(sourceNeName) Then sourcewsName = neBoardStyleMap.GetAt(sourceNeName)
        Set sourceWs = ThisWorkbook.Worksheets(sourcewsName)
        
        cancleMigrateAllRec = False
        If isConfirmMigrate(sourceWs, currentSheet, sourceNeName, groupName) Then
            Call getBoardStyleGroupNameStartAndEndRowNumber(sourceWs, groupName, sourcegroupNameStartRowNumber, sourcegroupNameEndRowNumber)
            
            moiNumber = sourcegroupNameEndRowNumber - sourcegroupNameStartRowNumber - 1
            Application.DisplayAlerts = False
            Application.EnableEvents = False
            Call addAllNewLines(sourceWs, sourcegroupNameStartRowNumber + 1, groupNameStartRowNumber + 2, groupNameEndRowNumber + 1, moiNumber, selectedGroupMappingDefData.totalColumnNumber)
            Application.EnableEvents = True
            Application.DisplayAlerts = True
        Else
            cancleMigrateAllRec = True
        End If
    Else
        moiNumber = CLng(Me.MocNumberComboBox.value)

        Call addNewLines(groupNameStartRowNumber + 2, groupNameEndRowNumber + 1, moiNumber, selectedGroupMappingDefData.totalColumnNumber)
        'Call setNewRangesStyle(groupNameEndRowNumber + 1, groupNameEndRowNumber + moiNumber, )
        Call addSourceBoardNoRangesBoxList(groupName, groupNameEndRowNumber + 1, groupNameEndRowNumber + moiNumber)
    End If
    
End Sub

'ȫ����Ǩʱ���������һ�������ݰ�Ǩ����Ŀ��վ������Ҫ��ʾ�û��Ƿ������Ǩ����˵��������Ǩ�ᵼ���ظ�
Private Function isConfirmMigrate(ByRef sourceWs As Worksheet, ByRef destWs As Worksheet, ByRef sourceNeName As String, ByRef groupName As String) As Boolean
    If Not isIncludeMigratedRecord(sourceWs, destWs, sourceNeName, groupName) Then
        isConfirmMigrate = True
        Exit Function
    End If
    
    Dim errMsg As String
    errMsg = getResByKey("ExistDuplMigrateRecords")
    
    If MsgBox(errMsg, vbQuestion + vbYesNo) = vbNo Then
        isConfirmMigrate = False
    Else
        isConfirmMigrate = True
    End If
End Function

Private Function isIncludeMigratedRecord(ByRef sourceWs As Worksheet, ByRef destWs As Worksheet, ByRef sourceNeName As String, ByRef groupName As String) As Boolean
    On Error GoTo ErrorHandler
    
    isIncludeMigratedRecord = False
    
    If Not boardStyleMappingDefMap.hasKey(groupName) Then Exit Function ' �쳣����
    Set selectedGroupMappingDefData = boardStyleMappingDefMap.GetAt(groupName)
        
    ' �ҵ�Դվ�����п��԰�Ǩ�����ݣ����keyValue��_ƴ��
    Dim srcKeyValueCol As Collection
    Set srcKeyValueCol = New Collection
    Call collectSourceKeyValueCol(sourceWs, groupName, srcKeyValueCol)
    
    ' ����Ŀ��վ�ϵİ�Ǩ���ݣ������Դվ�ϵİ�Ǩ���ݣ��򷵻�isIncludeMigratedRecord=True
    isIncludeMigratedRecord = isIncludeMigratedRecord_i(destWs, groupName, sourceNeName, srcKeyValueCol)
    
ErrorHandler:
End Function

Private Sub collectSourceKeyValueCol(ByRef sourceWs As Worksheet, ByRef groupName As String, ByRef keyValueCollection As Collection)
    
    ' ȡ�洢����������Columnλ�õ�Map����
    Dim columnNamePositionLetterMap As CMap
    Set columnNamePositionLetterMap = selectedGroupMappingDefData.columnNamePositionLetterMap
    
    ' ȡ��ǨĿ�����������б�
    Dim destColumnNameCol As Collection
    Set destColumnNameCol = selectedGroupMappingDefData.copyFillInSourceColumnName ' Ŀ����������
    
    ' ȡ���ݿ����ʼ��������
    Dim groupNameStartRowNumber As Long, groupNameEndRowNumber As Long
    Call getBoardStyleGroupNameStartAndEndRowNumber(sourceWs, groupName, groupNameStartRowNumber, groupNameEndRowNumber)
    
    ' �������н��б�����ȡÿ�����ݵ�keyStr����������ֵ��_����
    Dim keyStr As String
    Dim rowNumber As Long
    For rowNumber = groupNameStartRowNumber + 2 To groupNameEndRowNumber
        keyStr = extractValueCol(sourceWs, destColumnNameCol, columnNamePositionLetterMap, rowNumber)
        keyValueCollection.Add Item:=keyStr, key:=keyStr
    Next rowNumber
End Sub



' ����Ŀ��վ�ϵİ�Ǩ���ݣ������Դվ�ϵİ�Ǩ���ݣ��򷵻�isIncludeMigratedRecord=True
Private Function isIncludeMigratedRecord_i(ByRef ws As Worksheet, ByRef groupName As String, ByRef sourceNeName As String, ByRef srcKeyValueCol As Collection)
    isIncludeMigratedRecord_i = False
       
    ' ȡ�洢����������Columnλ�õ�Map����
    Dim columnNamePositionLetterMap As CMap
    Set columnNamePositionLetterMap = selectedGroupMappingDefData.columnNamePositionLetterMap
    
    ' ȡ��ǨԴ���������б�
    Dim sourceColumnNameCol As Collection
    Set sourceColumnNameCol = selectedGroupMappingDefData.autoFillInSourceColumnName ' Դ��������
    
    ' ȡ���ݿ����ʼ��������
    Dim groupNameStartRowNumber As Long, groupNameEndRowNumber As Long
    Call getBoardStyleGroupNameStartAndEndRowNumber(ws, groupName, groupNameStartRowNumber, groupNameEndRowNumber)
    
    ' ȡԴվ������
    Dim sourceNeNameLetter As String '���߰�Ǩ����У���Ϊ��Ч���ݴ���һ�п�ʼ
    sourceNeNameLetter = getSrcNeNameColWithGrpName(ws, groupNameStartRowNumber, groupName)
    If sourceNeNameLetter = "" Then Exit Function
    
    ' �������н��б�����ȡÿ�����ݵ�keyStr����������ֵ��_����
    Dim keyStr As String
    Dim rowNumber As Long
    Dim curSourceNeName As String
    For rowNumber = groupNameStartRowNumber + 2 To groupNameEndRowNumber
        curSourceNeName = ws.range(sourceNeNameLetter & rowNumber).value
        
        If curSourceNeName = sourceNeName Then
            keyStr = extractValueCol(ws, sourceColumnNameCol, columnNamePositionLetterMap, rowNumber)
            If isInCollection(srcKeyValueCol, keyStr) Then
                isIncludeMigratedRecord_i = True
                Exit Function
            End If
        End If
    Next rowNumber

End Function

' ����columnNameCol��ȡָ��sheet��ָ���е�ValueColStr
Private Function extractValueCol(ByRef ws As Worksheet, ByRef columnNameCol, ByRef columnNamePositionLetterMap As CMap, ByVal rowNumber As Long) As String
    Dim columnName As Variant
    Dim keyStr As String
    Dim tempValue As String
    Dim columnNamePositionLetter As String
    
    ' ÿ���ڰ�Ŀ���������Ʊ���
    For Each columnName In columnNameCol
        
        ' ȡ������λ��,���Բ����ڵ��쳣������������
        columnNamePositionLetter = columnNamePositionLetterMap.GetAt(columnName)
        
        ' �ϲ�����
        tempValue = ws.range(columnNamePositionLetter & rowNumber).value
        If keyStr = "" Then
            keyStr = tempValue
        Else
            keyStr = keyStr + "," + tempValue
        End If
    Next

    extractValueCol = keyStr
End Function
 

Private Sub addAllNewLines(ByRef ws As Worksheet, ByVal sourceStartRowNumber As Long, ByVal srcRowNumber As Long, ByVal dstRowNumber As Long, ByVal moiNumber As Long, ByVal maxColumnNumber As Long)
    Call insertNewLines(srcRowNumber, dstRowNumber, moiNumber)
    Call copyAllNewLines(ws, sourceStartRowNumber, dstRowNumber, moiNumber)
    Call setNewRangesStyle(srcRowNumber, dstRowNumber, moiNumber, maxColumnNumber, True)
End Sub

Private Sub copyAllNewLines(ByRef ws As Worksheet, ByVal srcRowNumber As Long, ByVal dstRowNumber As Long, ByVal moiNumber As Long)
    Dim k As Long
    Dim srcRowRange As range, dstRowRange As range
    Application.EnableEvents = False
    For k = 1 To moiNumber
        Set srcRowRange = ws.Rows(srcRowNumber + k)
        Set dstRowRange = currentSheet.Rows(dstRowNumber + k - 1)
        srcRowRange.Copy
        dstRowRange.PasteSpecial
    Next k
    Application.EnableEvents = True
    Application.CutCopyMode = False
End Sub

Private Sub addNewLines(ByVal srcRowNumber As Long, ByVal dstRowNumber As Long, ByVal moiNumber As Long, ByVal maxColumnNumber As Long)
    Call insertNewLines(srcRowNumber, dstRowNumber, moiNumber)
    Call setNewRangesStyle(srcRowNumber, dstRowNumber, moiNumber, maxColumnNumber, False)
End Sub

Private Sub insertNewLines(ByVal srcRowNumber As Long, ByVal dstRowNumber As Long, ByVal moiNumber As Long)
    Dim k As Long
    Dim srcRowRange As range, dstRowRange As range
    Set srcRowRange = currentSheet.Rows(srcRowNumber)
    Set dstRowRange = currentSheet.Rows(dstRowNumber)
        Application.EnableEvents = False
    For k = 1 To moiNumber
        srcRowRange.Copy
        dstRowRange.Insert Shift:=xlShiftDown
    Next k
        Application.EnableEvents = True
    Application.CutCopyMode = False
End Sub

Private Sub setNewRangesStyle(ByVal srcRowNumber As Long, ByVal dstRowNumber As Long, ByVal moiNumber As Long, ByVal maxColumnNumber As Long, migrateFlag As Boolean)
    Call initMoiRowsManager(srcRowNumber, dstRowNumber, moiNumber, maxColumnNumber)
    Call setNewRangesBackgroundColour(NewMoiRangeColorIndex)
    
    Dim endRowNumber As Long
    endRowNumber = dstRowNumber + moiNumber - 1
    Call setNeedFillInRangesStyles(dstRowNumber, endRowNumber)
    If migrateFlag = False Then Call clearAllBoardNoRanges(dstRowNumber, endRowNumber, maxColumnNumber)
    
    
    Call selectCertainCell(currentSheet, "A" & srcRowNumber - 2)
End Sub

Private Sub initMoiRowsManager(ByVal srcRowNumber As Long, ByVal dstRowNumber As Long, ByVal moiNumber As Long, ByVal maxColumnNumber As Long)
    moiRowsManager.groupNameRowNumber = srcRowNumber - 2
    moiRowsManager.startRowNumber = dstRowNumber
    moiRowsManager.endRowNumber = dstRowNumber + moiNumber - 1
    moiRowsManager.maxColumnNumber = maxColumnNumber
End Sub

Private Sub setNeedFillInRangesStyles(ByVal startRowNumber As Long, ByVal endRowNumber As Long)
    Dim needFillColumnNameCol As Collection
    Set needFillColumnNameCol = selectedGroupMappingDefData.needFillColumnNameCol
    Dim columnNamePositionLetterMap As CMap
    Set columnNamePositionLetterMap = selectedGroupMappingDefData.columnNamePositionLetterMap
    Dim columnName As Variant
    Dim columnNamePositionLetter As String
    Dim needFillInRange As range
    For Each columnName In needFillColumnNameCol
                If Not columnNamePositionLetterMap.hasKey(columnName) Then GoTo NextLoop
        columnNamePositionLetter = columnNamePositionLetterMap.GetAt(columnName)
        Set needFillInRange = currentSheet.range(currentSheet.range(columnNamePositionLetter & startRowNumber), currentSheet.range(columnNamePositionLetter & endRowNumber))
        needFillInRange.Interior.colorIndex = NeedFillInRangeColorIndex
        needFillInRange.ClearContents
        Call moiRowsManager.addNeedFillInRange(CStr(columnName), needFillInRange)
NextLoop:
    Next columnName
End Sub

Private Sub initMocNameComboBox()
    Set boardStyleMappingDefMap = boardStyleData.getBoardStyleMappingDefMap
    
    Dim groupNameVar As Variant
    Dim groupName As String
    Dim groupNameCollection As Collection
    Set groupNameCollection = getKeyValueCollection(boardStyleMappingDefMap.KeyCollection)
    
    If currentSheet Is Nothing Then Set currentSheet = ThisWorkbook.ActiveSheet

    Me.MocNameComboBox.Clear
    With Me.MocNameComboBox
        For Each groupNameVar In groupNameCollection
            .AddItem groupNameVar
        Next
    End With
    
    Dim selectedGroupName As String
    Call getSelectedGroupName(selectedGroupName)
    
    If Me.MocNameComboBox.ListCount <> 0 Then
        If Contains(groupNameCollection, selectedGroupName) Then 'Ĭ����ʾѡ���ķ���
            Me.MocNameComboBox.value = selectedGroupName
        Else
            Me.MocNameComboBox.listIndex = 0
        End If
    End If
End Sub
Private Sub initBaseStationList()
    Dim currentneName As String
    Dim boardStyleSheetName As String
    
    Dim migrationData As CMigrationDataManager
    Dim targetSourceNeMap As CMapValueObject
    Set migrationData = New CMigrationDataManager
    Call migrationData.init
    Set targetSourceNeMap = migrationData.targetSourceNeMap
    
    Me.SourceBaseStationList.Enabled = True
    Me.NextButton.Enabled = True
    SourceBaseStationList.Clear

    If currentSheet Is Nothing Then Set currentSheet = ThisWorkbook.ActiveSheet
    boardStyleSheetName = currentSheet.name
    If boardStyleNeMap.hasKey(boardStyleSheetName) Then currentneName = boardStyleNeMap.GetAt(boardStyleSheetName)
    
    Dim srcneName As Variant
    Dim onerowRec As CMap
    Dim srcNeMap As CMap
    Set srcNeMap = New CMap
    If targetSourceNeMap.hasKey(currentneName) Then
        Set onerowRec = targetSourceNeMap.GetAt(currentneName)
        Set srcNeMap = getSrcNeMap(onerowRec)
        For Each srcneName In srcNeMap.KeyCollection
            SourceBaseStationList.AddItem (srcneName)
        Next
    End If
    
    If SourceBaseStationList.ListCount <> 0 Then
        Me.SourceBaseStationList.listIndex = 0
    Else
        Me.SourceBaseStationList.Enabled = False
        Me.NextButton.Enabled = False
    End If
End Sub
Private Function getSrcNeMap(ByRef onerowRec As CMap) As CMap
    Dim keyValue As Variant
    Dim valueStr As String
    Dim valueStrArry() As String
    Dim index As Long
    Dim tempMap As CMap
    Set tempMap = New CMap
    For Each keyValue In onerowRec.KeyCollection
        valueStr = onerowRec.GetAt(keyValue)
        If valueStr <> "" Then
            valueStrArry = Split(valueStr, ",")
            For index = LBound(valueStrArry) To UBound(valueStrArry)
                If Not tempMap.hasKey(valueStrArry(index)) Then Call tempMap.SetAt(valueStrArry(index), valueStrArry(index))
            Next
        End If
    Next
    Set getSrcNeMap = tempMap
End Function
Private Sub getSelectedGroupName(ByRef selectedGroupName As String)
    Dim selectionRange As range
    Set selectionRange = Selection
    selectedGroupName = selectionRange(1).value
End Sub

Private Sub initMocNumberComboBox()
    Dim number As Long
    With Me.MocNumberComboBox
        For number = 1 To MaxMocNumber
            .AddItem number
        Next number
    End With
    If Me.MocNumberComboBox.ListCount <> 0 Then Me.MocNumberComboBox.listIndex = 0
End Sub


Private Sub SourceBaseStationList_Change()
    Dim sourceNeName As String
    Dim groupName As Variant
    Dim sourceBoardstyleSheetName As String
    Dim sourceBoardstylesheet As Worksheet
    Dim rowNumber As Long
    
    sourceNeName = Me.SourceBaseStationList.value
    
    If sourceNeName = "" Then Exit Sub
    
    Set boardStyleMappingDefMap = boardStyleData.getBoardStyleMappingDefMap
    Dim groupNameCollection As Collection
    Dim boardStyleMappingDefData As CBoardStyleMappingDefData
    Set groupNameCollection = getKeyValueCollection(boardStyleMappingDefMap.KeyCollection)

    If neBoardStyleMap.hasKey(sourceNeName) Then sourceBoardstyleSheetName = neBoardStyleMap.GetAt(sourceNeName)
    
    If containsASheet(ThisWorkbook, sourceBoardstyleSheetName) Then
        Set sourceBoardstylesheet = ThisWorkbook.Worksheets(sourceBoardstyleSheetName)
        Me.MocNameComboBox.Clear
        With Me.MocNameComboBox
            For Each groupName In groupNameCollection
                Set boardStyleMappingDefData = boardStyleMappingDefMap.GetAt(groupName)
                If isSourceMocRecEmpty(sourceBoardstylesheet, CStr(groupName)) = False And isneedMigrateSourceMocRec(boardStyleMappingDefData) Then .AddItem groupName
            Next
        End With
    End If
    
    If Me.MocNameComboBox.ListCount <> 0 Then Me.MocNameComboBox.listIndex = 0
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    Set addBoardStyleMoiInProcess = New CAddingBoardStyleMoc
    Set moiRowsManager = New CMoiRowsManager
    Call addBoardStyleMoiInProcess.init
    If boardStyleData Is Nothing Then
        Call initBoardStyleMappingDataPublic
    End If
    
    Call initFormBox
    
    Call initBaseStationDataPublic
    Call initAllBoardStyleDataPublic
    
    Call initMocNameComboBox
    Call initMocNumberComboBox

    Exit Sub
ErrorHandler:
End Sub

Private Sub UserForm_Terminate()
    Set addBoardStyleMoiInProcess = Nothing
    Set moiRowsManager = Nothing
End Sub

Private Sub initFormBox()
    Me.SourceBaseStationList.Enabled = False
    Me.MocNumberComboBox.Enabled = True
End Sub

Private Function isSourceMocRecEmpty(ByRef ws As Worksheet, ByRef groupName As String) As Boolean
    Dim groupNameStartRowNumber As Long, groupNameEndRowNumber As Long, sourceMocRecNum As Long
    isSourceMocRecEmpty = False
    Call getBoardStyleGroupNameStartAndEndRowNumber(ws, groupName, groupNameStartRowNumber, groupNameEndRowNumber)
    sourceMocRecNum = groupNameEndRowNumber - groupNameStartRowNumber + 1
    If sourceMocRecNum = 2 Then isSourceMocRecEmpty = True
End Function

Private Function isneedMigrateSourceMocRec(ByRef boardStyleMappingDefData As CBoardStyleMappingDefData) As Boolean
    Dim sourceAttrColl As Collection
    Set sourceAttrColl = boardStyleMappingDefData.autoFillInSourceColumnName
    isneedMigrateSourceMocRec = False
    If sourceAttrColl.count <> 0 Then isneedMigrateSourceMocRec = True
End Function


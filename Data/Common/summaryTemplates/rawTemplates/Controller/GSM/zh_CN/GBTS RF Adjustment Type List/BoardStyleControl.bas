Attribute VB_Name = "BoardStyleControl"

Option Explicit
Private groupAndRowNumMap As CMap

'ҳǩ�л�ʱ���ã������к͵�Ԫ�͵Ĳ���ɾ������
Public Sub insertAndDeleteControl(ByRef flag As Boolean)
    On Error Resume Next
    
    With Application
        With .CommandBars("Column")
            .FindControl(ID:=3183).Enabled = flag '����
            .FindControl(ID:=297).Enabled = flag '��
            .FindControl(ID:=294).Enabled = flag 'ɾ��
        End With
        With .CommandBars("Cell")
            .FindControl(ID:=3181).Enabled = flag '����
            .FindControl(ID:=295).Enabled = flag '��
            .FindControl(ID:=292).Enabled = flag 'ɾ��
        End With
    End With
End Sub

'��Ԫ���л�ʱ���ã������顢�к���ǰ���еĲ����ɾ��
Public Sub rowControl(ByVal sh As Object, ByVal Target As Range)
    Set groupAndRowNumMap = New CMap
    '����Group���������У���ʼ��BoardStyleҳǩ����
    Call initBoardStyleGroupRowNumMap
    '�����顢�к���ǰһ�����У�����ִ�в��롢ɾ������
    Call setGroupAndColumnInsertContorl(sh, Target)
    
End Sub

Private Sub setRowControl(ByRef flag As Boolean)
    On Error Resume Next
    With Application
        '������ɾ��������
        With .CommandBars("Row")
            .FindControl(ID:=3183).Enabled = flag '����
            .FindControl(ID:=296).Enabled = flag '��
            .FindControl(ID:=293).Enabled = flag 'ɾ��
        End With
    End With
End Sub

Private Sub setRowInsertControl(ByRef flag As Boolean)
    On Error Resume Next
    With Application
        '������ɾ��������
        With .CommandBars("Row")
            .FindControl(ID:=3183).Enabled = flag '����
            .FindControl(ID:=296).Enabled = flag '��
            .FindControl(ID:=293).Enabled = True 'ɾ��
        End With
    End With
End Sub

Private Sub setGroupAndColumnInsertContorl(ByVal sh As Object, ByVal Target As Range)
    
    If groupAndRowNumMap.isEmpty Then
        Exit Sub
    End If
    
    Dim index As Long
    Dim rowFlage As Long
    '��Ҫ�ж���ѡ����������
    For index = Target.row To Target.row + Target.Rows.count - 1
        rowFlage = isGroupMapHasValue(index)
        '����������ǰһ�У�������һ�У����Ʋ����ɾ������
        If rowFlage = 1 Then
            Call setRowControl(False)
            Exit For
        '������һ�У����Ʋ��룬����ɾ����ֻ��ѡ�е���ʱ����
        ElseIf rowFlage = 2 And Target.Rows.count = 1 Then
            Call setRowInsertControl(False)
            Exit For
        Else
            Call setRowControl(True)
        End If
    Next
End Sub

Private Function isGroupMapHasValue(ByRef value As Long) As Long
    Dim rowNum As Variant
    isGroupMapHasValue = -1
    For Each rowNum In groupAndRowNumMap.ValueCollection
        If value = rowNum Or value = rowNum + 1 Or value = rowNum - 1 Then
            isGroupMapHasValue = 1
            Exit For
        ElseIf value = rowNum + 2 Then
            isGroupMapHasValue = 2
            Exit For
        Else
            isGroupMapHasValue = -1
        End If
    Next
End Function

Private Sub initBoardStyleGroupRowNumMap()
    Dim mappingDefSheet As Worksheet
    Set mappingDefSheet = ThisWorkbook.Worksheets("MAPPING DEF")
    
    Dim sheetName As String
    Dim groupName As String
    Dim columnName As String
    Dim mocName As String
    Dim attributeName As String
    
    Dim RowNumber As Long
    
    '����MappingDefҳǩ
    For RowNumber = 2 To mappingDefSheet.Range("a1048576").End(xlUp).row
        '��ȡMappingDef��Ϣ
        Call clearMappingData(sheetName, groupName, columnName, mocName, attributeName)
        Call getMappingData(sheetName, groupName, columnName, mocName, attributeName, mappingDefSheet, RowNumber)
        '��ʼ��BoardStyleҳǩ���ݣ��ų��������ͳ�����
        If sheetName = getResByKey("Board Style") And InStr(groupName, getResByKey("Operation")) = 0 And InStr(mocName, "Customization") = 0 Then
            Call getGroupAndRowNumMap(groupName)
        End If
    Next
End Sub

Private Sub getGroupAndRowNumMap(ByRef groupName As String)
    Dim boardStyleSheet As Worksheet
    Set boardStyleSheet = ThisWorkbook.ActiveSheet
    Dim groupNameRowNumber As Long
    
    
    '��ȡ���������к�
    groupNameRowNumber = getGroupNameStartRowNumber(boardStyleSheet, CStr(groupName))
    'δ��ȡ�����������кţ��˳�
    If groupNameRowNumber = -1 Then
        Exit Sub
    End If
    
    Dim mappingKey As String
    '����"sheetName#groupName"��Key���к���Value
    mappingKey = boardStyleSheet.name + "#" + groupName
    
    If Not groupAndRowNumMap.hasKey(mappingKey) Then
        groupAndRowNumMap.SetAt mappingKey, groupNameRowNumber
    End If
End Sub


Private Sub clearMappingData(ByRef sheetName As String, ByRef groupName As String, ByRef columnName As String, ByRef mocName As String, _
    ByRef attributeName As String)
    sheetName = ""
    groupName = ""
    columnName = ""
    mocName = ""
    attributeName = ""
End Sub

Private Sub getMappingData(ByRef sheetName As String, ByRef groupName As String, ByRef columnName As String, ByRef mocName As String, _
    ByRef attributeName As String, ByRef mappingDefSheet As Worksheet, ByRef RowNumber As Long)
    sheetName = mappingDefSheet.Range("A" & RowNumber).value
    groupName = mappingDefSheet.Range("B" & RowNumber).value
    columnName = mappingDefSheet.Range("C" & RowNumber).value
    mocName = mappingDefSheet.Range("D" & RowNumber).value
    attributeName = mappingDefSheet.Range("E" & RowNumber).value
End Sub

Private Function getBoardStyleSheet() As Worksheet
    Dim boardStyleSheetName As String
    Dim sheet As Worksheet
    Dim sheetName As String
    boardStyleSheetName = getResByKey("Board Style")
    If containsASheet(ThisWorkbook, boardStyleSheetName) Then
        Set getBoardStyleSheet = ThisWorkbook.Worksheets(boardStyleSheetName)
    Else
        For Each sheet In ThisWorkbook.Worksheets
            sheetName = sheet.name
            If InStr(sheetName, boardStyleSheetName) <> 0 Then
                Set getBoardStyleSheet = ThisWorkbook.Worksheets(sheetName)
                Exit Function
            End If
        Next sheet
    End If
End Function



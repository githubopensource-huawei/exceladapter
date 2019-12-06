Attribute VB_Name = "neighborPolicySub"
Option Explicit
Public bandDefStr As String


Public Sub initBandDefData()
    Dim colNum As Long
    Dim rowIndex As Long
    Dim sheet As Worksheet
    
    If IsExistsSheet(getResByKey("BAND_DEFINITION")) Then
        Set sheet = ThisWorkbook.Sheets(getResByKey("BAND_DEFINITION"))
    Else
        Exit Sub
    End If
    
    bandDefStr = ""
    
    '��ǰҳǩ�ҵ�*BandInd������
    colNum = getColNum(sheet.name, 2, "BANDIND", "BANDDEF")
    
    For rowIndex = 3 To ThisWorkbook.Sheets(sheet.name).range(getColumnNameFromColumnNum(colNum) + "1048576").End(xlUp).row
        If ThisWorkbook.Sheets(sheet.name).Cells(rowIndex, colNum) <> "" Then
            If bandDefStr = "" Then
                bandDefStr = ThisWorkbook.Sheets(sheet.name).Cells(rowIndex, colNum).value
            Else
                bandDefStr = bandDefStr + "," + ThisWorkbook.Sheets(sheet.name).Cells(rowIndex, colNum).value
            End If
        End If
    Next
    'ȥ��
    bandDefStr = abandonRepeatedString(bandDefStr)
    
End Sub

Public Sub neighborPolicyShtSelectionChange(ByVal sheet As Worksheet, ByVal target As range)
    Dim maxRowNum As Long
    '*BandInd�����ݣ��������ݳ�ʼ��ʧ�ܣ��˳�
    If bandDefStr = "" Then Exit Sub
    
    maxRowNum = ThisWorkbook.ActiveSheet.range("a1048576").End(xlUp).row
    If maxRowNum < 3 Then Exit Sub
    
    '����*Source BandInd�У����������б�
    Call setSourceBandIndValidation
    '����Target BandInd�У����������б�
    Call setTargetBandIndValidation
    '����Neighbor Reference Source BandInd�У����������б�
    Call setNeighborSourceValidation
    '����Neighbor Reference Target BandInd�У����������б�
    Call setNeighborTargetValidation
    
End Sub

Private Sub setSourceBandIndValidation()
    Dim colNum As Long
    Dim maxRowNum As Long
    
    colNum = getColNum(ThisWorkbook.ActiveSheet.name, 2, "SRCBANDIND", "NBPOLICY")
    'maxRowNum = ThisWorkbook.ActiveSheet.Range(getColumnNameFromColumnNum(colNum) + "1048576").End(xlUp).row
    maxRowNum = ThisWorkbook.ActiveSheet.range("a1048576").End(xlUp).row
    
    Dim rng As range
    Set rng = ThisWorkbook.ActiveSheet.range(getColumnNameFromColumnNum(colNum) + "3:" + getColumnNameFromColumnNum(colNum) + CStr(maxRowNum))
    
    With rng.Validation
        .delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=bandDefStr
    End With
    
    
End Sub

Private Sub setTargetBandIndValidation()
    Dim colNum As Long
    Dim maxRowNum As Long
    
    colNum = getColNum(ThisWorkbook.ActiveSheet.name, 2, "TARGETBANDIND", "NBPOLICY")
    'maxRowNum = ThisWorkbook.ActiveSheet.Range(getColumnNameFromColumnNum(colNum) + "1048576").End(xlUp).row
    maxRowNum = ThisWorkbook.ActiveSheet.range("a1048576").End(xlUp).row
    Dim rng As range
    Set rng = ThisWorkbook.ActiveSheet.range(getColumnNameFromColumnNum(colNum) + "3:" + getColumnNameFromColumnNum(colNum) + CStr(maxRowNum))
    
    With rng.Validation
        .delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=bandDefStr
    End With
    
    
End Sub

Private Sub setNeighborSourceValidation()
    Dim colNum As Long
    Dim maxRowNum As Long
    
    colNum = getColNum(ThisWorkbook.ActiveSheet.name, 2, "NBREFSRCBANDIND", "NBPOLICY")
    'maxRowNum = ThisWorkbook.ActiveSheet.Range(getColumnNameFromColumnNum(colNum) + "1048576").End(xlUp).row
    maxRowNum = ThisWorkbook.ActiveSheet.range("a1048576").End(xlUp).row
    Dim rng As range
    Set rng = ThisWorkbook.ActiveSheet.range(getColumnNameFromColumnNum(colNum) + "3:" + getColumnNameFromColumnNum(colNum) + CStr(maxRowNum))
    
    With rng.Validation
        .delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=bandDefStr
    End With
    
    
End Sub

Private Sub setNeighborTargetValidation()
    Dim colNum As Long
    Dim maxRowNum As Long
    
    colNum = getColNum(ThisWorkbook.ActiveSheet.name, 2, "NBREFTARGETBANDIND", "NBPOLICY")
    'maxRowNum = ThisWorkbook.ActiveSheet.Range(getColumnNameFromColumnNum(colNum) + "1048576").End(xlUp).row
    maxRowNum = ThisWorkbook.ActiveSheet.range("a1048576").End(xlUp).row
    Dim rng As range
    Set rng = ThisWorkbook.ActiveSheet.range(getColumnNameFromColumnNum(colNum) + "3:" + getColumnNameFromColumnNum(colNum) + CStr(maxRowNum))
    
    With rng.Validation
        .delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=bandDefStr
    End With
    
    
End Sub

'**********************************************************
'???????:1->A,27->AA
'**********************************************************
Private Function getColumnNameFromColumnNum(iColumn As Long) As String
  If iColumn >= 257 Or iColumn < 0 Then
    getColumnNameFromColumnNum = ""
    Return
  End If
  
  Dim result As String
  Dim High, Low As Long
  
  High = Int((iColumn - 1) / 26)
  Low = iColumn Mod 26
  
  If High > 0 Then
    result = Chr(High + 64)
  End If

  If Low = 0 Then
    Low = 26
  End If
  
  result = result & Chr(Low + 64)
  getColumnNameFromColumnNum = result
End Function

'������ǰ�����жϵ�ǰ�������Ƿ���ڸ�ҳǩ
Private Function IsExistsSheet(sheetName As String) As Boolean
  Dim ShtIdx As Long
  Dim OpSht As Worksheet
  
  ShtIdx = 1
  Do While (ShtIdx <= ActiveWorkbook.Sheets.count)
      Set OpSht = ActiveWorkbook.Sheets(ShtIdx)
      If OpSht.name = sheetName Then
        IsExistsSheet = True
        Exit Function
      End If
      ShtIdx = ShtIdx + 1
  Loop
  IsExistsSheet = False
End Function


'�����б��ַ���ȥ�أ���","�ָ����ַ�����
Public Function abandonRepeatedString(referencedString As String)

    Dim valueArr
    Dim valueColl As New Collection
    Dim value As Variant
    Dim i As Integer
    Dim j As Integer
    
    valueArr = Split(referencedString, ",")
        
    For Each value In valueArr
        valueColl.Add value
    Next
        
    abandonRepeatedString = ""
        
    If valueColl.count >= 2 Then
        For i = 1 To valueColl.count
            For j = i + 1 To valueColl.count
                If valueColl.Item(j) = valueColl.Item(i) Then
                    valueColl.Remove (j)
                    j = j - 1
                End If
                If j = valueColl.count Then
                    Exit For
                End If
            Next
        Next
    End If
    For i = 1 To valueColl.count
        If i = 1 Then
            abandonRepeatedString = valueColl.Item(i)
        Else
            abandonRepeatedString = abandonRepeatedString + "," + valueColl.Item(i)
        End If
    Next

End Function

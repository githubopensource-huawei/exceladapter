VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************
'类模块名称：TableClass
'类模块作用：方便在内存中构建一张表
'***************************************************************
Option Explicit

Private Records As New Collection

Public Sub AddField(Attr As AttrClass)
    If Records.Count > 1 Then 'If records is not empty, can not add field.
        Exit Sub
    End If

    Dim Record As Collection
    If Records.Count <= 0 Then
        Set Record = New Collection
        Records.Add Item:=Record
    Else
        Set Record = Records.Item(1)
    End If
    
    Record.Add Item:=Attr, Key:=Attr.Name
End Sub

Public Sub AddRecord()
    If Records.Count <= 0 Then 'If no field defination, can not add record.
        Exit Sub
    End If

    Dim Record1 As Collection, Record2 As Collection
    Set Record1 = Records.Item(1)
    Set Record2 = New Collection
    Records.Add Item:=Record2
    
    Dim i As Integer
    Dim Attr As AttrClass
    For i = 1 To Record1.Count
        Set Attr = Record1.Item(i)
        Set Attr = Attr.Clone
        Attr.Value = ""
        Record2.Add Item:=Attr, Key:=Attr.Name
    Next i
End Sub

Public Function GetValueAt(RecordIndex As Long, FieldIndex As Variant) As AttrClass
    Set GetValueAt = Nothing

    Dim iFieldIndex As Long
    If Not IsNumeric(FieldIndex) Then
        iFieldIndex = GetFieldIndex(CStr(FieldIndex))
    Else
        iFieldIndex = FieldIndex
    End If

    Dim Record As Collection
    If (1 <= RecordIndex) And (RecordIndex <= Records.Count - 1) Then
        Set Record = Records.Item(RecordIndex + 1)
    Else
        Exit Function
    End If
    
    If (1 <= iFieldIndex) And (iFieldIndex <= Record.Count) Then
        Set GetValueAt = Record.Item(iFieldIndex)
    Else
        Exit Function
    End If
End Function

Private Function GetFieldIndex(AttrName As String) As Long
    GetFieldIndex = -1
    If Records.Count <= 0 Then
        Exit Function
    End If
    Dim Record As Collection
    Set Record = Records.Item(1)
    Dim i As Integer, Attr As AttrClass
    For i = 1 To Record.Count
        Set Attr = Record.Item(i)
        If Attr.Name = AttrName Then
            GetFieldIndex = i
            Exit For
        End If
    Next i
End Function

Public Function GetRecordCount() As Long
    GetRecordCount = Records.Count - 1
End Function

Public Function GetFieldCount() As Long
    GetFieldCount = -1
    If Records.Count <= 0 Then
        Exit Function
    End If
    Dim Record As Collection
    Set Record = Records.Item(1)
    GetFieldCount = Record.Count
End Function

Public Sub Clear(Optional IsKeepField As Boolean = True)
    Dim Record As Collection, Attr As AttrClass
    Dim iRecordIndex As Long
    iRecordIndex = IIf(IsKeepField, 1, 0)
    Do While Records.Count > iRecordIndex
        Set Record = Records.Item(2)
        Records.Remove (2)
        Do While Record.Count > 0
            Set Attr = Record.Item(1)
            Record.Remove (1)
            Set Attr = Nothing
        Loop
        Set Record = Nothing
    Loop
End Sub

Public Sub Remove(RecordIndex As Long)
    Dim Record As Collection
    If (1 <= RecordIndex) And (RecordIndex <= Records.Count - 1) Then
        Set Record = Records.Item(RecordIndex + 1)
    Else
        Exit Sub
    End If

    Dim Attr As AttrClass
    Do While Record.Count > 0
        Set Attr = Record.Item(1)
        Record.Remove (1)
        Set Attr = Nothing
    Loop
    Records.Remove (RecordIndex + 1)
    Set Record = Nothing
End Sub

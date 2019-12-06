VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CReferenceMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private referenceMessMap_ As CMap '引用提示信息
Private Const ReferenceRow_ As Long = 3
Private Const delimiter As String = ","

Private Sub init()
    Dim address As Collection
    Dim iRange As range
    Set address = New Collection
    Dim ws As Worksheet
    Dim sheetnam As String
    sheetnam = GetMainSheetName()
    Dim colNum As Integer
    Dim mainsheet As Worksheet
    Set mainsheet = ThisWorkbook.Sheets(sheetnam)
    
    Dim addressValues As Variant, ele As Variant
    
    
    For colNum = 1 To mainsheet.range("IV" + CStr(ReferenceRow_)).End(xlToLeft).column
        If VBA.Trim(mainsheet.Cells(ReferenceRow_, colNum).value) <> "" Then
           Dim tmpvalue As String
           tmpvalue = mainsheet.Cells(ReferenceRow_, colNum).value
           If InStr(tmpvalue, delimiter) <> 0 Then
            addressValues = Split(tmpvalue, delimiter)
            For Each ele In addressValues
                address.Add Item:=CStr(ele)
            Next ele
           Else
             address.Add Item:=tmpvalue
           End If
           
        End If
    Next
    
    For Each ws In ThisWorkbook.Worksheets
        If isIubStyleWorkSheetByParameterWs(ws) Then
                Dim i As Integer
                Dim element As Variant
                For Each element In address
                  Set iRange = ws.range(CStr(element))
                  Dim key As String
                  Dim value As String
                  key = CStr(element)
                  value = iRange.Validation.inputmessage
                 Call referenceMessMap_.SetAt(key, value)
                Next element
                 Exit For
        End If
       
    Next
    
End Sub

Public Sub generateCurrentReferenceStr()
    Call init '每次加载一次
End Sub

Public Sub addReferenceMess(ByRef address As String, ByRef inputmessage As String)
   Call referenceMessMap_.SetAt(address, inputmessage)
End Sub

Public Function isReferenced(ByRef address As String) As Boolean
    isReferenced = referenceMessMap_.hasKey(address)
End Function

Public Sub remmoveReferenceMess(ByRef address As String)
   Call referenceMessMap_.RemoveKey(address)
End Sub

Public Function getReferencMess(ByRef address As String) As String
   On Error GoTo ErrorHandle
  Dim value As String
  value = ""
  value = referenceMessMap_.GetAt(address)
  getReferencMess = value
  Exit Function
ErrorHandle:
  getReferencMess = ""
End Function

Public Sub refreshReferenceMess(ByRef ws As Worksheet, ByRef row As Integer)
 Dim colNum As Integer
 Dim address As String
 Dim iRange As range
 For colNum = 2 To ws.range("IV" + CStr(row)).End(xlToLeft).column
  address = Replace(ws.Cells(row, colNum).address, "$", "")
  If isReferenced(address) Then
      If (ws.Cells(row, colNum).Interior.colorIndex = NullPattern And ws.Cells(row, colNum).Interior.Pattern = NullPattern) Or ws.Cells(row, colNum).Interior.colorIndex = HyperLinkColorIndex Then
            Set iRange = ws.range(address)
            Call setStyle(iRange, getReferencMess(address), getResByKey("Reference Address"), HyperLinkColorIndex)
      End If
  End If
 Next colNum
End Sub

Private Sub setStyle(ByRef iRange As range, ByRef text As String, title As String, colorIndex As Long)
        Call addValidation(iRange)
        With iRange.Validation
                .inputTitle = title
                .inputmessage = text
                .ShowInput = True
                .ShowError = False
        End With
        With iRange.Interior
            .colorIndex = colorIndex
        End With
End Sub

Private Sub addValidation(iRange As range)
On Error Resume Next
        With iRange.Validation
                .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertInformation
        End With
End Sub

Private Sub Class_Initialize()
    Set referenceMessMap_ = New CMap
End Sub

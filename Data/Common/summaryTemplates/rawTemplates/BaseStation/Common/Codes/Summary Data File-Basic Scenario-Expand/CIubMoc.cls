VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIubMoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mocName_ As String '在IUB页签中的A列名称，从SheetDef中获取
Private mocSelectedFlag_ As String '是否选定的标志，未选定用字符串N来标识，与SheetDef中的定义一致
Private mocRowNumber_ As Long 'moc所在行数
Private mocStartRowNumber_ As Long '在IUB页签的起始行
Private mocEndRowNumber_ As Long '在IUB页签的结束行

Public Sub init(ByRef mocName As String, ByRef selectedString As String, ByRef rowNumber As Long, ByRef startRow As Long, ByRef endRow As Long)
    mocName_ = mocName
    mocSelectedFlag_ = selectedString
    mocRowNumber_ = rowNumber
    mocStartRowNumber_ = startRow
    mocEndRowNumber_ = endRow
End Sub

Public Function isMocSelected() As Boolean
    If mocSelectedFlag_ = "N" Then
        isMocSelected = False
    Else
        isMocSelected = True
    End If
End Function

Public Function getRowNumber() As Long
    getRowNumber = mocRowNumber_
End Function

Property Get startRow() As Long
    startRow = mocStartRowNumber_
End Property

Property Get endRow() As Long
    endRow = mocEndRowNumber_
End Property



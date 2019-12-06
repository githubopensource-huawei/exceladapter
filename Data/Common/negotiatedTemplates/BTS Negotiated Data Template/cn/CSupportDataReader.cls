VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSupportDataReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_strPathDataSource As String

' Excel文件路径所在的Cell位置
Dim m_nPathDataSourceRow As Integer
Dim m_nPathDataSourceColumn As Integer

Dim m_nDisplayNameRow As Integer
Dim m_nDetailNameRow As Integer

Dim m_strParaListSheetName As String


Const C_Row_FieldName As Integer = 3
Const C_Row_ObjId As Integer = 1
Const C_Row_ObjName As Integer = 2

Const C_DstSheet_Row_FieldName As Integer = 2
Const C_DstSheet_Row_DetailName As Integer = 6
Const C_DstSheet_Row_DisplayName As Integer = 3


Enum SDR_Language
    SDR_Language_English = 1
    SDR_Language_Chinese = 2
End Enum

Public Sub InitDataSource(language As SDR_Language)
  
  ' 初始化
  If (language = SDR_Language_English) Then
    m_nPathDataSourceRow = 1
    m_nPathDataSourceColumn = 6
    m_nDisplayNameRow = 16
    m_nDetailNameRow = 17
    m_strParaListSheetName = "Parameter List"
  End If
  
  If (language = SDR_Language_Chinese) Then
    m_nPathDataSourceRow = 2
    m_nPathDataSourceColumn = 6
    m_nDisplayNameRow = 14
    m_nDetailNameRow = 19
    m_strParaListSheetName = "参数列表"
  End If
  

  ' 读取
  Dim sheetTableDef As Worksheet
  Dim sheetDataSource As Worksheet
  
  Set sheetTableDef = GetTableDefSheet()
  
  m_strPathDataSource = Trim(sheetTableDef.Cells(m_nPathDataSourceRow, m_nPathDataSourceColumn).Value)
  
  Set sheetDataSource = GetDataSourceSheet()
  
  Call doWrite(sheetDataSource, sheetTableDef)
  
    
  
 
End Sub

Private Sub doWrite(sheetDataSource As Worksheet, sheetTableDef As Worksheet)
  
  Dim strFieldName As String
  Dim strObjName As String
  Dim strObjId As String
  Dim nRowIndex As Integer
    
  nRowIndex = 15
  Do
    '校验字段是否存在
    strFieldName = Trim(sheetTableDef.Cells(nRowIndex, C_Row_FieldName).Value)
    If strFieldName = "" Then
      Exit Do
    End If
    
    strObjId = Trim(sheetTableDef.Cells(nRowIndex, C_Row_ObjId).Value)
    If Not ("" = strObjId) Then
      strObjName = Trim(sheetTableDef.Cells(nRowIndex, C_Row_ObjName).Value)
    End If
    
    Application.StatusBar = "Processing : " + strObjName + "-" + strFieldName
    
    Call doWriteDescription(sheetTableDef, nRowIndex, sheetDataSource, strObjName, strFieldName)
    
    nRowIndex = nRowIndex + 1
  Loop

End Sub


Private Function GetTableDefSheet() As Worksheet
  Set GetTableDefSheet = ThisWorkbook.Worksheets("TableDef")
End Function

Private Function GetDataSourceSheet() As Worksheet
  Application.DisplayAlerts = False
  Workbooks.Open FileName:=m_strPathDataSource, ReadOnly:=False
  Application.DisplayAlerts = True
  Set GetDataSourceSheet = ActiveWorkbook.Worksheets(m_strParaListSheetName)
End Function

Private Sub doWriteDescription(sheetTableDef As Worksheet, srcCurrentRow As Integer, sheetDataSource As Worksheet, strObjName As String, strFieldName As String)
  
  Dim nRowIndex As Long
  Dim strMocName As String
    
  For nRowIndex = 2 To 65535
    strMocName = Trim(sheetDataSource.Cells(nRowIndex, 1).Value)
    If "" = strMocName Then
      Exit Sub
    End If

    'Application.StatusBar = "Processing : " + strObjName + "-" + strFieldName + "  Line: " + Str(nRowIndex)
    
    If (UCase(strFieldName) = UCase(sheetDataSource.Cells(nRowIndex, C_DstSheet_Row_FieldName).Value)) And UCase(strObjName) = UCase(strMocName) Then
      sheetTableDef.Cells(srcCurrentRow, m_nDetailNameRow) = sheetDataSource.Cells(nRowIndex, C_DstSheet_Row_DetailName).Value
      sheetTableDef.Cells(srcCurrentRow, m_nDisplayNameRow) = sheetDataSource.Cells(nRowIndex, C_DstSheet_Row_DisplayName).Value
      Exit Sub
    End If
    
  Next
  
End Sub




















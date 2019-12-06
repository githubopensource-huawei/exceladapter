Attribute VB_Name = "SupportInfo"

Public g_XCTConn As New Connection


'Attribute VB_Name = "PublicFun"
'该模块为公共函数库
'连接Sybase数据库
Public Function ConnectDatabase(Server As String, DataBase As String, UserName As String, Password As String, conn As Connection)
On Error GoTo ErrHandler
  If conn.State = adStateOpen Then
    conn.Close
  End If
  strConn = "Provider=Sybase OLEDB Provider;Persist Security Info=False;Initial Catalog=" + Trim(DataBase) + ";Data Source=" + Trim(Server) + ";User Id= " + Trim(UserName) + ";Password=" + Trim(Password) + ";"
  conn.Open strConn
 ConnectDatabase = True
 Exit Function
ErrHandler:
  ConnectDatabase = False
End Function
'连接SQL Server数据库
Public Function ConnectDatabaseSQL(Server As String, DataBase As String, UserName As String, Password As String, conn As Connection)
On Error GoTo ErrHandler
  If conn.State = adStateOpen Then
    conn.Close
  End If
  strConn = "driver={SQL SERVER};server=" + Trim(Server) + ";database=" + Trim(DataBase) + ";User Id= " + Trim(UserName) + ";Password=" + Trim(Password) + ";"
  conn.Open strConn
  ConnectDatabaseSQL = True
  Exit Function
ErrHandler:
  ConnectDatabaseSQL = False
End Function

Public Function ConnectMySQLDatabase(Server As String, sDataBase As String, UserName As String, Password As String, conn As Connection)
On Error GoTo ErrHandler
  If conn.State = adStateOpen Then
    conn.Close
  End If
  strConn = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=" + Trim(Server) + ";database=" + sDataBase + _
  ";UID=" + Trim(UserName) + ";PASSWORD=" + Trim(Password) + ";OPTION=3;stmt=SET NAMES GB2312"
  'strConn = strConn + "PORT=3306;OPTION=3;"
  conn.Open strConn
  ConnectMySQLDatabase = True
  Exit Function
ErrHandler:
  ConnectMySQLDatabase = False
End Function


'释放数据库链接
Public Sub ReleaseConnection(conn As ADODB.Connection)
On Error Resume Next
  conn.Close
  Set conn = Nothing
End Sub

Public Sub UpdateSupportInfo()
  Dim CurSheet As Worksheet
  Set CurSheet = Sheets("TableDef")
  
  Dim iRow As Integer

  Dim iColumnIndex As Integer
  
  Dim sVersion As String
  Dim sSheetName As String
  Dim sMoName As String
  Dim sFieldName As String
    
  Dim sRange As String
  sRange = "E" + CStr(StartTblDataRow) + ":K" + CStr(StartTblDataRow + CInt(CurSheet.Cells(5, 7)))
  Range(sRange).ClearContents
  
  sRange = "O" + CStr(StartTblDataRow) + ":W" + CStr(StartTblDataRow + CInt(CurSheet.Cells(5, 7)))
  Range(sRange).ClearContents
  
  sVersion = CurSheet.Cells(5, 8)
  
  For iRow = StartTblDataRow To StartTblDataRow + CInt(CurSheet.Cells(5, 7)) - 1
    sSheetName = CurSheet.Cells(iRow, SheetNameCol)
    
    If CurSheet.Cells(iRow, iSpecialMoName) = "" Then
      sMoName = CurSheet.Cells(iRow, iMocName + 2)
    Else
      sMoName = CurSheet.Cells(iRow, iSpecialMoName)
    End If
  
    sFieldName = CurSheet.Cells(iRow, iColumnFieldName + 1)
    
    CurSheet.Cells(iRow, iMapTableName + 1).Value = "t_tmp_" + sSheetName + "_" + CurSheet.Cells(iRow, iMocName + 2) + "_" + sVersion
    CurSheet.Cells(iRow, iMapFieldName + 1).Value = sFieldName
    
    Call GenerateColumnInfo(CurSheet, iRow, sVersion, sMoName, sFieldName)
    
    Call GenerateRangeValue(CurSheet, iRow, sVersion, sMoName, sFieldName)
  
    Call GenerateFieldColumn(CurSheet, iRow, iColumnIndex)
    
    CurSheet.Cells(iRow, iColumnWidth + 1) = Int(Len(Trim(CurSheet.Cells(iRow, iFieldDisplayName_ENG + 1))) * 0.6 + 6)
      
  Next
  
  Call GenerateDspChsName(CurSheet, sVersion)
End Sub

Public Sub GenerateColumnInfo(CurSheet As Worksheet, iRow As Integer, sVersion As String, sMoName As String, sFieldName As String)
    Dim iFieldType As Integer
    iFieldType = -1
    
    Dim sSQL As String
    sSQL = "select iFieldType, sDspName, iMustGive from view_FieldAllInfo where sVersion = '" + Trim(sVersion) + "' and sTableName = '" + Trim(sMoName) + "' and sFieldName = '" + Trim(sFieldName) + "' and iMode = 2 "

    Dim rs As Recordset
    Set rs = CreateObject("ADODB.RecordSet")
      
    rs.Open sSQL, g_XCTConn
    While Not rs.EOF
      CurSheet.Cells(iRow, iFieldDisplayName_ENG + 1).Value = rs("sDspName")
      CurSheet.Cells(iRow, iFieldPostil + 1).Value = rs("sDspName")
      CurSheet.Cells(iRow, iCheckNull + 1).Value = 1 - CInt(rs("iMustGive"))
      iFieldType = CInt(rs("iFieldType"))
      rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
    
    If sMoName = "M3LKS" And sFieldName = "SIGLKSX" Then  ' 非必给参数
      CurSheet.Cells(iRow, iCheckNull + 1).Value = "1"
    End If
      
    If sMoName = "AAL2RT" And sFieldName = "RTX" Then  ' 非必给参数
      CurSheet.Cells(iRow, iCheckNull + 1).Value = "0"
    End If
    
    If iFieldType = 0 Or iFieldType = 11 Then                  ' int and bigint
      CurSheet.Cells(iRow, iColumnType + 1).Value = "INT"
      CurSheet.Cells(iRow, iColumnType2 + 1).Value = "INT"
    ElseIf iFieldType = 2 Then
      CurSheet.Cells(iRow, iColumnType + 1).Value = "STRING"   ' string
      CurSheet.Cells(iRow, iColumnType2 + 1).Value = "STRING"
    ElseIf iFieldType = 101 Then
      CurSheet.Cells(iRow, iColumnType + 1).Value = "LIST"     'enum
      CurSheet.Cells(iRow, iColumnType2 + 1).Value = "LIST"
    ElseIf iFieldType = 5 Then
      CurSheet.Cells(iRow, iColumnType + 1).Value = "STRING"     'bitmap
      CurSheet.Cells(iRow, iColumnType2 + 1).Value = "BITMAP"
    ElseIf iFieldType = 6 Then
      CurSheet.Cells(iRow, iColumnType + 1).Value = "STRING"     'ip4
      CurSheet.Cells(iRow, iColumnType2 + 1).Value = "IP"
    End If
End Sub

Public Sub ExecuteInitvalPython(sFileName As String)
    Dim TextLine As String
    Dim FileNum As Integer
    FileNum = FreeFile()

    Open sFileName For Input As #FileNum
    
    Do While Not EOF(FileNum)    ' 循环至文件尾。
        Line Input #FileNum, TextLine    ' 读入一行数据并将其赋予某变量。
        If Trim(TextLine) <> "go" And Trim(TextLine) <> "" Then
          g_XCTConn.Execute TextLine
        End If
    Loop
    Close #FileNum    ' 关闭文件。
End Sub

Public Sub GenerateDspChsName(CurSheet As Worksheet, sVersion As String)
    ExecuteInitvalPython (CurSheet.Cells(4, 15)) 'build chs support info.
  
    For iRow = StartTblDataRow To StartTblDataRow + CInt(CurSheet.Cells(5, 7)) - 1
     If CurSheet.Cells(iRow, iSpecialMoName) = "" Then
        sMoName = CurSheet.Cells(iRow, iMocName + 2)
      Else
        sMoName = CurSheet.Cells(iRow, iSpecialMoName)
      End If
    
      sFieldName = CurSheet.Cells(iRow, iColumnFieldName + 1)
      
      Dim sSQL As String
      sSQL = "select sDspName from view_FieldAllInfo where sVersion = '" + Trim(sVersion) + "' and sTableName = '" + Trim(sMoName) + "' and sFieldName = '" + Trim(sFieldName) + "' and iMode = 2 "
    
      Dim rs As Recordset
      Set rs = CreateObject("ADODB.RecordSet")
        
      rs.Open sSQL, g_XCTConn
      While Not rs.EOF
        CurSheet.Cells(iRow, iFieldDisplayName_CHS + 1).Value = rs("sDspName")
        rs.MoveNext
      Wend
      rs.Close
      Set rs = Nothing
    Next
    
    ExecuteInitvalPython (CurSheet.Cells(3, 15)) 'recover english environment
End Sub

Public Sub GenerateRangeValue(CurSheet As Worksheet, iRow As Integer, sVersion As String, sMoName As String, sFieldName As String)
    If CurSheet.Cells(iRow, iColumnType2 + 1) = "INT" Or CurSheet.Cells(iRow, iColumnType2 + 1) = "STRING" Then
      Call GenerateIntRange(CurSheet, iRow, sVersion, sMoName, sFieldName)
    ElseIf CurSheet.Cells(iRow, iColumnType2 + 1) = "LIST" Then
      Call GenerateListValue(CurSheet, iRow, sVersion, sMoName, sFieldName)
    ElseIf CurSheet.Cells(iRow, iColumnType2 + 1) = "BITMAP" Then
      Call GenerateBitmapRange(CurSheet, iRow, sVersion, sMoName, sFieldName)
    ElseIf CurSheet.Cells(iRow, iColumnType2 + 1) = "IP" Then
      CurSheet.Cells(iRow, iMin + 1).Value = 7
      CurSheet.Cells(iRow, iMax + 1).Value = 15
    End If
End Sub

Public Sub GenerateIntRange(CurSheet As Worksheet, iRow As Integer, sVersion As String, sMoName As String, sFieldName As String)
    Dim sSQL As String, sMinRange As String, sMaxRange As String
    Dim rs As Recordset

    sSQL = "select iMinValue, iMaxValue from view_FieldRange where sVersion = '" + Trim(sVersion) + "' and sTableName = '" + Trim(sMoName) + "' and sFieldName = '" + Trim(sFieldName) + "' and iMode = 2 "
   
    Set rs = CreateObject("ADODB.RecordSet")
    rs.Open sSQL, g_XCTConn
    While Not rs.EOF
      sMinRange = sMinRange + "," + CStr(rs("iMinValue"))
      sMaxRange = sMaxRange + "," + CStr(rs("iMaxValue"))
      rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
    
    If Trim(CurSheet.Cells(iRow, iManualValue)) <> "" Then
      Dim sValue As String, iValues
      sValue = CurSheet.Cells(iRow, iManualValue)
      iValues = Split(sValue, ",")
      CurSheet.Cells(iRow, iMin + 1).Value = iValues(0)
      CurSheet.Cells(iRow, iMax + 1).Value = iValues(1)
    Else
      CurSheet.Cells(iRow, iMin + 1).Value = Right(sMinRange, Len(sMinRange) - 1)
      CurSheet.Cells(iRow, iMax + 1).Value = Right(sMaxRange, Len(sMaxRange) - 1)
    End If
End Sub

Public Sub GenerateListValue(CurSheet As Worksheet, iRow As Integer, sVersion As String, sMoName As String, sFieldName As String)
    Dim sListValue As String
    Dim sSQL As String
    Dim rs As Recordset
    
    sSQL = "select sInput from view_FieldEnum where sVersion = '" + Trim(sVersion) + "' and sTableName = '" + Trim(sMoName) + "' and sFieldName = '" + Trim(sFieldName) + "' and iMode = 2 "
   
    Set rs = CreateObject("ADODB.RecordSet")
    rs.Open sSQL, g_XCTConn
    While Not rs.EOF
      sListValue = sListValue + "," + rs("sInput")
      rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
    
    If Trim(CurSheet.Cells(iRow, iManualValue)) <> "" Then
      CurSheet.Cells(iRow, iListValue + 1).Value = CurSheet.Cells(iRow, iManualValue)
    Else
      CurSheet.Cells(iRow, iListValue + 1).Value = Right(sListValue, Len(sListValue) - 1)
    End If
End Sub

Public Sub GenerateBitmapRange(CurSheet As Worksheet, iRow As Integer, sVersion As String, sMoName As String, sFieldName As String)
    Dim sListValue As String
    Dim sSQL As String
    Dim rs As Recordset
    
    sSQL = "select count(1) as bitNumber from view_FieldEnum where sVersion = '" + Trim(sVersion) + "' and sTableName = '" + Trim(sMoName) + "' and sFieldName = '" + Trim(sFieldName) + "' and iMode = 2 "
   
    Set rs = CreateObject("ADODB.RecordSet")
    rs.Open sSQL, g_XCTConn
    While Not rs.EOF
      CurSheet.Cells(iRow, iMin + 1).Value = rs("bitNumber")
      CurSheet.Cells(iRow, iMax + 1).Value = rs("bitNumber")
      rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
End Sub
Function ConvertToLetter(iCol As Integer) As String
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   If iAlpha > 0 Then
      ConvertToLetter = Chr(iAlpha + 64)
   End If
   If iRemainder > 0 Then
      ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
   End If
End Function

Public Sub GenerateFieldColumn(CurSheet As Worksheet, iRow As Integer, iColumnIndex As Integer)
    
    If CurSheet.Cells(iRow, iMocName + 2) <> CurSheet.Cells(iRow - 1, iMocName + 2) Then
      iColumnIndex = 2
    Else
      iColumnIndex = iColumnIndex + 1
    End If
    
    CurSheet.Cells(iRow, iFieldBeginColumn + 1) = ConvertToLetter(iColumnIndex)
    CurSheet.Cells(iRow, iFieldEndColumn + 1) = CurSheet.Cells(iRow, iFieldBeginColumn + 1)
End Sub


Attribute VB_Name = "SetValidDef"
Public GeneratingFlag As Integer  '0表示正在生成
Public Sub BrushValidDef()

'Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset
Dim sheetTableDef As Worksheet
Dim sheetTableList As Worksheet
Dim sheetValidDef As Worksheet
Dim sheetRefresh As Worksheet
Dim sFiledName As String
Dim sDspName As String
Dim sTableName As String
Dim sBranchFieldName As String
Dim oldFiledName As String
Dim oldBranchFieldName As String
Dim sValue As String
Dim iValidFlag As Integer
Dim sPostil1 As String, sPostil2 As String, sPostil As String, sPostilSign As String
Dim iRows As Integer, icount As Integer, iSign  As Integer, iSheetNum  As Integer, iRsCount  As Integer, i  As Integer
Dim sVersion As String




icount = 0
iSign = 0

On Error GoTo ErrHandler

 If conn.State = adStateClosed Then
    MsgBox "请先连接数据库。"
     Exit Sub
  End If
  
Set cmd.ActiveConnection = conn
Set sheetTableDef = ThisWorkbook.Sheets("TableDef")

Set sheetTableList = ThisWorkbook.Sheets("TableList")

Set sheetValidDef = ThisWorkbook.Sheets("ValidDef")

Set sheetRefresh = ThisWorkbook.Sheets("Refresh")

sVersion = Trim(sheetRefresh.Cells(1, 5).Value)
     If sVersion = "" Then
        MsgBox "网元版本编号不能为空"
            Exit Sub
       End If

' 先清空以前表里数据
   sheetValidDef.Activate
   Range("A4:J5000").Select
   Selection.Delete
   
' 设置行标
iRows = 3
For iSheetNum = 2 To 30000
''''''''对每个Sheet进行编辑
' ThisWorkbook.Sheets(sTableName).Rows(1).Clear
' ThisWorkbook.Sheets(sTableName).Rows(2).Clear

     If Trim(sheetTableList.Cells(iSheetNum, 1).Value) = "" Then
                Exit For
       End If
       
sTableName = Trim(sheetTableList.Cells(iSheetNum, 1).Value)



'***分支控制参数关系查询***
Application.StatusBar = "正在取得表" + sTableName + "分支参数数据，请稍候……"
cmd.CommandText = "SELECT  D.sVersion,D.iTableId,E.sTableName,E.sFieldName,E.iValue,E.sInput,D.iValidFlag,D.iFieldId,c.sFieldName iBranchFieldName,c.iFieldType,D.iSign,D.iFatherBranchId,D.iBranchId,D.iMode  FROM    Utils_BranchDef AS D,view_FieldEnum  AS E ,view_FieldAllInfo c  WHERE   E.sVersion=D.sVersion AND   E.iTableId=D.iTableId AND   E.iFieldId=D.iBranchFieldId AND     D.sVersion='" + sVersion + "'" + " and D.sBranchMinValue=E.iValue AND  D.sBranchMaxValue=E.iValue AND D.iMode=E.iMode AND D.iValidFlag=0 AND D.iMode=2  and c.sVersion =D.sVersion  AND c.iTableId=D.iTableId   AND c.iFieldId=D.iFieldId   AND c.iMode = D.iMode and c.iVisible  = 1  and E.sTableName='" + sTableName + "' order by D.iTableId,D.iFieldId,D.sBranchMinValue "

rs.CursorLocation = adUseClient
rs.Open cmd
Application.StatusBar = ""
oldBranchFieldName = ""
oldFiledName = ""
Do While Not rs.EOF
    sFiledName = rs("sFieldName")
   sBranchFieldName = rs("iBranchFieldName")
   If oldBranchFieldName <> sBranchFieldName Or oldFiledName <> sFiledName Then
      iRows = iRows + 1
      sheetValidDef.Rows(iRows).Clear
      oldBranchFieldName = sBranchFieldName
      oldFiledName = sFiledName
    End If
    sValue = sheetValidDef.Cells(iRows, 6).Value
 ' 不清楚当前编辑行
   If sValue <> "" Then
      sValue = sValue + "," + rs("sInput")
     End If
   If sValue = "" Then
      sValue = rs("sInput")
     End If
    sheetValidDef.Cells(iRows, 1) = sTableName
    sheetValidDef.Cells(iRows, 2) = sFiledName
    sheetValidDef.Cells(iRows, 6) = sValue
    sheetValidDef.Cells(iRows, 7) = sBranchFieldName
    sheetValidDef.Cells(iRows, 9) = "NO"
    rs.MoveNext
    Loop
    rs.Close
    
    
 

''''''''对每个Sheet进行编辑
Next iSheetNum

'获取分支列信息
GetFieldColRow

MsgBox "OK"
 
 Exit Sub

ErrHandler:
  'ConnectDatabase = False
End Sub
Public Sub GetFieldColRow()

Dim sheetTableDef As Worksheet
Dim sheetValidDef As Worksheet
Dim sFiledName As String
Dim sTableName As String
Dim sTableNameDef As String
Dim sBranchFieldName As String
Dim iRows As Integer, iSheetNum As Integer

Set sheetTableDef = ThisWorkbook.Sheets("TableDef")
Set sheetValidDef = ThisWorkbook.Sheets("ValidDef")

For iSheetNum = 4 To 30000
     If Trim(sheetValidDef.Cells(iSheetNum, 7).Value) = "" Then
                Exit For
       End If
       
sTableName = Trim(sheetValidDef.Cells(iSheetNum, 1).Value)
sFiledName = Trim(sheetValidDef.Cells(iSheetNum, 2).Value)
sBranchFieldName = Trim(sheetValidDef.Cells(iSheetNum, 7).Value)
  For iRows = 15 To 30000
       If Trim(sheetTableDef.Cells(iRows, 3).Value) = "" Then
           Exit For
       End If
       If Trim(sheetTableDef.Cells(iRows, 1).Value) <> "" Then
            sTableNameDef = Trim(sheetTableDef.Cells(iRows, 2).Value)
       End If
       If sTableName = sTableNameDef And sBranchFieldName = Trim(sheetTableDef.Cells(iRows, 3).Value) Then
          sheetValidDef.Cells(iSheetNum, 8) = Trim(sheetTableDef.Cells(iRows, 5).Value)
          iRows = 30000
        End If
    Next iRows

''''''''对每个Sheet进行编辑
Next iSheetNum

End Sub


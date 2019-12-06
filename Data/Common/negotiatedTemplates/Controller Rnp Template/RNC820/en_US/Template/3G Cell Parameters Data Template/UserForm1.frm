VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function Check(ServerName As String, DataBase As String, UserName As String) As Boolean
    Dim strErr As String
    If ServerName = "" Then
        strErr = strErr & "请输入Server机器名或IP地址!" & vbCrLf
    End If
    If DataBase = "" Then
        strErr = strErr & "请数据库名!" & vbCrLf
    End If
    If UserName = "" Then
        strErr = strErr & "请输入用户名!" & vbCrLf
    End If
        
    If strErr <> "" Then
        MsgBox strErr, vbCritical
        Exit Function
    End If
    Check = True
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo HandleErr
    '首先检查输入数据的正确型
    If Not Check(TextBox1.Text, TextBox2.Text, TextBox3.Text) Then
        Exit Sub
    End If
    
    '连接到数据库
    If optSybase.Value Then
        If Not ConnectDatabase(TextBox1.Text, TextBox2.Text, TextBox3.Text, TextBox4.Text, TextBox5.Text, conn) Then
             MsgBox "连接数据库失败", vbCritical
              Exit Sub
          Else: MsgBox "连接数据库成功"
         End If
      End If
      
    If optSQL.Value Then
        If Not ConnectDatabaseSQL(TextBox1.Text, TextBox3.Text, TextBox4.Text, TextBox5.Text, conn) Then
             MsgBox "连接数据库失败", vbCritical
              Exit Sub
          Else: MsgBox "连接数据库成功"
         End If
      End If
    Me.Hide
        
    Exit Sub
HandleErr:
    MsgBox Err.Description
End Sub


Private Sub optSybase_Click()

End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub UserForm_Terminate()
  ReleaseConnection conn
End Sub
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


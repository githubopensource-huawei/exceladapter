VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} XCT_100 
   Caption         =   "Connect to cmedb"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   3840
   OleObjectBlob   =   "XCT_100.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "XCT_100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function Check(ServerName As String, DataBase As String, UserName As String) As Boolean
    Dim strErr As String
    If ServerName = "" Then
        strErr = strErr & "请输入XCT的Server机器名或IP地址!" & vbCrLf
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

Private Sub cmdOK_Click()
On Error GoTo HandleErr
    '首先检查输入数据的正确型
    If Not Check(TextBox1.Text, TextBox2.Text, TextBox3.Text) Then
        Exit Sub
    End If
    
    '连接到数据库
    If optSybase.Value Then
      If Not ConnectDatabase(TextBox1.Text, TextBox2.Text, TextBox3.Text, TextBox4.Text, g_XCTConn) Then
          MsgBox "连接XCT数据库失败，请注意部分功能需要连接数据库！", vbCritical
           'Exit Sub
      End If
    Else  'SQL Server
      If Not ConnectDatabaseSQL(TextBox1.Text, TextBox2.Text, TextBox3.Text, TextBox4.Text, g_XCTConn) Then
          MsgBox "连接XCT数据库失败", vbCritical
          Exit Sub
      End If
    End If
    Me.Hide

    Call UpdateSupportInfo
    Exit Sub
HandleErr:
    MsgBox Err.Description
End Sub

Private Sub initDBParameter(DBType As Integer)
  TextBox1.Text = IIf(DBType = 0, "127.0.0.1", "10.141.142.10,4100")
  TextBox2.Text = "cmedb"
  TextBox3.Text = "sa"
  TextBox4.Text = IIf(DBType = 0, "emsems", "emsems")
End Sub

Private Sub optSQL_Click()
  initDBParameter 0
End Sub

Private Sub optSybase_Click()
  initDBParameter 1
End Sub

Private Sub UserForm_Initialize()
    Me.optSybase.Value = True
    optSybase_Click
End Sub

Private Sub UserForm_Terminate()
  ReleaseConnection g_XCTConn
End Sub



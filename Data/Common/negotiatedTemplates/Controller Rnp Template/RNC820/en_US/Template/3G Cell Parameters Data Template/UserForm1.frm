VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function Check(ServerName As String, DataBase As String, UserName As String) As Boolean
    Dim strErr As String
    If ServerName = "" Then
        strErr = strErr & "������Server��������IP��ַ!" & vbCrLf
    End If
    If DataBase = "" Then
        strErr = strErr & "�����ݿ���!" & vbCrLf
    End If
    If UserName = "" Then
        strErr = strErr & "�������û���!" & vbCrLf
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
    '���ȼ���������ݵ���ȷ��
    If Not Check(TextBox1.Text, TextBox2.Text, TextBox3.Text) Then
        Exit Sub
    End If
    
    '���ӵ����ݿ�
    If optSybase.Value Then
        If Not ConnectDatabase(TextBox1.Text, TextBox2.Text, TextBox3.Text, TextBox4.Text, TextBox5.Text, conn) Then
             MsgBox "�������ݿ�ʧ��", vbCritical
              Exit Sub
          Else: MsgBox "�������ݿ�ɹ�"
         End If
      End If
      
    If optSQL.Value Then
        If Not ConnectDatabaseSQL(TextBox1.Text, TextBox3.Text, TextBox4.Text, TextBox5.Text, conn) Then
             MsgBox "�������ݿ�ʧ��", vbCritical
              Exit Sub
          Else: MsgBox "�������ݿ�ɹ�"
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
'����SQL Server���ݿ�
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


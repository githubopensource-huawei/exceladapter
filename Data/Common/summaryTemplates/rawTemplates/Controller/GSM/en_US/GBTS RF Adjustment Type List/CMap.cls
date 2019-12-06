VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolKey As Collection '关键字集合
Private mcolKey2Data As Collection  '数据的集合，由Key访问
Private mMulKeySep As String '多键值之间的分隔符

Public Sub Clean()
    Set mcolKey = New Collection
    Set mcolKey2Data = New Collection
End Sub

Public Function GetAt(ByVal strKey, Optional Default) As Variant
    On Error Resume Next
    Call EncrptKey(strKey)
    GetAt = mcolKey2Data(strKey)
    If Err <> 0 Then
        GetAt = Default
        Err.Clear
    End If
End Function

Public Sub SetAt(ByVal strKey, ByVal vVal)
    Call EncrptKey(strKey)
    If hasKey(strKey) Then
        mcolKey2Data.Remove strKey
    Else
        mcolKey.Add strKey, strKey
    End If

    mcolKey2Data.Add vVal, strKey
End Sub

Public Function hasKey(ByVal strKey) As Boolean
    Dim v
    
    On Error Resume Next
    Call EncrptKey(strKey)
    v = mcolKey2Data(strKey)
    hasKey = (Err = 0)
End Function

Public Function MakeKey(ParamArray key() As Variant) As String
    Dim colX As New Collection
    Dim I As Long
    
    For I = 0 To UBound(key)
        colX.Add key(I)
    Next I

    MakeKey = MakeKeyByCol(colX)
End Function

Public Function MakeKeyByCol(colX As Collection) As String
    Dim I As Long
    Dim s As String
    
    s = colX(1)
    For I = 2 To colX.count
        s = s & mMulKeySep & colX(I)
    Next I
    Call EncrptKey(s)
    
    MakeKeyByCol = s
End Function

Public Property Get KeyCollection() As Collection
    Dim colKey As New Collection
    Dim key
    
    For Each key In mcolKey
        DisEncrptKey key
        colKey.Add key
    Next key
   
    Set KeyCollection = colKey
End Property

Public Property Get ValueCollection() As Collection
    Set ValueCollection = mcolKey2Data
End Property

Public Property Get count() As Long
    count = mcolKey.count
End Property

Public Property Let MulKeySep(ByVal s As String)
    mMulKeySep = s
End Property

Private Sub Class_Initialize()
    Me.Clean
    mMulKeySep = "\"
End Sub

Public Sub RemoveKey(ByVal key)
    On Error Resume Next
    Call EncrptKey(key)
    mcolKey.Remove key
    mcolKey2Data.Remove key
End Sub

Public Property Get isEmpty() As Boolean
    isEmpty = (Me.count <= 0)
End Property

Private Sub EncrptKey(key)
    If IsNumeric(key) Then key = Chr(8) & key
End Sub

Private Sub DisEncrptKey(key)
    If Left(key, 1) = Chr(8) Then key = Mid(key, 2)
End Sub


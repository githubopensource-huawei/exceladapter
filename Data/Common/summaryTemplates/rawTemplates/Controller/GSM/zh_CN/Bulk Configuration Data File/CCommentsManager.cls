VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCommentsManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sheetName As String
Private m_sheetType As String
Private m_comments As Collection
Private m_iubKeyMapping As Collection
        
' get property
Property Get sheetName() As String
    sheetName = m_sheetName
End Property

Property Get sheetType() As String
    sheetType = m_sheetType
End Property

Property Get comments() As Collection
    Set comments = m_comments
End Property


' set property
Property Let sheetName(shtName As String)
    m_sheetName = shtName
End Property

Property Let sheetType(shtType As String)
    m_sheetType = shtType
End Property




Public Sub insertComment(grpName As String, attrName As String, comment As String)
    On Error GoTo ErrorHandler
    Dim key As String
    key = grpName & "_" & attrName
    If Contains(m_comments, key) Then m_comments.Remove (key)
    m_comments.Add Item:=comment, key:=key
    
    Dim iubKey As String
    iubKey = getIubKey(attrName)
    If Contains(m_iubKeyMapping, iubKey) Then m_iubKeyMapping.Remove (iubKey)
    m_iubKeyMapping.Add Item:=key, key:=iubKey
    
    Exit Sub
ErrorHandler:
    Debug.Print "some exception in CommentsManager::insertComment, " & Err.Description
End Sub

Public Function getKey(grpName As String, attrName As String) As String
    getKey = grpName & "_" & attrName
End Function

Public Function getIubKey(attrName As String) As String
    getIubKey = attrName
End Function

Public Function hasKey(key As String) As Boolean
    On Error GoTo ErrorHandler
    hasKey = True
    If Contains(m_comments, key) Then Exit Function
    
    hasKey = False
    
    Exit Function
ErrorHandler:
    hasKey = False
    Debug.Print "some exception in " & Err.Description
End Function

Public Function hasIubKey(iubKey As String, Optional realKey As String) As Boolean
    On Error GoTo ErrorHandler
    hasIubKey = True
    If Contains(m_iubKeyMapping, iubKey) Then
        realKey = m_iubKeyMapping(iubKey)
        If hasKey(realKey) Then Exit Function
    End If
    
    hasIubKey = False
    
    Exit Function
ErrorHandler:
    hasIubKey = False
    Debug.Print "some exception in " & Err.Description
End Function


Private Sub Class_Initialize()
    Set m_comments = Nothing
    Set m_comments = New Collection
    Set m_iubKeyMapping = Nothing
    Set m_iubKeyMapping = New Collection
    m_sheetName = ""
    m_sheetType = ""
End Sub

Private Sub Class_Terminate()
    Set m_comments = Nothing
    Set m_iubKeyMapping = Nothing
End Sub

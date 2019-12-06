VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClinkType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Private m_linkSheetName As String '被链接Cell的sheet名称
    Private m_linkGroupName As String '被链接group的名称
    Private m_linkColumName As String '被链接列的名称
    Private m_linkRowNum As Long ' 被链接的cell的实际行号
    Private m_linkColNum As Long '被链接cell的实际列号
    
    Private m_linkColumRowNum As Long ' 被链接的的列头所在的位置
    
    Private m_sheetName As String '引用cell所在的sheet
    Private m_groupName As String
    Private m_columName As String
    
    Private m_rowNum As Long '引用cell所在的行
    Private m_colNum As Long '引用cell所在的列
    Private m_address As String
Property Get GroupName() As String
        GroupName = m_groupName
End Property
Property Let columName(str As String)
        m_columName = str
End Property

Property Get columName() As String
        columName = m_columName
End Property

Property Let GroupName(str As String)
        m_groupName = str
End Property

Property Get linkColumRowNum() As Long
                linkColumRowNum = m_linkColumRowNum
End Property

Property Let linkColumRowNum(row As Long)
                m_linkColumRowNum = row
End Property


Property Get linkSheetName() As String
        linkSheetName = m_linkSheetName
End Property

Property Let linkSheetName(str As String)
        m_linkSheetName = str
End Property

Property Get linkGroupName() As String
        linkGroupName = m_linkGroupName
End Property

Property Let linkGroupName(str As String)
        m_linkGroupName = str
End Property

Property Get linkColumName() As String
        linkColumName = m_linkColumName
End Property

Property Let linkColumName(str As String)
        m_linkColumName = str
End Property

Property Get linkRowNum() As Long
        linkRowNum = m_linkRowNum
End Property

Property Let linkRowNum(str As Long)
        m_linkRowNum = str
End Property

Property Get linkColNum() As Long
        linkColNum = m_linkColNum
End Property

Property Let linkColNum(str As Long)
        m_linkColNum = str
End Property

Property Get sheetName() As String
        sheetName = m_sheetName
End Property

Property Let sheetName(str As String)
        m_sheetName = str
End Property

Property Get rowNum() As Long
        rowNum = m_rowNum
End Property

Property Let rowNum(str As Long)
        m_rowNum = str
End Property

Property Get colNum() As Long
        colNum = m_colNum
End Property

Property Let colNum(str As Long)
        m_colNum = str
End Property
Property Get address() As String
        address = m_address
End Property
Property Let address(str As String)
        m_address = str
End Property

Public Function getKey() As String
    getKey = m_sheetName + ", " + m_groupName + ", " + m_columName
End Function



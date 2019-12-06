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
    Private m_linkSheetName As String '������Cell��sheet����
    Private m_linkGroupName As String '������group������
    Private m_linkColumName As String '�������е�����
    Private m_linkRowNum As Long ' �����ӵ�cell��ʵ���к�
    Private m_linkColNum As Long '������cell��ʵ���к�
    
    Private m_linkColumRowNum As Long ' �����ӵĵ���ͷ���ڵ�λ��
    
    Private m_sheetName As String '����cell���ڵ�sheet
    Private m_groupName As String
    Private m_columName As String
    
    Private m_rowNum As Long '����cell���ڵ���
    Private m_colNum As Long '����cell���ڵ���
    Private m_address As String
Property Get groupName() As String
        groupName = m_groupName
End Property
Property Let columName(str As String)
        m_columName = str
End Property

Property Get columName() As String
        columName = m_columName
End Property

Property Let groupName(str As String)
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

Property Get RowNum() As Long
        RowNum = m_rowNum
End Property

Property Let RowNum(str As Long)
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
            getKey = m_sheetName + "," + m_groupName + "," + m_columName
End Function

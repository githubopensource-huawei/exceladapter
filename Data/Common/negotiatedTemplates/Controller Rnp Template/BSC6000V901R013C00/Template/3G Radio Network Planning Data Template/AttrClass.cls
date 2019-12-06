VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AttrClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************
'��ģ�����ƣ�AttrClass
'��ģ�����ã��������RNPģ���漰��MOC��������
'***************************************************************
Option Explicit

Private strMocName As String '����MOC����

Private strName As String '��������
Private strCaption As String '������ʾ����
Private strColName As String '������ A, B, C
Private strDataType As String '�������� INT,LIST,STRING,TITLE
Private strMinValue As String '������Сֵ
Private strMaxValue As String '�������ֵ
Private strValueList As String '����ö��ֵ�б�
Private strDefaultValue As String 'ΪCT��DF����ģ������

'-------------------------------------------------------------------------------
'���²�����Ϊ�����������
'-------------------------------------------------------------------------------
Private strRefMocName As String '����MOC����
Private strRefAttrName As String '������������

'-------------------------------------------------------------------------------
'���²����Ƿ����������������������������Լ���������ģ�����AttrClass�����ʹ��
'-------------------------------------------------------------------------------
Private enumDataType As mocAttrDataType '��������ö��ֵ
Private iColIndex As Integer '������ 1, 2, 3

Private strColNameVDF As String '��ΪCvtTemplateģ������
Private iColIndexVDF As Integer '��ΪCvtTemplateģ������

Private strValue As String '�����ȡExcel

Property Get Value()
    Value = strValue
End Property

Property Let Value(Value)
    strValue = Value
End Property

Property Get ColIndexVDF()
    ColIndexVDF = iColIndexVDF
End Property

Property Let ColIndexVDF(Value)
    iColIndexVDF = Value
End Property

Property Get ColNameVDF()
    ColNameVDF = strColNameVDF
End Property

Property Let ColNameVDF(Value)
    strColNameVDF = Value
End Property

Property Get DefaultValue()
    DefaultValue = strDefaultValue
End Property

Property Let DefaultValue(Value)
    strDefaultValue = Value
End Property

Property Get MocName()
    MocName = strMocName
End Property

Property Let MocName(Value)
    strMocName = Value
End Property

Property Get RefMocName()
    RefMocName = strRefMocName
End Property

Property Let RefMocName(Value)
    strRefMocName = Value
End Property

Property Get RefAttrName()
    RefAttrName = strRefAttrName
End Property

Property Let RefAttrName(Value)
    strRefAttrName = Value
End Property

'�ж��Ƿ���������
Property Get IsVirtualAttr()
    If strRefMocName = "" Or strRefAttrName = "" Then
        IsVirtualAttr = False
    Else
        IsVirtualAttr = True
    End If
End Property

Property Get Name()
    Name = strName
End Property

Property Let Name(Value)
    strName = Value
End Property

Property Get caption()
    caption = strCaption
End Property

Property Let caption(Value)
    strCaption = Value
End Property

Property Get ColIndex()
    ColIndex = iColIndex
End Property

Property Let ColIndex(Value)
    iColIndex = Value
End Property

Property Get ColName()
    ColName = strColName
End Property

Property Let ColName(Value)
    strColName = Value
End Property

Property Get DataType()
    DataType = enumDataType
End Property

Property Let DataType(Value)
    enumDataType = Value
End Property

Property Get DataTypeName()
    DataTypeName = strDataType
End Property

Property Let DataTypeName(Value)
    strDataType = Value
End Property

Property Get MinValue()
    MinValue = strMinValue
End Property

Property Let MinValue(Value)
    strMinValue = Value
End Property

Property Get MaxValue()
    MaxValue = strMaxValue
End Property

Property Let MaxValue(Value)
    strMaxValue = Value
End Property

Property Get ValueList()
    ValueList = strValueList
End Property


Property Let ValueList(Value)
    strValueList = Value
End Property

'��¡����
Public Function Clone() As AttrClass
    Set Clone = New AttrClass
    With Clone
        .MocName = strMocName
        .Name = strName
        .caption = strCaption
        .ColName = strColName
        .DataTypeName = strDataType
        .MinValue = strMinValue
        .MaxValue = strMaxValue
        .ValueList = strValueList
        .RefMocName = strRefMocName
        .RefAttrName = strRefAttrName
        .DataType = enumDataType
        .ColIndex = iColIndex
    End With
End Function
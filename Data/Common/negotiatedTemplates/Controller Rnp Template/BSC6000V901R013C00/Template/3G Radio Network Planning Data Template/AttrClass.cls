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
'类模块名称：AttrClass
'类模块作用：方便操作RNP模板涉及的MOC及其属性
'***************************************************************
Option Explicit

Private strMocName As String '所属MOC名称

Private strName As String '属性名称
Private strCaption As String '属性显示名称
Private strColName As String '列名称 A, B, C
Private strDataType As String '属性类型 INT,LIST,STRING,TITLE
Private strMinValue As String '属性最小值
Private strMaxValue As String '属性最大值
Private strValueList As String '属性枚举值列表
Private strDefaultValue As String '为CT和DF两个模块所用

'-------------------------------------------------------------------------------
'以下参数仅为虚拟属性设计
'-------------------------------------------------------------------------------
Private strRefMocName As String '引用MOC名称
Private strRefAttrName As String '引用属性名称

'-------------------------------------------------------------------------------
'以下参数是非输入参数，是依据上述输入参数自己计算出来的，方便AttrClass对象的使用
'-------------------------------------------------------------------------------
Private enumDataType As mocAttrDataType '属性类型枚举值
Private iColIndex As Integer '列索引 1, 2, 3

Private strColNameVDF As String '仅为CvtTemplate模块所用
Private iColIndexVDF As Integer '仅为CvtTemplate模块所用

Private strValue As String '方便读取Excel

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

'判断是否虚拟属性
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

'克隆属性
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

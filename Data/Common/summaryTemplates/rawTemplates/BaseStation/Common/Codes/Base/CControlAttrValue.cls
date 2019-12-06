VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CControlAttrValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private shortName_ As String '短名称
Private longName_ As String '长名称，即页签上的显示名称
Private value_ As String '该参数的值
Private valueType_ As String '类型，string/enum/uinteger等
Private matchFlag_ As Boolean

Private valueEmptyFlag_ As Boolean '主控值是否为空的标志位
Private valueReferenceFlag_ As Boolean '主控值是否为引用的标志位
Private valueCellGrayFlag_ As Boolean '主控单元格是否为灰化的标志位

Public Sub init(ByRef shortName As String, ByRef longName As String, ByRef value As String, ByRef valueType As String, _
    ByRef valueEmptyFlag As Boolean, ByRef valueReferenceFlag As Boolean, ByRef valueCellGrayFlag As Boolean)
    shortName_ = shortName
    longName_ = longName
    value_ = value
    valueType_ = valueType
    
    valueEmptyFlag_ = valueEmptyFlag
    valueReferenceFlag_ = valueReferenceFlag
    valueCellGrayFlag_ = valueCellGrayFlag
End Sub

Public Function checkAttrValidation(ByRef equalsNode As Variant) As Boolean
    If valueCellGrayFlag_ = True Then '如果主控灰化，则该参数无效，不在范围内
        checkAttrValidation = False
        Exit Function
    ElseIf valueReferenceFlag_ = True Or (valueEmptyFlag_ = True And valueCellGrayFlag_ = False) Then '如果单元格为引用或（空且非灰化），则认定为值有效
        checkAttrValidation = True
        matchFlag_ = True
        Exit Function
    End If
    
    Dim validationFlag As Boolean
    Dim childNodes As Variant
    Set childNodes = equalsNode.childNodes
    
    '要校验的类型只有如下六种，其余全部不校验，认定在范围内
    If valueType_ = "Enum" Then
        validationFlag = checkEnumValidation(childNodes)
    ElseIf valueType_ = "UInteger" Or valueType_ = "Integer" Then
        validationFlag = checkIntegerValidation(childNodes)
    ElseIf valueType_ = "String" Or valueType_ = "Password" Or valueType_ = "ATM" Then
        validationFlag = checkStringValidation(childNodes)
    Else
        validationFlag = True
    End If
    checkAttrValidation = validationFlag
    matchFlag_ = validationFlag
End Function

'校验一个参数值是否符合Branch分支节点中的各值
Public Function checkABranchAttrValues(ByRef branchNode As Variant) As Boolean
    On Error GoTo ErrorHandler
    Dim validationFlag As Boolean
    If valueType_ = "Enum" Then
        validationFlag = checkEnumValidation(branchNode.SelectNodes("EnumItem")) 'EnumItem是枚举的一种标签
        If validationFlag = False Then
            validationFlag = checkEnumValidation(branchNode.SelectNodes("BitEnumItem")) 'BitEnumItem是枚举的第二种标签
        End If
    ElseIf valueType_ = "UInteger" Or valueType_ = "Integer" Then
        validationFlag = checkIntegerValidation(branchNode.SelectNodes("NumBoundTrait")) 'NumBoundTrait是数字类型的标签
    ElseIf valueType_ = "String" Or valueType_ = "Password" Or valueType_ = "ATM" Then
        validationFlag = checkStringValidation(branchNode.SelectNodes("StringLenTrait")) 'StringLenTrait是字符串类型的标签
    End If
    checkABranchAttrValues = validationFlag
    Exit Function
ErrorHandler: '出异常则说明输入值类型不对，校验失败
    checkABranchAttrValues = False
End Function

Private Function checkEnumValidation(ByRef childNodes As Variant) As Boolean
    checkEnumValidation = False
    Dim enumItemNode As Variant
    For Each enumItemNode In childNodes
        If enumItemNode.getAttribute("name") = value_ Then 'if enumItemNode.nodeName = "EnumItem" or "BitEnumItem" then
            checkEnumValidation = True
            Exit Function
        End If
    Next enumItemNode
End Function

Private Function checkIntegerValidation(ByRef childNodes As Variant) As Boolean
    checkIntegerValidation = False
    Dim traitNode As Variant
    Dim minValue As Double
    Dim maxValue As Double
    
    Dim value As Double
    Dim valueString As String
    value = CDbl(value_)
    valueString = CStr(value)
    '如果转换成double之后的数值还有小数点，则不是整数类型，是个小数，不在范围内
    If InStr(valueString, ".") <> 0 Then
        Exit Function
    End If
    
    For Each traitNode In childNodes
        minValue = CDbl(traitNode.getAttribute("min"))
        maxValue = CDbl(traitNode.getAttribute("max"))
        If value <= maxValue And value >= minValue Then
            checkIntegerValidation = True
            Exit Function
        End If
    Next traitNode
End Function

Private Function checkStringValidation(ByRef childNodes As Variant) As Boolean
    checkStringValidation = False
    Dim traitNode As Variant
    Dim minValue As Double
    Dim maxValue As Double
    For Each traitNode In childNodes
        minValue = CDbl(traitNode.getAttribute("min"))
        maxValue = CDbl(traitNode.getAttribute("max"))
        If LenB(StrConv(value_, vbFromUnicode)) <= maxValue And LenB(StrConv(value_, vbFromUnicode)) >= minValue Then
            checkStringValidation = True
            Exit Function
        End If
    Next traitNode
End Function

Property Get shortName() As String
    shortName = shortName_
End Property

Property Get valueEmptyFlag() As Boolean
    valueEmptyFlag = valueEmptyFlag_
End Property

Property Get valueReferenceFlag() As Boolean
    valueReferenceFlag = valueReferenceFlag_
End Property

Private Sub Class_Initialize()
    matchFlag_ = False
    valueEmptyFlag_ = False
    valueReferenceFlag_ = False
    valueCellGrayFlag_ = False
End Sub

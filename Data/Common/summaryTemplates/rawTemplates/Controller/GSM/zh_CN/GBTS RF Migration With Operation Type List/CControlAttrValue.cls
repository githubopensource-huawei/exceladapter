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

Private shortName_ As String '������
Private longName_ As String '�����ƣ���ҳǩ�ϵ���ʾ����
Private value_ As String '�ò�����ֵ
Private valueType_ As String '���ͣ�string/enum/uinteger��
Private matchFlag_ As Boolean

Private valueEmptyFlag_ As Boolean '����ֵ�Ƿ�Ϊ�յı�־λ
Private valueReferenceFlag_ As Boolean '����ֵ�Ƿ�Ϊ���õı�־λ
Private valueCellGrayFlag_ As Boolean '���ص�Ԫ���Ƿ�Ϊ�һ��ı�־λ

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
    If valueCellGrayFlag_ = True Then '������ػһ�����ò�����Ч�����ڷ�Χ��
        checkAttrValidation = False
        Exit Function
    ElseIf valueReferenceFlag_ = True Or (valueEmptyFlag_ = True And valueCellGrayFlag_ = False) Then '�����Ԫ��Ϊ���û򣨿��ҷǻһ��������϶�Ϊֵ��Ч
        checkAttrValidation = True
        matchFlag_ = True
        Exit Function
    End If
    
    Dim validationFlag As Boolean
    Dim childNodes As Variant
    Set childNodes = equalsNode.childNodes
    
    'ҪУ�������ֻ���������֣�����ȫ����У�飬�϶��ڷ�Χ��
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

'У��һ������ֵ�Ƿ����Branch��֧�ڵ��еĸ�ֵ
Public Function checkABranchAttrValues(ByRef branchNode As Variant) As Boolean
    On Error GoTo ErrorHandler
    Dim validationFlag As Boolean
    If valueType_ = "Enum" Then
        validationFlag = checkEnumValidation(branchNode.SelectNodes("EnumItem")) 'EnumItem��ö�ٵ�һ�ֱ�ǩ
        If validationFlag = False Then
            validationFlag = checkEnumValidation(branchNode.SelectNodes("BitEnumItem")) 'BitEnumItem��ö�ٵĵڶ��ֱ�ǩ
        End If
    ElseIf valueType_ = "UInteger" Or valueType_ = "Integer" Then
        validationFlag = checkIntegerValidation(branchNode.SelectNodes("NumBoundTrait")) 'NumBoundTrait���������͵ı�ǩ
    ElseIf valueType_ = "String" Or valueType_ = "Password" Or valueType_ = "ATM" Then
        validationFlag = checkStringValidation(branchNode.SelectNodes("StringLenTrait")) 'StringLenTrait���ַ������͵ı�ǩ
    End If
    checkABranchAttrValues = validationFlag
    Exit Function
ErrorHandler: '���쳣��˵������ֵ���Ͳ��ԣ�У��ʧ��
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
    For Each traitNode In childNodes
        minValue = CDbl(traitNode.getAttribute("min"))
        maxValue = CDbl(traitNode.getAttribute("max"))
        If CDbl(value_) <= maxValue And CDbl(value_) >= minValue Then
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
        If Len(value_) <= maxValue And Len(value_) >= minValue Then
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

Private Sub Class_Initialize()
    matchFlag_ = False
    valueEmptyFlag_ = False
    valueReferenceFlag_ = False
    valueCellGrayFlag_ = False
End Sub
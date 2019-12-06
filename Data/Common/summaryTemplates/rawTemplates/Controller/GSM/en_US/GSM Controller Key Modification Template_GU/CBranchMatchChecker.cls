VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBranchMatchChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private controlAttrValueManager_ As CControlAttrValueManager '�������ز������ݵĹ�����
Private controlInfo_ As String 'XML����Ϣ
Private oneBranchMatchFlag_ As Boolean '�Ƿ���һ����֧ƥ��ı�־
Private matchBranchXml_ As String 'ƥ��ķ�֧xml��Ϣ�����ڶԵ�ǰ����ֵ��У��
Private matchBranchNode_ As Variant 'ƥ���Branch�ڵ㣬�Ƿ����ú����ٿ�

Public Sub init(ByRef controlAttrValueManager As CControlAttrValueManager, ByRef controlInfo As String)
    Set controlAttrValueManager_ = controlAttrValueManager
    controlInfo_ = controlInfo
End Sub

'�õ�ƥ���֧�ڵ�
Public Function getMatchBranchNode() As Variant
    Set getMatchBranchNode = matchBranchNode_
End Function

'��ƥ���֧�����ز���ֵ�Ƿ�Ϊ�գ�Ҫ��oneBranchMatchFlag_ΪTrue
Public Function getMatchBranchAttrEmptyFlag() As Boolean
    getMatchBranchAttrEmptyFlag = False
    If oneBranchMatchFlag_ = False Then
        getMatchBranchAttrEmptyFlag = True
        Exit Function '���ƥ���֧Ϊfalse��˵��û��ƥ�䣬�ٶ�ĳ�����ز���Ϊ�գ�������
    End If
    Dim controlAttrValue As CControlAttrValue
    Dim equalsNode As Variant
    Dim attributeName As String
    
    For Each equalsNode In matchBranchNode_.GetElementsByTagName("Equals")
        attributeName = equalsNode.getAttribute("attribute") '�õ����ز�������
        Set controlAttrValue = controlAttrValueManager_.getControlAttrValue(attributeName) '�õ����ز���ֵ���͵���
        
        '��ƥ���֧��ÿ�����ز���ֵ�Ƿ�Ϊ�ջ����Ƿ������ã�����ǣ����˳����ز�������Ҫ����
        If controlAttrValue.valueEmptyFlag = True Or controlAttrValue.valueReferenceFlag = True Then
            getMatchBranchAttrEmptyFlag = True
            Exit Function
        End If
    Next equalsNode
End Function

Public Function getOneBranchMatchFlag() As Boolean
    Dim xmlDoc As Variant
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    xmlDoc.LoadXML (controlInfo_)
    
    Dim root As Variant
    Set root = xmlDoc.DocumentElement
    
    oneBranchMatchFlag_ = checkOneConditionMatch(root)
    getOneBranchMatchFlag = oneBranchMatchFlag_
End Function

'�˲�һ����֧���Equals���Ƿ���ƥ�����
Private Function checkOneBranchItemMatch(ByRef rootNode As Variant) As Boolean
    checkOneBranchItemMatch = False
    
    Dim attributeName As String
    
    Dim controlAttrValue As CControlAttrValue
    Dim xmlInfo As String
    
    Dim equalsNode As Variant
    For Each equalsNode In rootNode.SelectNodes("Equals")
        attributeName = equalsNode.getAttribute("attribute") '�õ����ز�������
        Set controlAttrValue = controlAttrValueManager_.getControlAttrValue(attributeName) '�õ����ز���ֵ���͵���
        checkOneBranchItemMatch = controlAttrValue.checkAttrValidation(equalsNode) '����ǰ���ز����ڵ㴫�����ز����࣬����ƥ���У��
        
        If checkOneBranchItemMatch = True Then Exit Function
    Next equalsNode
End Function

'�˲���Condition�Ƿ�ͬʱƥ�䣬һ�������ֻ��һ��Condition
Private Function checkAllConditionMatch(ByRef rootNode As Variant) As Boolean
    Dim matchFlag As Boolean, conditionMatchFlag As Boolean
    matchFlag = True
    Dim conditionNode As Variant
    For Each conditionNode In rootNode.SelectNodes("Condition")
        conditionMatchFlag = checkOneConditionMatch(conditionNode)
        matchFlag = matchFlag And conditionMatchFlag
    Next conditionNode
    checkAllConditionMatch = matchFlag
End Function

'�˲�һ����֧�Ƿ�ƥ��
Private Function checkOneBranchMatch(ByRef rootNode As Variant) As Boolean
    Dim enumItemNode As Variant
    Dim currentBrachEnumItemMatchFlag As Boolean, allConditionMatchFlag As Boolean
    currentBrachEnumItemMatchFlag = checkOneBranchItemMatch(rootNode) '��ǰ��֧���ӽڵ�equalsֵ�Ƿ�ƥ��
    allConditionMatchFlag = checkAllConditionMatch(rootNode) '���е�Condition��֧�Ƿ�ͬʱ����
    checkOneBranchMatch = currentBrachEnumItemMatchFlag And allConditionMatchFlag '���������㣬��ǰ��֧ƥ��ɹ�
End Function

'���һ��Condition�Ƿ�ƥ��
Private Function checkOneConditionMatch(ByRef rootNode As Variant) As Boolean
    checkOneConditionMatch = False
    Dim branchMatchFlag As Boolean
    Dim branchNode As Variant
    For Each branchNode In rootNode.SelectNodes("Branch") 'rootNode.ChildNodes
        branchMatchFlag = checkOneBranchMatch(branchNode)
        If branchMatchFlag = True Then
            checkOneConditionMatch = True
            matchBranchXml_ = branchNode.XML
            Set matchBranchNode_ = branchNode '�õ�ƥ���֧��Branch�ڵ�
            Exit Function
        End If
    Next branchNode
End Function

Private Sub Class_Initialize()
    oneBranchMatchFlag_ = False
End Sub
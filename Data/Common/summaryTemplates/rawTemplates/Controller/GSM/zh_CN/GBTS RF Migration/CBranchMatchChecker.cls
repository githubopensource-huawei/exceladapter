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

Private controlAttrValueManager_ As CControlAttrValueManager '所有主控参数数据的管理类
Private controlInfo_ As String 'XML的信息
Private oneBranchMatchFlag_ As Boolean '是否有一个分支匹配的标志
Private matchBranchXml_ As String '匹配的分支xml信息，用于对当前参数值的校验
Private matchBranchNode_ As Variant '匹配的Branch节点，是否能用后面再看

Public Sub init(ByRef controlAttrValueManager As CControlAttrValueManager, ByRef controlInfo As String)
    Set controlAttrValueManager_ = controlAttrValueManager
    controlInfo_ = controlInfo
End Sub

'得到匹配分支节点
Public Function getMatchBranchNode() As Variant
    Set getMatchBranchNode = matchBranchNode_
End Function

'看匹配分支的主控参数值是否为空，要求oneBranchMatchFlag_为True
Public Function getMatchBranchAttrEmptyFlag() As Boolean
    getMatchBranchAttrEmptyFlag = False
    If oneBranchMatchFlag_ = False Then
        getMatchBranchAttrEmptyFlag = True
        Exit Function '如果匹配分支为false，说明没有匹配，假定某个主控参数为空，不处理
    End If
    Dim controlAttrValue As CControlAttrValue
    Dim equalsNode As Variant
    Dim attributeName As String
    
    For Each equalsNode In matchBranchNode_.GetElementsByTagName("Equals")
        attributeName = equalsNode.getAttribute("attribute") '得到主控参数名称
        Set controlAttrValue = controlAttrValueManager_.getControlAttrValue(attributeName) '得到主控参数值类型的类
        If controlAttrValue.valueEmptyFlag = True Then '看匹配分支的每个主控参数值是否为空
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

'核查一个分支里的Equals里是否有匹配对象
Private Function checkOneBranchItemMatch(ByRef rootNode As Variant) As Boolean
    checkOneBranchItemMatch = False
    
    Dim attributeName As String
    
    Dim controlAttrValue As CControlAttrValue
    Dim xmlInfo As String
    
    Dim equalsNode As Variant
    For Each equalsNode In rootNode.SelectNodes("Equals")
        attributeName = equalsNode.getAttribute("attribute") '得到主控参数名称
        Set controlAttrValue = controlAttrValueManager_.getControlAttrValue(attributeName) '得到主控参数值类型的类
        checkOneBranchItemMatch = controlAttrValue.checkAttrValidation(equalsNode) '将当前主控参数节点传给主控参数类，进行匹配的校验
        
        If checkOneBranchItemMatch = True Then Exit Function
    Next equalsNode
End Function

'核查多个Condition是否同时匹配，一般情况下只有一个Condition
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

'核查一个分支是否匹配
Private Function checkOneBranchMatch(ByRef rootNode As Variant) As Boolean
    Dim enumItemNode As Variant
    Dim currentBrachEnumItemMatchFlag As Boolean, allConditionMatchFlag As Boolean
    currentBrachEnumItemMatchFlag = checkOneBranchItemMatch(rootNode) '当前分支的子节点equals值是否匹配
    allConditionMatchFlag = checkAllConditionMatch(rootNode) '所有的Condition分支是否同时满足
    checkOneBranchMatch = currentBrachEnumItemMatchFlag And allConditionMatchFlag '两个都满足，则当前分支匹配成功
End Function

'检查一个Condition是否匹配
Private Function checkOneConditionMatch(ByRef rootNode As Variant) As Boolean
    checkOneConditionMatch = False
    Dim branchMatchFlag As Boolean
    Dim branchNode As Variant
    For Each branchNode In rootNode.SelectNodes("Branch") 'rootNode.ChildNodes
        branchMatchFlag = checkOneBranchMatch(branchNode)
        If branchMatchFlag = True Then
            checkOneConditionMatch = True
            matchBranchXml_ = branchNode.XML
            Set matchBranchNode_ = branchNode '得到匹配分支的Branch节点
            Exit Function
        End If
    Next branchNode
End Function

Private Sub Class_Initialize()
    oneBranchMatchFlag_ = False
End Sub


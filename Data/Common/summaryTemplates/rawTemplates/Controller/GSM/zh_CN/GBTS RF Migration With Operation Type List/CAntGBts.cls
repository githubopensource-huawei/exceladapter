VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAntGBts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private p_cellId As String
Private p_brds As String
Private p_portNo As String
Private p_antNo As String
Private p_rowCount As Long
Private p_range As String
Private p_trxNum As Long
Private p_freqNos As String
Private p_antGrpId As String
Private Const BRACKET_RIGHT = "]"
Private Const BRACKET_LEFT = "["
Private Const COMMA = ","
Dim antenneCollection As Collection

    

Property Let cellId(str As String)
        p_cellId = str
End Property

Property Let trxNum(str As Long)
        p_trxNum = str
End Property

Property Let brds(str As String)
        p_brds = expandStr(str, p_trxNum)
End Property

Property Let freqNos(str As String)
       p_freqNos = str
End Property

Property Let antNo(str As String)
        p_antNo = expandStr(str, p_trxNum)
End Property
Property Let antGrpId(str As String)
        p_antGrpId = expandStr(str, p_trxNum)
End Property

Property Let ranges(str As String)
        p_range = str
End Property

Property Let portNos(str As String)
        p_portNo = expandStr(str, p_trxNum)
End Property

Property Get cellId() As String
        cellId = p_cellId
End Property

Property Get trxNum() As Long
        trxNum = p_trxNum
End Property

Property Get brds() As String
       brds = p_brds
End Property

Property Get ranges() As String
       ranges = p_range
End Property

Property Get portNos() As String
       portNos = p_portNo
End Property

Property Get freqNos() As String
       freqNos = p_freqNos
End Property

Property Get antNo() As String
       antNo = p_antNo
End Property
Property Get antGrpId() As String
       antGrpId = p_antGrpId
End Property

Property Get rowCount() As String
        rowCount = p_rowCount
End Property

Private Function getSplitStr(srcStr As String)
    Dim splitStr As String
    
    splitStr = COMMA
    If Not (InStr(srcStr, BRACKET_RIGHT) = 0) Then
        splitStr = BRACKET_RIGHT
    End If
    
    getSplitStr = splitStr
End Function

Private Function expandStr(srcStr As String, cnt As Long)
        Dim dstStr As String
        Dim strArray() As String
        Dim splitStr As String
        splitStr = getSplitStr(srcStr)
        
        Dim conStr As String
        Dim fmtStr As String
        fmtStr = srcStr
        If COMMA = splitStr Then
            conStr = COMMA
        Else
            conStr = BRACKET_RIGHT
            If Right(srcStr, 1) = BRACKET_RIGHT Then
                fmtStr = Left(srcStr, Len(srcStr) - 1)
            End If
        End If
  
        strArray = Split(fmtStr, splitStr)

      
        Dim iNo As Long
        
        If UBound(strArray) = 0 Then
            '1扩展为1,1,1,1,1连接符号为逗号 [1,2扩展为[1,2][1,2][1,2][1,2][1,2连接符号为右括号
            iNo = 0
            
            While (iNo < cnt)
                If (0 = iNo) Then
                    dstStr = strArray(iNo)
                Else
                    dstStr = dstStr + conStr + strArray(0)
                End If
            
                iNo = iNo + 1
            Wend
            
            dstStr = dstStr + conStr
        Else
            '1,2 扩展为1,2,,, [1,2][3,4]扩展为[1,2][3,4]]]]
            dstStr = srcStr
            For iNo = UBound(strArray) + 1 To (cnt - 1)
                dstStr = dstStr + conStr
            Next
        End If
        
        expandStr = dstStr
End Function

Private Function cutTail(str As String) As String
        Dim strArr() As String
        Dim finaStr As String
        
        Dim splitStr As String
        splitStr = getSplitStr(str)
        Dim conStr As String
        
        If COMMA = splitStr Then
            conStr = COMMA
        Else
            conStr = BRACKET_RIGHT
        End If
        
        strArr = Split(str, splitStr)
        finaStr = ""
        Dim iNo As Long
        iNo = 0
        While (iNo < p_trxNum And iNo <= UBound(strArr))
            If (0 = iNo) Then
                finaStr = strArr(iNo)
            Else
                finaStr = finaStr + conStr + strArr(iNo)
            End If
            
            iNo = iNo + 1
        Wend
        
        
        If COMMA = splitStr Then
            cutTail = finaStr
        Else
            cutTail = finaStr + conStr
        End If
End Function

Private Function getCurValueSafty(strArr() As String, index As Long) As String

    If index <= UBound(strArr) Then
        Dim loc As Long
        loc = InStr(strArr(index), BRACKET_LEFT)
        If loc = 0 Then
            getCurValueSafty = Replace(strArr(index), BRACKET_LEFT, "")
        Else
            getCurValueSafty = Right(strArr(index), Len(strArr(index)) - loc)
        End If
    Else
        getCurValueSafty = ""
    End If
End Function

Private Function getStrToArrSize(srcStr As String, delimiter As String)
    Dim strArr() As String
    
    strArr = Split(srcStr, delimiter)
    getStrToArrSize = UBound(strArr)
End Function

Private Function getMaxStrToArrUBound(strArr() As String) As Long
    Dim index As Long
    Dim maxSize As Long
    
    maxSize = 0
        
    For index = LBound(strArr) To UBound(strArr)
        maxSize = max(maxSize, getStrToArrSize(strArr(index), COMMA))
    Next
    
    getMaxStrToArrUBound = maxSize
End Function

Private Function autoSplit(srcStr As String) As String()
    Dim deliStr As String
    deliStr = getSplitStr(srcStr)
    
    Dim fmtStr As String
    fmtStr = srcStr
    If BRACKET_RIGHT = deliStr Then
        fmtStr = Left(srcStr, Len(srcStr) - 1)
    End If
    
    autoSplit = Split(srcStr, deliStr)
End Function


Function getAntenneCollection() As Collection
    Set getAntenneCollection = antenneCollection
End Function

Public Function commitData() As Boolean

    Dim brdArray() As String
    Dim frqArray() As String
    Dim portNoArray() As String
    Dim antArray() As String
    Dim antGrpIdArray() As String
    Dim index As Long
    Dim innerIndex As Long
    Set antenneCollection = New Collection
    
    brdArray = autoSplit(p_brds)
    frqArray = autoSplit(p_freqNos)
    portNoArray = autoSplit(p_portNo)
    antArray = autoSplit(antNo)
    antGrpIdArray = autoSplit(antGrpId)
    
    commitData = True
    
    If (0 = p_trxNum) Or (-1 = UBound(frqArray)) Then
        Call MsgBox(getResByKey("TrxNumCannotBeZeroOrEmpty") & ":" & p_cellId, vbInformation, getResByKey("Warning"))
        commitData = False
        Exit Function
    End If
    
    If p_trxNum > 1 + UBound(frqArray) Then
        Call MsgBox(getResByKey("freqNumShouldBeBiggerThanTrxNum") & ":" & p_cellId, vbInformation, getResByKey("Error"))
        commitData = False
        Exit Function
    End If
    
'    If UBound(frqArray) > UBound(brdArray) Or UBound(frqArray) > UBound(portNoArray) Or UBound(frqArray) > UBound(antArray) Then
'        Call MsgBox(getResByKey("DebugArrayCrash") & ":" & p_cellId, vbInformation, getResByKey("Warning"))
'    End If
    
    'process per freq / per cell
    For index = LBound(frqArray) To UBound(frqArray)
                Call expandBinds(antenneCollection, _
                    p_cellId, _
                    getCurValueSafty(frqArray, index), _
                    getCurValueSafty(brdArray, index), _
                    getCurValueSafty(portNoArray, index), _
                    getCurValueSafty(antArray, index), _
                    getCurValueSafty(antGrpIdArray, index))
                    
    Next
    
End Function

Private Sub expandBinds(ByRef antCollection As Collection, sCellId As String, sFreq As String, sBrd As String, sPortNo As String, sAntNo As String, sAntGrpId As String)
    Dim strArr(4) As String
    strArr(0) = sBrd
    strArr(1) = sPortNo
    strArr(2) = sAntNo
    strArr(3) = sAntGrpId
    
    Dim maxUbound As Long
    maxUbound = getMaxStrToArrUBound(strArr)
    
    sBrd = expandStr(sBrd, maxUbound + 1)
    sPortNo = expandStr(sPortNo, maxUbound + 1)
    sAntNo = expandStr(sAntNo, maxUbound + 1)
    sAntGrpId = expandStr(sAntGrpId, maxUbound + 1)
    
    
    Dim brdArray() As String
    Dim portNoArray() As String
    Dim antArray() As String
    Dim antGrpIdArray() As String
    brdArray = Split(sBrd, COMMA)
    portNoArray = Split(sPortNo, COMMA)
    antArray = Split(sAntNo, COMMA)
    antGrpIdArray = Split(sAntGrpId, COMMA)
    
    Dim index As Long
    For index = LBound(brdArray) To maxUbound
        Dim antenne(5) As String
        antenne(0) = sCellId
        antenne(1) = sFreq
        antenne(2) = getCurValueSafty(brdArray, index)
        antenne(3) = getCurValueSafty(portNoArray, index)
        antenne(4) = getCurValueSafty(antArray, index)
        antenne(5) = getCurValueSafty(antGrpIdArray, index)
        Call antCollection.Add(antenne)
    Next
End Sub







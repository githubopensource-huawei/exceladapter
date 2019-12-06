VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGBtsTrxBind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_cellId As String
Private p_freqNo As String
Private p_brdNo As String
Private p_portNo As String
Private p_antNo As String
Private p_antGrpId As String

    

Property Let cellId(str As String)
        p_cellId = str
End Property



Property Get cellId() As String
        cellId = p_cellId
End Property



    

Property Let freqNo(str As String)
        p_freqNo = str
End Property



Property Get freqNo() As String
        freqNo = p_freqNo
End Property




    

Property Let brdNo(str As String)
        p_brdNo = str
End Property



Property Get brdNo() As String
        brdNo = p_brdNo
End Property



    

Property Let portNo(str As String)
        p_portNo = str
End Property



Property Get portNo() As String
        portNo = p_portNo
End Property



    

Property Let antNo(str As String)
        p_antNo = str
End Property




Property Get antNo() As String
        antNo = p_antNo
End Property
Property Let antGrpId(str As String)
        p_antGrpId = str
End Property
Property Get antGrpId() As String
        antGrpId = p_antGrpId
End Property




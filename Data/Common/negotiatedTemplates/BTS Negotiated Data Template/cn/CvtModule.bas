Attribute VB_Name = "CvtModule"
  Const Col_CellName As Integer = 2
  Const Col_BrdType As Integer = 3
  Const Col_PassNo As Integer = 4
  Const Col_CN As Integer = 5
  Const Col_SRN As Integer = 6
  Const Col_SN As Integer = 7
  Const Col_BCCHFREQ As Integer = 8
  Const Col_Non_BCCHFREQ As Integer = 9
  Const Row_Begin As Integer = 6
  Const Row_Begin_Target As Integer = 6
  Const BCCH As String = "YES"
  Const NON_BCCH As String = "NO"
  
  Private GCellSht As Worksheet
  

Public Sub CvtTRXMAPtoGTRX()
  Dim sCellName As String
  Dim sPassNo As String
  Dim sCN As String
  Dim sSRN As String
  Dim sSN As String
  Dim sBCCH As String
  Dim sNonBCCH As String
  Dim sFreq As String
  
  
  Dim iSrcRow As Integer
  Dim iTgtRow As Integer
  
  Dim SrcSht As Worksheet
  Dim TgtSht As Worksheet
  
  Set SrcSht = ThisWorkbook.Worksheets("Frequency Tool")
  Set TgtSht = ThisWorkbook.Worksheets("GTRX")
  Set GCellSht = ThisWorkbook.Worksheets("GCELL")
  
  iSrcRow = Row_Begin
  iTgtRow = Row_Begin_Target
  
  Do
    '校验小区是否存在
    sCellName = SrcSht.Cells(iSrcRow, Col_CellName)
    If sCellName = "" Then
      Exit Do
    End If
    
    Application.StatusBar = "Processing Line : " + str(iSrcRow)
    
    '处理主B频点
    sBCCH = SrcSht.Cells(iSrcRow, Col_BCCHFREQ)
    If Not sBCCH = "" Then
      Call WriteLineToTarget(SrcSht, iSrcRow, TgtSht, iTgtRow, BCCH, sBCCH)
      iTgtRow = iTgtRow + 1
    End If
    
    '处理非主B频点
    sNonBCCH = Replace(SrcSht.Cells(iSrcRow, Col_Non_BCCHFREQ), "，", ",")
    If Not sNonBCCH = "" Then
      arrFreqs = Split(sNonBCCH, ",")
      For i = 0 To UBound(arrFreqs)
        sFreq = arrFreqs(i)
        Call WriteLineToTarget(SrcSht, iSrcRow, TgtSht, iTgtRow, NON_BCCH, sFreq)
        iTgtRow = iTgtRow + 1
      Next
    End If
  
    iSrcRow = iSrcRow + 1
  Loop
  
  If iSrcRow = Row_Begin Then
    MsgBox "The Cell Name should be assigned."
  Else
    MsgBox "Converting GTRX finished."
  End If
  
  Set GCellSht = Nothing
  Application.StatusBar = ""
End Sub

'GTRX sheet
Public Sub WriteLineToTarget(SrcSht As Worksheet, iSrcRow As Integer, TgtSht As Worksheet, iTgtRow As Integer, sBCCH As String, sFreq As String)
  TgtSht.Cells(iTgtRow, 3) = SrcSht.Cells(iSrcRow, Col_CellName)
  TgtSht.Cells(iTgtRow, 4) = sFreq
  TgtSht.Cells(iTgtRow, 5) = sBCCH
  TgtSht.Cells(iTgtRow, 7) = SrcSht.Cells(iSrcRow, Col_BrdType)
  TgtSht.Cells(iTgtRow, 9) = SrcSht.Cells(iSrcRow, Col_PassNo)
  TgtSht.Cells(iTgtRow, 10) = SrcSht.Cells(iSrcRow, Col_CN)
  TgtSht.Cells(iTgtRow, 11) = SrcSht.Cells(iSrcRow, Col_SRN)
  TgtSht.Cells(iTgtRow, 12) = SrcSht.Cells(iSrcRow, Col_SN)
End Sub

'取得小区名称
Function GetCellIdByCellName(CellName As String) As String
  Dim sTmpCellName As String
   
  For i = 6 To 65535
    sTmpCellName = GCellSht.Cells(i, 4)
    If sTmpCellName = "" Then
        Exit Function
    End If
  
    Application.StatusBar = "Searching Cell: " + CellName + "   Line: " + str(i)
    If UCase(Trim(sTmpCellName)) = UCase(Trim(CellName)) Then
        GetCellIdByCellName = GCellSht.Cells(i, 3)
        Exit Function
    End If
    
  Next
End Function






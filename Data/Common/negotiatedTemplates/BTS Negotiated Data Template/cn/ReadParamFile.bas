Attribute VB_Name = "ReadParamFile"

Public HWTmpltDoc As Workbook
Public VFTmpltDoc As Workbook

'Public Function GetParameterFile() As Workbook
'  Dim PeerFile As Workbook
'  Dim LocalFile As Workbook
'
'  Set LocalFile = ActiveWorkbook
'
'  FileName = Trim(LocalFile.Worksheets("TableDef").Cells(1, 6).value)
'
'  Application.StatusBar = "Open file : " + FileName
'  If Not FileName = "" Then
'    Application.DisplayAlerts = False
'    Workbooks.Open FileName:=FileName, ReadOnly:=False
'    Application.DisplayAlerts = True
'    Set PeerFile = ActiveWorkbook
'    ThisWorkbook.Activate
'
'    Set GetParameterFile = PeerFile
'  Else
'    MsgBox "文件未指定"
'  End If
'End Function
'
'Public Sub LoadParams()
'    Const FieldName As Integer = 3
'    Const CfgId As Integer = 1
'    Const ObjName As Integer = 2
'
'    Dim sFieldName As String
'    Dim sId As String
'    Dim sObjectName As String
'
'    Dim PeerFile As Workbook
'
'    Dim iSrcRow As Integer
'
'    Dim SrcSht As Worksheet
'    Dim ParamSht As Worksheet
'
'    Application.StatusBar = "Beginning ."
'    Set SrcSht = ThisWorkbook.Worksheets("TableDef")
'
'    iSrcRow = 15
'
'    Set ParamSht = GetParameterFile.Worksheets("Parameter List")
'
'    Do
'      '校验字段是否存在
'      sFieldName = SrcSht.Cells(iSrcRow, FieldName)
'      If sFieldName = "" Then
'        Exit Do
'      End If
'
'      If Not ("" = Trim(SrcSht.Cells(iSrcRow, CfgId))) Then
'        sId = Trim(SrcSht.Cells(iSrcRow, CfgId))
'        sObjectName = Trim(SrcSht.Cells(iSrcRow, ObjName))
'      End If
'
'      Application.StatusBar = "Processing : " + sObjectName + "-" + sFieldName
'
'      Call CopyParaDefinitions(SrcSht, ParamSht, iSrcRow, sObjectName, sFieldName)
'
'      iSrcRow = iSrcRow + 1
'    Loop
'
'    MsgBox "完成"
'    Application.StatusBar = ""
'End Sub
'
'Sub CopyParaDefinitions(SrcSht As Worksheet, ParamSht As Worksheet, iCurRow As Integer, sObj As String, sField As String)
'    Const Detail_Name As Integer = 5
'    Const Dsp_Name As Integer = 2
'    Const BTS_Detail_Name As Integer = 17
'    Const BTS_Dsp_Name As Integer = 16
'
'
'    SrcSht.Cells(iCurRow, 18) = iCurRow
'    For iRow = 2 To 65535
'      If ParamSht.Cells(iRow, 1) = "" Then
'        Exit Sub
'      End If
'
'      Application.StatusBar = "Processing : " + sObj + "-" + sField + "  Line: " + Str(iRow)
'
'      If sField = UCase(ParamSht.Cells(iRow, 1)) Then
'        SrcSht.Cells(iCurRow, BTS_Detail_Name) = ParamSht.Cells(iRow, Detail_Name).value
'        SrcSht.Cells(iCurRow, BTS_Dsp_Name) = ParamSht.Cells(iRow, Dsp_Name).value
'        Exit Sub
'      End If
'    Next
'End Sub



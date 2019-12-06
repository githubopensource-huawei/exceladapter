Attribute VB_Name = "mod_public"

'*******************************************************************'
'********************* Version infomation start ********************'
'*******************************************************************'

' Embed Summary's Collection Table Release Notes

' 2011-04-18 17:59    Version: 1.4
' 1. ������������

' 2011-04-12 19:59    Version: 1.3
' 1. ��������֮ǰ���Ǻ�

' 2011-03-07 18:59    Version: 1.2
' 1. Ƶ������ֲ�� R11 Summary ������
' 2. �����Զ���������ļ��Ĺ���

' 2010-12-30 15:59    Version: 1.1
' 1. ѡ�������ļ����Զ���ӵ������б�
' 2. ��дƵ����

' 2010-11-30 15:59    Version: 1.0
' 1. ʵ��Ƶ����



'*******************************************************************'
'********************** Version infomation end *********************'
'*******************************************************************'


Global Const gShtName_bts = "BTS Transport Layer"
Global Const gShtName_cell = "Cell Basic Info"
Global gColName_btsName As String
Global gColName_btsType As String
Global gColName_cellName As String
Global gColName_cellName2 As String
Global gColName_cellType As String
Global gColName_bcch As String
Global gColName_noBcch As String
Global gColName_fc As String
Global gColName_btsTpltName As String
Global Const constRecordRow = 2

'************Template ������Ӣ��***************
Global gCaption_Label1 As String
Global gCaption_Label2 As String
Global gCaption_OptionButton1 As String
Global gCaption_OptionButton2 As String
Global gCaption_SubmitAdd As String
Global gCaption_CancelButton As String
Global gCaption_OKButton As String
Global gCaption_SubmitDelete As String
Global gCaption_TemplateForm As String

Global gCaption_CustomizeTemplate As String

Global gCaptionCreateBTS As String
Global gCaptionRpsTDMInBSC As String
Global gCaptionRpsBetweenBSC As String
Global gCaptionAll As String
Global gCaptionSceneFrame As String


 Global gMsg_AddEmpty As String
 Global gMsg_AddExistH As String
 Global gMsg_AddExistE As String
 Global gMsg_AddSuccH As String
 Global gMsg_AddSuccE As String
 
 Global gMsg_DelEmpty As String
 Global gMsg_DelExistH As String
 Global gMsg_DelExistE As String
 Global gMsg_DelSuccH As String
 Global gMsg_DelSuccE As String
 
 Global gMsg_OperWarning As String
 Global gMsg_OperInfo As String
 
'************Check frequency band ������Ӣ��***************
 Global gCaption_BandTitle As String
 Global gCaption_BandCheck As String
 Global gCaption_BandCancel As String
 Global gMsg_BandCheckPass As String
 Global gMsg_BandCheckError As String
 
   
Public Type Cell
    BTSName As String
    GSMCellName As String
    celltype As String
    BCCH As String
    FreqListCount As String
    
    CellFreqs() As String
End Type
Public gPath As String
Public FreqType_BTS3900_900M As String
Public FreqType_BTS3900_1800M As String
Public FreqType_DBS3900_900M As String
Public FreqType_DBS3900_1800M As String

Global Const InvalidColNameCol = 1
Global Const InvalidGroupNameCol = 2
Global Const InvalidShtNameCol = 3

'Public FreqTypeXml As New DOMDocument

Global gErrMsg As String
Global Const gLogFileName = "\error.log"

Global Const gSpaceCnt = 10         ' ����Ŀհ���(��)�������Ϊ 10
Global Const gRowMax = 60000        ' ֻ����ǰ 60000 �е�����
Global Const gColMax = 200          ' ֻ����ǰ 200  �е�����

Global Const gCreateBTS = "GSM_SUMMARY_CREATEBTS"
Global Const gRpsTDMInBSC = "GSM_BTS_REPARENT_TDM_INBSC"
Global Const gRpsIPInBSC = "GSM_BTS_REPARENT_IP_INBSC"
Global Const gRpsBetweenBSC = "GSM_BTS_REPARENT"

Global Const gMappingDefShtName = "MAPPING DEF"
Global Const gShtNameInvalidFields = "InvalidFields"
Global Const gShtNameSpecialFields = "SpecialFields"
Global Const gShtNameFuctionMocs = "FuctionMocs"
Global Const gRxuSpecShtName = "RXU Specification"


' Map ҳ����
Global Const gColName_srcShtName = "Sheet Name"
Global Const gColName_groupName = "Group Name"
Global Const gColName_srcColName = "Column Name"
Global Const gColName_dstShtName = "MOC Name"
Global Const gColName_dstColName = "Attribute Name"
Global gCurScene As String

' Sheet ҳ����
Public Type shtPage                 ' ��վ
    shtName As String               ' Sheetҳ����
    shtData() As String             ' Sheetҳ�����ݣ���ά����
End Type

Public Sub frmShow()
    frmChooseFreqType.cbo_bts_900.Clear
    frmChooseFreqType.cbo_bts_1800.Clear
    frmChooseFreqType.cbo_dbs_900.Clear
    frmChooseFreqType.cbo_dbs_1800.Clear
    
    frmChooseFreqType.cbo_bts_900.AddItem ("15M")
    frmChooseFreqType.cbo_bts_900.AddItem ("20M")
    frmChooseFreqType.cbo_bts_900.AddItem ("20.2M")
    
    frmChooseFreqType.cbo_bts_1800.AddItem ("15M")
    frmChooseFreqType.cbo_bts_1800.AddItem ("20M")
    frmChooseFreqType.cbo_bts_1800.AddItem ("20.2M")
    
    frmChooseFreqType.cbo_dbs_900.AddItem ("12.5M")
    frmChooseFreqType.cbo_dbs_900.AddItem ("15M")
    
    frmChooseFreqType.cbo_dbs_1800.AddItem ("12.5M")
    frmChooseFreqType.cbo_dbs_1800.AddItem ("15M")
    
    frmChooseFreqType.cbo_bts_900.ListIndex = 0
    frmChooseFreqType.cbo_bts_1800.ListIndex = 1
    frmChooseFreqType.cbo_dbs_900.ListIndex = 0
    frmChooseFreqType.cbo_dbs_1800.ListIndex = 0
        
    LoadXmlToFrm
    
    'init ThisWorkbook
    frmChooseFreqType.InitGUI
    frmChooseFreqType.Show
End Sub

' ��ȡ para lsit �ļ�
Private Sub readBandFile()
    Dim i, j  As Integer
    Dim readBuf As String
    Dim tmp() As String, substr() As String
    
    ' ��ȡ�����ļ�
    If Dir(ThisWorkbook.Path & "\FreqType.txt") = "" Then
        Exit Sub
    End If
    readTxtFile readBuf, ThisWorkbook.Path & "\FreqType.txt"
    readBuf = Replace(readBuf, vbCr, vbLf)
    
    ' ����ַ���
    tmp = Split(readBuf, vbLf, , vbTextCompare)
    For i = 0 To UBound(tmp)
    
        If 0 < Len(tmp(i)) Then     ' �� tmp(i) Ϊ��ʱ�� substr Ϊ��
            substr = Split(tmp(i), "=", , vbTextCompare)
            
            Select Case Trim(UCase(substr(0)))
                Case "BTS3900_900M"
                    FreqType_BTS3900_900M = Trim(substr(1))
                    
                Case "BTS3900_1800M"
                    FreqType_BTS3900_1800M = Trim(substr(1))
                
                Case "DBS3900_900M"
                    FreqType_DBS3900_900M = Trim(substr(1))
                
                Case "DBS3900_1800M"
                    FreqType_DBS3900_1800M = Trim(substr(1))
            End Select
        End If
    Next i
End Sub

'*************************************************
' ��ȡXML�����ļ�����Ϣ����¼���һ��ѡ���Ƶ������
'*************************************************
Private Function LoadXmlToFrm()
    Dim i As Integer
    Dim j As Integer
    Dim strPath As String
    Dim ParentNodeName As String
    Dim ChildNodeName As String

    readBandFile
    
    If FreqType_BTS3900_900M <> "" Then
        frmChooseFreqType.cbo_bts_900.Text = FreqType_BTS3900_900M
    End If
    If FreqType_BTS3900_1800M <> "" Then
        frmChooseFreqType.cbo_bts_1800.Text = FreqType_BTS3900_1800M
    End If
    If FreqType_DBS3900_900M <> "" Then
        frmChooseFreqType.cbo_dbs_900.Text = FreqType_DBS3900_900M
    End If
    If FreqType_DBS3900_1800M <> "" Then
        frmChooseFreqType.cbo_dbs_1800.Text = FreqType_DBS3900_1800M
    End If
End Function

'*************************************************
' ��ѡ���Ƶ����Ϣд��XML�ļ���
'*************************************************
Public Function WriteFreqTypeToXML()
    Dim xmlFile As String, msg As String

    xmlFile = ThisWorkbook.Path & "\FreqType.txt"
    
    msg = "BTS3900_900M=" & frmChooseFreqType.cbo_bts_900.Text & vbLf & _
          "BTS3900_1800M=" & frmChooseFreqType.cbo_bts_1800.Text & vbLf & _
          "DBS3900_900M=" & frmChooseFreqType.cbo_dbs_900.Text & vbLf & _
          "DBS3900_1800M=" & frmChooseFreqType.cbo_dbs_1800.Text
    writeTxtFile xmlFile, msg
End Function

'*************************************************
' �������е����ݽ�����������
'*************************************************
Public Function sortArrayAsInteger(ByRef sa() As String)
    Dim x, y, cnt, ia, ib As Integer
    Dim st As String
    
    cnt = UBound(sa)
    For x = 0 To cnt
        For y = cnt - 1 To x Step -1
            If "" = sa(y) Then
                ia = 0
            Else
                ia = CInt(sa(y))
            End If
            
            If "" = sa(y + 1) Then
                ib = 0
            Else
                ib = CInt(sa(y + 1))
            End If
            
            If ia > ib Then
                st = sa(y)
                sa(y) = sa(y + 1)
                sa(y + 1) = st
            End If
        Next y
    Next x
End Function

' д�ı��ļ�
Public Function writeTxtFile(ByRef fileName As String, ByRef msg As String)
    Dim fileNum As Integer
    On Error GoTo E
    fileNum = FreeFile()
    Open fileName For Output As #fileNum
        Print #fileNum, msg
    Close #fileNum
    Exit Function
E:
End Function

' д������־�ļ�
Public Function writeLogFile(ByVal fileName As String, ByVal msg As String)
    msg = "***************** error log start *****************" & vbLf & _
          "log time: " & Date$ & " " & Time$ & vbLf & vbLf & _
           msg & vbLf & _
          "****************** error log end ******************"
    writeTxtFile fileName, msg
End Function

' ��¼��־��Ϣ��д���ļ�
Public Function Log(Optional ByRef msg As String = "", Optional ByRef bClear As Boolean = False)
    Dim logFile As String
    logFile = ThisWorkbook.Path & gLogFileName
    
    If True = bClear Then
        gErrMsg = ""
        If "" <> Dir(logFile) Then
            Kill logFile
        End If
    ElseIf "" <> msg Then
        gErrMsg = gErrMsg & msg & vbLf
        writeLogFile logFile, gErrMsg
    End If
End Function

' �˳�����
Public Function Finish(bFinishProgram As Boolean)
    WriteFreqTypeToXML
    frmChooseFreqType.Hide
    If True = bFinishProgram Then
        End
    End If
End Function

' ��鵱ǰ�Ƿ��Ѿ����ִ���
Public Function CheckError()
    If "" <> gErrMsg Then
        MsgBox (gMsg_BandCheckError & " [ " & ThisWorkbook.Path & gLogFileName & " ] ")
        Finish True
    End If
End Function

' ��ȡ Sheet ҳ�����ݵ� sp ����
Public Function readSheetData(ByRef dstSP As shtPage, ByRef srcSht As Worksheet, Optional ByVal delBlankRow As Boolean = True)
    Dim row, col, srcRow, dstRow, rowCnt As Long, colCnt As Integer, dataRow As Long
    Dim allBuf() As String, delBuf() As String, tmp() As String, str As String
    
    ' ��ȡ������
    FindLocateFromSht srcSht, gColName_btsName, dataRow, colCnt
    'dataRow = getDataRowStart(srcSht.Name)
    GetShtRange srcSht, rowCnt, colCnt
    ReDim allBuf(rowCnt, colCnt), delBuf(rowCnt, colCnt)
    
    ' ��������
    dstRow = 0
    For srcRow = 1 To rowCnt ' ��ȡ���ݣ��� 1 �� 1 �п�ʼ
        str = ""
        dstRow = dstRow + 1
        For col = 1 To colCnt
            allBuf(srcRow, col) = Trim(srcSht.Cells(srcRow, col))
            delBuf(dstRow, col) = allBuf(srcRow, col)
            str = str & allBuf(srcRow, col)
        Next col
        
        ' ɾ������
        If (True = delBlankRow) And ("" = str) And (dstRow >= dataRow) Then
            dstRow = dstRow - 1
        End If
    Next srcRow
    
    ' ��������
    dstSP.shtName = srcSht.name
    If (True = delBlankRow) And (dstRow < srcRow) Then
        ReDim tmp(dstRow, colCnt)
        For row = 1 To dstRow       ' �����޿��е�����
            For col = 1 To colCnt
                tmp(row, col) = delBuf(row, col)
            Next col
        Next row
        dstSP.shtData = tmp         ' �����޿��е�����
    Else
        dstSP.shtData = allBuf      ' �����п��е�����
    End If
End Function

' ѡ����򿪵� excel�ļ���ֻ֧�� xls �� xlsx �ĺ�׺��
Public Function ChooseFile(ByRef fileName() As String)
    Dim i As Integer
    Dim FD As FileDialog
    Dim FDFilter As FileDialogFilter
    Dim FDFilters As FileDialogFilters
    Dim sFileNameIncorrect As String
    ReDim fileName(0)
    
    Application.ScreenUpdating = False  '�ر���Ļ���¿ɼӿ���ִ���ٶ�
    Set FD = Application.FileDialog(msoFileDialogOpen)
    Set FDFilters = FD.Filters
    FDFilters.Clear
    Set FDFilter = FDFilters.Add("Template Files", "*.xls")     '����ѡ�� xls ���͵��ļ�
    Set FDFilter = FDFilters.Add("Template Files", "*.xlsx")    '�Լ� xlsx ���͵��ļ�
    Set FDFilter = FDFilters.Add("Template Files", "*.*")       '�Լ� ���� ���͵��ļ�
    
    With FD
        .AllowMultiSelect = True
        .Show
        If (.SelectedItems.count = 0) Then
            Exit Function
        Else
            ReDim fileName(.SelectedItems.count) As String
            For i = 1 To .SelectedItems.count
                fileName(i) = .SelectedItems(i)
            Next i
        End If
    End With
    Set FD = Nothing
End Function

' �� Sheet �в���ָ���е�λ��
Public Function FindLocateFromSht(ByRef sht As Worksheet, ByVal sValue As String, ByRef retRow As Long, ByRef retCol As Integer)
    Dim row, rowCnt As Long, col, colCnt As Integer
    Dim rowBak, colBak As Integer
    Dim str As String
    
    retRow = 0
    retCol = 0
    GetShtRange sht, rowCnt, colCnt
    sValue = Replace(UCase(sValue), " ", "", , , vbTextCompare)
    
    For row = 1 To rowCnt
        For col = 1 To colCnt
            str = sht.Cells(row, col)
            If str = sValue Then
                retRow = row
                retCol = col
                Exit For
            End If
        Next col
    Next row
    
    If (0 = retRow) Then
        retRow = rowBak
        retCol = colBak
    End If
End Function

' ���� BTS Template Name ��λ��
Public Function FindLocateFromSP(ByRef sp As shtPage, ByVal sValue As String, ByRef retRow As Long, ByRef retCol As Integer)
    Dim row, rowCnt As Long, col, colCnt As Integer
    Dim rowBak, colBak As Integer
    Dim str As String
    retRow = 0
    retCol = 0
    sValue = Replace(UCase(sValue), " ", "", , , vbTextCompare)
    
    For row = 1 To UBound(sp.shtData, 1)
        For col = 1 To UBound(sp.shtData, 2)
            str = Replace(UCase(sp.shtData(row, col)), " ", "", , , vbTextCompare)
            If str = sValue Then
                retRow = row
                retCol = col
                Exit For
            ElseIf 0 < InStr(1, str, sValue) Then
                rowBak = row
                colBak = col
                Exit For
            End If
        Next col
    Next row
    
    If (0 = retRow) Then
        retRow = rowBak
        retCol = colBak
    End If
End Function

' ɾ������ǰ���Ǻ�
Public Function delTitleAsterisk(ByRef sp As shtPage, ByVal tRow As Integer)
    Dim row, col As Integer, str As String
    
    For col = 1 To UBound(sp.shtData, 2)
        str = Trim(sp.shtData(tRow, col))
        If 1 < Len(str) Then
            If "*" = Left(str, 1) Then
                sp.shtData(tRow, col) = Trim(Right(str, Len(str) - 1))      ' ɾ���ؼ�����֮ǰ�� *
            End If
        End If
    Next col
End Function

' ����ָ�� Sheet ҳ�����ݷ�Χ
Public Function GetShtRange(ByRef sht As Worksheet, ByRef rowCnt As Long, ByRef colCnt As Integer)
    colCnt = getColMax(sht)
    rowCnt = getRowMax(sht, colCnt)
End Function

' �������� Sheet ҳ��ǰ���������
Public Function getRowMax(ByRef sht As Worksheet, ByVal colCnt As Integer) As Long
    Dim row As Long, col As Integer, spcCnt As Integer
    Dim bFlag As Boolean
    
    spcCnt = 0
    For row = 1 To gRowMax + gSpaceCnt
        bFlag = False
        For col = 1 To colCnt                   ' �����кŵķ�Χ
            If "" <> Trim(sht.Cells(row, col)) Then
                bFlag = True
                spcCnt = 0
                Exit For
            End If
        Next col
        
        If False = bFlag Then                   ' ��ĳ�е������ж�Ϊ��ʱ������Ϊ��
            spcCnt = spcCnt + 1
            If gSpaceCnt <= spcCnt Then
                Exit For
            End If
        End If
    Next row
    
    row = row - spcCnt      ' ��ȡ�кţ���Խ��ʱ�����⴦��
    If 0 >= row Then
        row = 1
    End If
    
    getRowMax = row
End Function

' �������� Sheet ҳ��ǰ���������
Public Function getColMax(ByRef sht As Worksheet) As Integer
    Dim col As Integer
    Dim row As Integer, spcCnt As Integer
    Dim bFlag As Boolean
    
    spcCnt = 0
    For col = 1 To gColMax + gSpaceCnt      ' ������������������� 200 ��
        bFlag = False
        For row = 1 To gSpaceCnt            ' ��ǰ 10 �е���������������
            If "" <> Trim(sht.Cells(row, col)) Then
                bFlag = True
                spcCnt = 0
                Exit For
            End If
        Next row
        
        If False = bFlag Then
            spcCnt = spcCnt + 1             ' ��ĳһ�е�ǰ 10 ������ȫ��Ϊ��ʱ����Ϊ����Ϊ��
            If gSpaceCnt <= spcCnt Then
                Exit For
            End If
        End If
    Next col
    
    col = col - spcCnt      ' ��ȡ�кţ���Խ��ʱ�����⴦��
    If 0 >= col Then
        col = 1
    End If
    
    getColMax = col
End Function

' ��ǿ���ַ��滻��findCharList �е�ÿ���ַ������滻Ϊ replac�����Ҿ����������ظ��� replac
' �� findCharList Ϊ��ʱ��ʹ��Ĭ�ϵĲ����ַ����� ",��;��:��[]����"
' delBlank Ϊ true ʱɾ�����пո�
Public Function replacExt(ByVal srcStr As String, ByVal findCharList As String, ByVal replac As String, Optional ByVal delBlank As Boolean = True) As String
    Dim i As Integer
    
    ' ʹ��Ĭ�ϼ����
    If "" = findCharList Then
        findCharList = ",��;��:��[]����"
    End If
    
    ' ɾ�������ַ�
    For i = 1 To Len(findCharList)
        srcStr = Replace(srcStr, Mid(findCharList, i, 1), replac)
        srcStr = Replace(srcStr, replac & replac, replac)
    Next i
    
    ' ɾ����β�����
    srcStr = trimComma(srcStr, replac)
    
    ' ɾ���ո�
    If True = delBlank Then
        srcStr = Replace(srcStr, " ", "")
    End If
    
    replacExt = srcStr
End Function

' ɾ���ַ�����β����Ч�ַ���Ĭ��Ϊ����
Public Function trimComma(ByVal s As String, Optional ByVal ch As String = ",") As String
    Dim i, lch, pos As Integer, c As String
    
    lch = Len(ch)
    If ("" <> s) And ("" <> ch) Then
        ' ���Ҵ����
        pos = 0
        For i = 1 To Len(s)
            c = Mid(s, i, lch)
            If (ch = c) Then
                pos = i
            Else
                Exit For
            End If
        Next i
        s = Right(s, Len(s) - pos + lch - 1)
        
        ' ���Ҵ��ұ�
        pos = Len(s) + 1
        For i = Len(s) To 1 Step -1
            c = Mid(s, i, lch)
            If (ch = c) Then
                pos = i
            Else
                Exit For
            End If
        Next i
        s = Left(s, pos - 1)
    End If
    
    trimComma = s
End Function

' �� str ���� sArray() ���У�����Ѿ�����������
Public Function insertList(ByRef sArray() As String, ByVal str As String)
    Dim i, blank As Integer, bExist As Boolean
    
    bExist = False
    blank = 0
    For i = 1 To UBound(sArray)
        If "" = sArray(i) Then
            blank = i               ' ���ҿհ����鵥Ԫ
        End If
        If str = sArray(i) Then
            bExist = True           ' �Ѵ�������������
            Exit Function
        End If
    Next i
    
    If 1 > blank Then
        blank = i                   ' δ�ҵ��հ׵�Ԫ
    End If
    
    If False = bExist Then
        If UBound(sArray) < blank Then
            ReDim Preserve sArray(blank) As String  ' ��չ�����С
        End If
        sArray(blank) = str         ' ��������
    End If
End Function
Sub showCustomizeTemplateForm()
    Load CustomizeTemplateForm
    'init ThisWorkbook
    CustomizeTemplateForm.InitGUI
    CustomizeTemplateForm.Show
    
End Sub

Sub addTemplate()
    Load TemplateForm
    'init ThisWorkbook
    TemplateForm.InitGUI
    TemplateForm.Show
End Sub

' ��ʼ������
Public Function init(ByRef srcWB As Workbook)
       
    gColName_btsName = getResByKey("ColName_BTSName")
    gColName_btsType = getResByKey("ColumnName_BTSType")
    gColName_cellName = getResByKey("ColumnName_GCELLName")
    gColName_cellName2 = getResByKey("ColumnName_CELLName")
    gColName_cellType = getResByKey("ColumnName_CELLType")
    gColName_bcch = getResByKey("ColumnName_BCCH")
    gColName_noBcch = getResByKey("ColumnName_NonBCCH")
    gColName_fc = getResByKey("ColumnName_FreqClass")
    gColName_btsTpltName = getResByKey("ColumnName_BTSTemplateName")

    gCaption_TemplateForm = getResByKey("Caption_BTSTemplateForm")
    gCaption_Label1 = getResByKey("LabelName_BTSType")
    gCaption_Label2 = getResByKey("LabelName_TemplateName")
    gCaption_OptionButton1 = getResByKey("ButtonName_Add")
    gCaption_OptionButton2 = getResByKey("ButtonName_Delete")
    gCaption_SubmitAdd = getResByKey("ButtonName_Add")
    gCaption_SubmitDelete = getResByKey("ButtonName_Delete")
    gCaption_CancelButton = getResByKey("ButtonName_Cancle")
    gCaption_OKButton = getResByKey("ButtonName_OK")

    gMsg_AddEmpty = getResByKey("Message_TemplateNameEmpty")
    gMsg_AddExistH = getResByKey("Message_Template")
    gMsg_AddExistE = getResByKey("Message_AlreadExist")
    gMsg_AddSuccH = getResByKey("Message_AddTemplate")
    gMsg_AddSuccE = getResByKey("Message_Success")
    gMsg_DelEmpty = getResByKey("Message_InputTemplateName")
    gMsg_DelExistH = getResByKey("Message_Template")
    gMsg_DelExistE = getResByKey("Message_NotExist")
    gMsg_DelSuccH = getResByKey("Message_DeleteTemplate")
    gMsg_DelSuccE = getResByKey("Message_Success")

    gCaption_BandTitle = getResByKey("Caption_CheckFreqBand")
    gCaption_BandCheck = getResByKey("Caption_Check")
    gCaption_BandCancel = getResByKey("ButtonName_Cancle")

    gMsg_BandCheckPass = getResByKey("Message_CheckSuccess")
    gMsg_BandCheckError = getResByKey("Message_CheckFailed")
    gMsg_OperWarning = getResByKey("Message_Warning")
    gMsg_OperInfo = getResByKey("Message_Info")

    gCaption_CustomizeTemplate = getResByKey("CustomizeScene")
    gCaptionCreateBTS = getResByKey("Caption_CreateBTS")
    gCaptionRpsTDMInBSC = getResByKey("Caption_ReparentTDMInBSC")
    gCaptionRpsBetweenBSC = getResByKey("Caption_ReparentBTSBetweenBSC")
    gCaptionAll = getResByKey("Caption_AllScenes")
    gCaptionSceneFrame = getResByKey("Caption_Scene")

End Function

Function isCustomizedTpl() As Boolean
    Dim row As Integer
    isCustomizedTpl = False
    
    Dim MappingSiteTemplate As Worksheet
    Set MappingSiteTemplate = ThisWorkbook.Worksheets("MappingSiteTemplate")
    
    For row = 2 To MappingSiteTemplate.Range("a1048576").End(xlUp).row
        If Trim(MappingSiteTemplate.Cells(row, 2).value) <> "" Then
            isCustomizedTpl = True
            Exit Function
        End If
    Next
End Function
' ��ȡ�ı��ļ����ַ�������
Public Function readTxtFile(ByRef fileContent As String, ByVal fileName As String)
    Dim fileNum As Integer
    Dim fileBuf As String
    
    On Error GoTo E
    
    fileNum = FreeFile()
    Open fileName For Binary As #fileNum
        fileBuf = String(LOF(fileNum), Chr(0))
        Get #fileNum, , fileBuf
    Close #fileNum
    fileContent = fileBuf
    Exit Function
E:
    Log "Failed to read file." & vbCrLf & "FileName:" & fileName
End Function

Public Function isShtExists(ByVal shtName As String)
    isShtExists = False
    Dim shtIndex As Integer
    For shtIndex = 1 To ThisWorkbook.Sheets.count
        If ThisWorkbook.Sheets(shtIndex).name = shtName Then
            isShtExists = True
            Exit Function
        End If
    Next shtIndex
End Function
Public Function getPosRange(ByVal shtName As String, ByVal groupName As String, ByVal colName As String, Optional ByVal startRow = 1) As Range
    Dim findRange As Range
    Dim firAddr, curAddr As String
    firAddr = ""
    'ֻ����ǰ100��,��Ҫ�ǿ��ǵ�Common Dataҳ��������ͷ�������Ƚϴ�100�й�����
    If Not isShtExists(shtName) Then
        Exit Function
    End If
    Set findRange = ThisWorkbook.Sheets(shtName).Range("A" & startRow & ":IV" & "100").Find(colName, LookIn:=xlValues, LookAt:=xlWhole)
    '���ָ��������,����Ҫͬʱ�ж�����
    If Trim(groupName) <> "" Then
        Do While (Not findRange Is Nothing)
            If getGroupName(findRange.Offset(-1, 0)) = groupName Or findRange.Address = firAddr Then
                Exit Do
            End If
            If firAddr = "" Then
                firAddr = findRange.Address
            End If
            Set findRange = ThisWorkbook.Sheets(shtName).Range("A" & startRow & ":IV" & "100").FindNext(findRange)
        Loop
    End If
    Set getPosRange = findRange
End Function

Public Function getGroupName(ByRef startRange As Range) As String
    Dim curRange As Range
    Set curRange = startRange
    '��ͬһ�в���֮ǰ���У�����в�Ϊ����Ϊ����
    Do While Not curRange Is Nothing And curRange.value = ""
        Set curRange = curRange.Offset(0, -1)
    Loop
    If Not curRange Is Nothing Then
        getGroupName = curRange.value
    End If
End Function

Public Function getColNum(ByVal shtName As String, ByVal groupName As String, ByVal colName As String, Optional ByVal startRow = 1)
    Dim findRange As Range
    Set findRange = getPosRange(shtName, groupName, colName, startRow)
    If Not findRange Is Nothing Then
        getColNum = findRange.Column
    End If
End Function



Attribute VB_Name = "AutoDeployment"
Option Explicit

Private Const nameCol As Integer = 1 'Name�кţ�A��
Private Const EsnCol As Integer = 2 'ESN�кţ�B��
Private Const SubNetworkCol As Integer = 3 'SubNetwork�кţ�C��
Private Const SubAreaCol As Integer = 4 'Sub Area�кţ�D��

Private Const SolidColorIdx = 16 '�һ���Ԫ����ɫ
Private Const SolidPattern = xlGray16 '�һ���Ԫ����ʽ
Private Const NullPattern = xlNone '������Ԫ����ʽ
Private Const ESN_ValidLen = 23 'ESN�ֶεĳ�������
Private Const String_ValidLen = 64 'Name�еĳ�������

'��Ч�ַ�������SubNetwork�У������ж�Ҫ�����У�飬���£� ?:><*/\|"~!@#$^%&{}[]+=
'\?��\*��\\��\^��\[��\]����ת���ַ��������ַ�����ת��
Private Const InvalidCharCommonStr As String = "[\?:><\*/\\|""~!@#$\^%&{}\[\]+=]"
Private Const InvalidCharSubNetworkStr As String = "[\?:><\*""/\\|]" '��Ч�ַ���SubNetwork��Ҫ�����У�飬���£� ?:><*"/\|

Private Const DisplayInvalidCharCommonStr As String = """?:><*/\|""~!@#$^%&{}[]+=""" '��Msgbox����ʾ����Ч�ַ�
Private Const DisplayInvalidCharSubNetworkStr As String = """?:><*""/\|""" '��Msgbox����ʾ��SubNetwork��Ч�ַ�

Private regExp As Object '����һ��������ʽ����

Private connTypeCol As Long 'Connection Type�к�
Private authenticationTypeCol As Long 'Authentication Type�к�

'PNP�ַ��Ŀ���
Public Sub AutoDeploySheetChange(ByRef sheet As Worksheet, ByVal target As range)
    '����������к������
    Dim maxColumnNumber As Long, maxRowNumber As Long
    maxColumnNumber = sheet.range("XFD2").End(xlToLeft).column
    maxRowNumber = sheet.UsedRange.rows.count
    
    '�ҵ���Ҫ��֧���Ƶ�Connection Type��Authentication Type�����к�
    Call initConnTypeAndAuthenticationTypeCol(sheet)
    
    '��ʼ��������ʽ����
    Call initRegExp
    
    Dim cellRange As range
    '���޸ĵ�Range�е�ÿһ����Ԫ������У��
    For Each cellRange In target
        '����޸ĵĵ�Ԫ���ֵ��������л�����У���ֱ���˳�
        If (cellRange.row > maxRowNumber Or cellRange.column > maxColumnNumber) Then Exit For
        
        If cellRange.column = nameCol Then
            '��A��Name�н��г���У�飬���ܳ���64
            Call checkName(cellRange)
        ElseIf cellRange.column = EsnCol Then
            '����esn�г��ȵ�У��
            Call checkEsnLength(cellRange)
            
            '���������ַ���У��
            Call checkInvalidCharacter(cellRange, InvalidCharCommonStr, DisplayInvalidCharCommonStr)
            
        ElseIf cellRange.column = SubNetworkCol Then
            '����subNetwork�������ַ���У�飬��PNP����GUI���汣��һ�¼���
            Call checkInvalidCharacter(cellRange, InvalidCharSubNetworkStr, DisplayInvalidCharSubNetworkStr)
            
            '������Ҳ��ҪУ�鳤�ȣ�Ŀǰ����0-64���ʸ���Name�е�У�麯��
            Call checkName(cellRange)
        Else
            '��������ͨ�������ַ���У��
            Call checkInvalidCharacter(cellRange, InvalidCharCommonStr, DisplayInvalidCharCommonStr)
            
            '������Ҳ��ҪУ�鳤�ȣ�Ŀǰ����0-64���ʸ���Name�е�У�麯��
            Call checkName(cellRange)
            
            '����Connection Type��Authentication Type�����ֶεķ�֧����
            Call controlBranch(cellRange, sheet)
            
            '��Ҫ�Իһ���Ԫ���������������У�飬Ŀǰ�һ���Ԫ��ֻ�������һ��Authentication Type�����Ŀǰ����Else�н���У��
            Call checkGray(cellRange)
            
            '����Connection Type��Authentication Type�����ֶε�����ö��ֵУ��
            Call checkValueValidation(cellRange)
        End If
    Next cellRange
End Sub

'����һ����ѡ��ʼ�����л��в���
Private Function findCertainValColumnNumber(ByRef ws As Worksheet, ByVal rowNumber As Long, ByRef cellVal As Variant, Optional ByVal startColumn As Long = 1) As Long
    Dim currentCellVal As Variant
    Dim maxColumnNumber As Long, k As Long
    maxColumnNumber = ws.UsedRange.columns.count
    findCertainValColumnNumber = -1
    For k = startColumn To maxColumnNumber
        currentCellVal = ws.Cells(rowNumber, k).value
        If currentCellVal = cellVal Then
            findCertainValColumnNumber = k
            Exit For
        End If
    Next
End Function

'�ҵ���Ҫ��֧���Ƶ�Connection Type��Authentication Type�����к�
Private Sub initConnTypeAndAuthenticationTypeCol(ByRef ws As Worksheet)
    'ֻ�ڱ���״δ�ʱ��������ֵ��Ϊ0������Ҫ���ң���̲��ٽ��в���
    If connTypeCol = 0 Then
        connTypeCol = findCertainValColumnNumber(ws, 2, getResByKey("connType"))
    End If
    
    If authenticationTypeCol = 0 Then
        authenticationTypeCol = findCertainValColumnNumber(ws, 2, getResByKey("authenticationType"))
    End If
End Sub

'��ʼ��regExp
Private Sub initRegExp()
    If regExp Is Nothing Then
        Set regExp = CreateObject("VBSCRIPT.REGEXP")
    End If
End Sub

'У��A��Name�еĳ���
Private Sub checkName(ByRef cellRange As range)
    Dim nameValue As String
    Dim nResponse As Integer
    
    nameValue = Trim(cellRange.value)
    If cellRange.row > 2 And Len(nameValue) <> 0 Then
        If ((String_ValidLen < LenB(StrConv(nameValue, vbFromUnicode)))) Or (1 > LenB(StrConv(nameValue, vbFromUnicode))) Then
            nResponse = MsgBox(getResByKey("Limited Length") & "[0~64]", vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
            If nResponse = vbRetry Then
                cellRange.Select
            End If
            cellRange.value = ""
        End If
    End If
End Sub

'У��B��ESNֵ�ĳ���
Private Sub checkEsnLength(ByRef cellRange As range)
    Dim nResponse As Integer
    Dim esnValue As String

    esnValue = Trim(cellRange.value)
    If Len(esnValue) <> 0 Then
        If cellRange.row > 2 And (Len(esnValue) > ESN_ValidLen) Then
            nResponse = MsgBox(getResByKey("Length") & ("[0~" & ESN_ValidLen & "]"), vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
            If nResponse = vbRetry Then
                cellRange.Select
            End If
            cellRange.value = ""
        End If
    End If
End Sub

'У����д�ĵ�Ԫ���Ƿ�����Ч�ַ�����PNP����GUI���汣��һ�¼���
Private Sub checkInvalidCharacter(ByRef cellRange As range, ByRef invalidStr As String, ByRef displayStr As String)
    Dim cellValue As String
    Dim nResponse As Integer
    
    cellValue = Trim(cellRange.value)
    '��������ʽ�ж��Ƿ�����Ч�ַ�
    regExp.Pattern = invalidStr '����ƥ����ʽ
    If cellRange.row > 2 And regExp.test(cellValue) Then
       nResponse = MsgBox(getResByKey("InvalidCharacter") & displayStr, vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
       If nResponse = vbRetry Then
           cellRange.Select
       End If
       cellRange.value = ""
    End If
End Sub

'����Connection Type��Authentication Type�����ֶεķ�֧����
Private Sub controlBranch(ByRef cellRange As range, ByRef ws As Worksheet)
    If cellRange.row > 2 And cellRange.column = connTypeCol Then
        If cellRange.value = getResByKey("commConn") Then
            ws.Cells(cellRange.row, authenticationTypeCol).Interior.colorIndex = SolidColorIdx
            ws.Cells(cellRange.row, authenticationTypeCol).Interior.Pattern = SolidPattern
            ws.Cells(cellRange.row, authenticationTypeCol).value = ""
            ws.Cells(cellRange.row, authenticationTypeCol).Validation.ShowInput = False
        ElseIf (cellRange.value = getResByKey("sslConn") Or cellRange.value = "") Then
            ws.Cells(cellRange.row, authenticationTypeCol).Interior.colorIndex = NullPattern
            ws.Cells(cellRange.row, authenticationTypeCol).Interior.Pattern = NullPattern
            ws.Cells(cellRange.row, authenticationTypeCol).Validation.ShowInput = True
        End If
    End If
End Sub

'�һ���Ԫ����������ֵ
Private Sub checkGray(ByRef cellRange As range)
    If cellRange.Interior.colorIndex = SolidColorIdx And cellRange.Interior.Pattern = SolidPattern And cellRange.value <> "" Then
        cellRange.value = ""
        MsgBox (getResByKey("NoInput"))
    End If
End Sub

Private Sub checkValueValidation(ByRef cellRange As range)
    '���ֵΪ�գ�����ҪУ�飬ֱ���˳�
    If cellRange.value = "" Or cellRange.row < 3 Then Exit Sub
    
    Dim arrayRange As String
    If cellRange.column = connTypeCol Then
        arrayRange = getResByKey("connectionTypeRange")
    ElseIf cellRange.column = authenticationTypeCol Then
        arrayRange = getResByKey("authenticationTypeRange")
    End If

    Call Check_Value_In_Range("Enum", arrayRange, cellRange.value, cellRange, False)
End Sub



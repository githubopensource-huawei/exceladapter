Attribute VB_Name = "AutoDeployment"
Option Explicit

Private Const nameCol As Integer = 1 'Name列号，A列
Private Const EsnCol As Integer = 2 'ESN列号，B列
Private Const SubNetworkCol As Integer = 3 'SubNetwork列号，C列
Private Const SubAreaCol As Integer = 4 'Sub Area列号，D列

Private Const SolidColorIdx = 16 '灰化单元格颜色
Private Const SolidPattern = xlGray16 '灰化单元格样式
Private Const NullPattern = xlNone '正常单元格样式
Private Const ESN_ValidLen = 23 'ESN字段的长度限制
Private Const String_ValidLen = 64 'Name列的长度限制

'无效字符，除了SubNetwork列，其余列都要做这个校验，如下： ?:><*/\|"~!@#$^%&{}[]+=
'\?，\*，\\，\^，\[，\]都是转义字符，其余字符无需转义
Private Const InvalidCharCommonStr As String = "[\?:><\*/\\|""~!@#$\^%&{}\[\]+=]"
Private Const InvalidCharSubNetworkStr As String = "[\?:><\*""/\\|]" '无效字符，SubNetwork列要做这个校验，如下： ?:><*"/\|

Private Const DisplayInvalidCharCommonStr As String = """?:><*/\|""~!@#$^%&{}[]+=""" '在Msgbox中显示的无效字符
Private Const DisplayInvalidCharSubNetworkStr As String = """?:><*""/\|""" '在Msgbox中显示的SubNetwork无效字符

Private regExp As Object '声明一个正则表格式对象

Private connTypeCol As Long 'Connection Type列号
Private authenticationTypeCol As Long 'Authentication Type列号

'PNP字符的控制
Public Sub AutoDeploySheetChange(ByRef sheet As Worksheet, ByVal target As range)
    '先设置最大行和最大列
    Dim maxColumnNumber As Long, maxRowNumber As Long
    maxColumnNumber = sheet.range("XFD2").End(xlToLeft).column
    maxRowNumber = sheet.UsedRange.rows.count
    
    '找到需要分支控制的Connection Type和Authentication Type两列列号
    Call initConnTypeAndAuthenticationTypeCol(sheet)
    
    '初始化正则表达式对象
    Call initRegExp
    
    Dim cellRange As range
    '对修改的Range中的每一个单元格依次校验
    For Each cellRange In target
        '如果修改的单元格的值大于最大行或最大列，则直接退出
        If (cellRange.row > maxRowNumber Or cellRange.column > maxColumnNumber) Then Exit For
        
        If cellRange.column = nameCol Then
            '对A列Name列进行长度校验，不能超过64
            Call checkName(cellRange)
        ElseIf cellRange.column = EsnCol Then
            '进行esn列长度的校验
            Call checkEsnLength(cellRange)
            
            '进行特殊字符的校验
            Call checkInvalidCharacter(cellRange, InvalidCharCommonStr, DisplayInvalidCharCommonStr)
            
        ElseIf cellRange.column = SubNetworkCol Then
            '进行subNetwork列特殊字符的校验，与PNP导出GUI界面保持一致即可
            Call checkInvalidCharacter(cellRange, InvalidCharSubNetworkStr, DisplayInvalidCharSubNetworkStr)
            
            '其余列也需要校验长度，目前都是0-64，故复用Name列的校验函数
            Call checkName(cellRange)
        Else
            '其余列做通用特殊字符的校验
            Call checkInvalidCharacter(cellRange, InvalidCharCommonStr, DisplayInvalidCharCommonStr)
            
            '其余列也需要校验长度，目前都是0-64，故复用Name列的校验函数
            Call checkName(cellRange)
            
            '进行Connection Type和Authentication Type两个字段的分支控制
            Call controlBranch(cellRange, sheet)
            
            '需要对灰化单元格做不允许输入的校验，目前灰化单元格只会在最后一列Authentication Type，因此目前放在Else中进行校验
            Call checkGray(cellRange)
            
            '进行Connection Type和Authentication Type两个字段的输入枚举值校验
            Call checkValueValidation(cellRange)
        End If
    Next cellRange
End Sub

'新增一个可选起始搜索行或列参数
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

'找到需要分支控制的Connection Type和Authentication Type两列列号
Private Sub initConnTypeAndAuthenticationTypeCol(ByRef ws As Worksheet)
    '只在表格首次打开时，这两个值才为0，才需要查找，后继不再进行查找
    If connTypeCol = 0 Then
        connTypeCol = findCertainValColumnNumber(ws, 2, getResByKey("connType"))
    End If
    
    If authenticationTypeCol = 0 Then
        authenticationTypeCol = findCertainValColumnNumber(ws, 2, getResByKey("authenticationType"))
    End If
End Sub

'初始化regExp
Private Sub initRegExp()
    If regExp Is Nothing Then
        Set regExp = CreateObject("VBSCRIPT.REGEXP")
    End If
End Sub

'校验A列Name中的长度
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

'校验B列ESN值的长度
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

'校验填写的单元格是否含有无效字符，与PNP导出GUI界面保持一致即可
Private Sub checkInvalidCharacter(ByRef cellRange As range, ByRef invalidStr As String, ByRef displayStr As String)
    Dim cellValue As String
    Dim nResponse As Integer
    
    cellValue = Trim(cellRange.value)
    '用正则表格式判断是否含有无效字符
    regExp.Pattern = invalidStr '设置匹配样式
    If cellRange.row > 2 And regExp.test(cellValue) Then
       nResponse = MsgBox(getResByKey("InvalidCharacter") & displayStr, vbRetryCancel + vbCritical + vbApplicationModal + vbDefaultButton1, getResByKey("Warning"))
       If nResponse = vbRetry Then
           cellRange.Select
       End If
       cellRange.value = ""
    End If
End Sub

'进行Connection Type和Authentication Type两个字段的分支控制
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

'灰化单元格不允许输入值
Private Sub checkGray(ByRef cellRange As range)
    If cellRange.Interior.colorIndex = SolidColorIdx And cellRange.Interior.Pattern = SolidPattern And cellRange.value <> "" Then
        cellRange.value = ""
        MsgBox (getResByKey("NoInput"))
    End If
End Sub

Private Sub checkValueValidation(ByRef cellRange As range)
    '如果值为空，则不需要校验，直接退出
    If cellRange.value = "" Or cellRange.row < 3 Then Exit Sub
    
    Dim arrayRange As String
    If cellRange.column = connTypeCol Then
        arrayRange = getResByKey("connectionTypeRange")
    ElseIf cellRange.column = authenticationTypeCol Then
        arrayRange = getResByKey("authenticationTypeRange")
    End If

    Call Check_Value_In_Range("Enum", arrayRange, cellRange.value, cellRange, False)
End Sub



Attribute VB_Name = "MakeModifyStyleSheets"
Option Explicit

Private logPath As String

Private Function getSheetNumber(ByRef sheetNumber As Long, ByRef inputParamterStr As String, ByRef trace As String) As Boolean
    On Error GoTo ErrorHandler
    getSheetNumber = True
    Dim parameterArray As Variant, sheetNumberArray As Variant, ele As Variant
    
    inputParamterStr = FileSub.readUTF8File(logPath)
    parameterArray = Split(inputParamterStr, vbCr + vbLf)
    
    For Each ele In parameterArray
        sheetNumberArray = Split(ele, "=")
        If "ModifyStyleSheetNumber" = Trim(sheetNumberArray(0)) Then
            sheetNumber = CLng(sheetNumberArray(1))
            Exit Function
        End If
    Next ele

    Exit Function
ErrorHandler:
    getSheetNumber = False
    trace = trace & "GetSheetNumber Fails! " & Err.Description
End Function

'Public Sub releaseClassResource(ByRef applicationFlag As CApplicationFlags, ByRef boardStyle As CBoardStyle)
'    Set applicationFlag = Nothing
'    Set boardStyle = Nothing
'End Sub

Public Sub makeMultiBoardSytleSheets()
    On Error GoTo ErrorHandler

    logPath = ThisWorkbook.Path + "\Parameter.ini"
    
    Dim applicationFlag As New CApplicationFlags
    Call applicationFlag.init
    
    Dim trace As String
    trace = ""
    
    Dim sheetNumber As Long
    Dim inputParamterStr As String
    If False = getSheetNumber(sheetNumber, inputParamterStr, trace) Then GoTo ErrorHandler
    
    Dim boardStyle As New CBoardStyleSheetNumberChange
    Call boardStyle.keepOneBaseBoardStyleSheet
    Call boardStyle.addBoardStyleSheets(sheetNumber)
    
    ThisWorkbook.Save
    'Call releaseClassResource(applicationFlag, boardStyle)
    Call FileSub.WriteUTF8File(inputParamterStr & vbCrLf & "Log=Make board style sheets successfully." & trace, logPath, False)
    'ThisWorkbook.Close saveChanges:=False
    Exit Sub
ErrorHandler:
    ThisWorkbook.Save
    'Call releaseClassResource(applicationFlag, boardStyle)
    Call FileSub.WriteUTF8File(inputParamterStr & vbCrLf & "Log=Make board style sheets unsuccessfully!" & trace & " Error Info: " & Err.Description, logPath, False)
    'ThisWorkbook.Close saveChanges:=False
End Sub

'�޸İ汾�ţ��Ǳ�汾���ݵ�������Ҫ����ǰ̨�����޸İ汾��
Private Function getVersion(ByRef newVersion As String, ByRef inputParamterStr As String, ByRef trace As String, ByRef logPath As String) As Boolean
    On Error GoTo ErrorHandler
    getVersion = True
    Dim parameterArray As Variant, versionArray As Variant, ele As Variant
    
    inputParamterStr = FileSub.readUTF8File(logPath)
    parameterArray = Split(inputParamterStr, vbCr + vbLf)
    
    For Each ele In parameterArray
        versionArray = Split(ele, "=")
        If "NewVersion" = Trim(versionArray(0)) Then
            newVersion = versionArray(1)
            Exit Function
        End If
    Next ele

    Exit Function
ErrorHandler:
    getVersion = False
    trace = trace & "GetVersion Fails! " & Err.Description
End Function

Public Sub changeVersion()
    On Error GoTo ErrorHandler
    
    Dim changeVersionLogPath As String
    changeVersionLogPath = ThisWorkbook.Path + "\ChangeVersion.ini"
    
    Dim applicationFlag As New CApplicationFlags
    Call applicationFlag.init
    
    Dim trace As String
    trace = ""
    
    Dim newVersion As String
    Dim inputParamterStr As String
    '��ð汾�ź������ļ���ԭ�ַ���
    If False = getVersion(newVersion, inputParamterStr, trace, changeVersionLogPath) Then GoTo ErrorHandler

    If newVersion <> "" Then
        '�����µİ汾��
        Dim changeVersionClass As New CChangeVersion
        Call changeVersionClass.changeVersion(ThisWorkbook, newVersion)
    End If
    
    ThisWorkbook.Save
    Call FileSub.WriteUTF8File(inputParamterStr & vbCrLf & "Log=Change version successfully." & trace, changeVersionLogPath, False)
    
    Exit Sub
ErrorHandler:
    ThisWorkbook.Save
    Call FileSub.WriteUTF8File(inputParamterStr & vbCrLf & "Log=Change version unsuccessfully!" & trace & " Error Info: " & Err.Description, changeVersionLogPath, False)
End Sub




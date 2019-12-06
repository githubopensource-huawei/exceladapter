Attribute VB_Name = "MakeBoardStyleSheets"
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
        If "BoardStyleSheetNumber" = Trim(sheetNumberArray(0)) Then
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

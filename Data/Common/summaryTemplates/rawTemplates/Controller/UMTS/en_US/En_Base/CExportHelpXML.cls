VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CExportHelpXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const strFileName As String = "PreInfo.xml"
Private strFilePath As String


'导出 帮助文档
Public Function ExportXMLFile() As Boolean
    
'    Dim fWriPreInfo As New CUtf8File
'    Dim SgwAddrType As String
'    Dim M2000AddrType As String
'    Dim path As String
'
'    Dim fWriErr As New CUtf8File
'
'    If shtExport.checkdata = False Then
'        Exit Function
'    End If
'
'
'    path = Application.GetSaveAsFilename("PreInfo.xml", "XML Files (*.xml), *.xml", 1, "Export PreInfo.xml")
'
'    If path = "False" Then
'        ExportXMLFile = False
'        Exit Function
'    End If
'
'    If Not fWriPreInfo.CreateXmlFile(path) Then
'        MsgBox "Create PreInfo.xml Failed！"
'        ExportXMLFile = False
'        Exit Function
'    End If
'
'    writeHead fWriPreInfo
'
'    SgwAddrType = GetAddrTypeStr(Trim(GetCell(shtExport, ROW_SEGW_ADDRTYPE, COL_VALUE)))
'    M2000AddrType = GetAddrTypeStr(Trim(GetCell(shtExport, ROW_M2000_ADDRTYPE, COL_VALUE)))
'
'
'    fWriPreInfo.AppendLine GetLeftBlankStr(2) & "<DHCPSRVADDR AddrType=""" & M2000AddrType & """ IpAddr=""" & Trim(GetCell(shtExport, ROW_M2000_IP, COL_VALUE)) _
'                        & """ URL=""" & Trim(GetCell(shtExport, ROW_M2000_DOMAIN, COL_VALUE)) & """/>"
'
'    fWriPreInfo.AppendLine GetLeftBlankStr(2) & "<SEGWADDR AddrType=""" & SgwAddrType & """ IpAddr=""" & Trim(GetCell(shtExport, ROW_SEGW_IP, COL_VALUE)) _
'                        & """ URL=""" & Trim(GetCell(shtExport, ROW_SEGW_DOMAIN, COL_VALUE)) & """/>"
'
'
'    fWriPreInfo.AppendLine GetLeftBlankStr(2) & "<AUTHMETH AuthMeth=""" & shtExport.getAuthMethInt(GetCell(shtExport, ROW_AUTH_METH, COL_VALUE)) & """ PreShareKey=""" & Trim(GetCell(shtExport, ROW_KEY, COL_VALUE)) & """/>"
'    fWriPreInfo.AppendLine GetLeftBlankStr(2) & "<DHCPVLANID VlanId=""" & Trim(GetCell(shtExport, ROW_DHCP_VLANID, COL_VALUE)) & """/>"
'    fWriPreInfo.AppendLine GetLeftBlankStr(2) & "<SEGWLOCNAME Name=""" & Trim(GetCell(shtExport, ROW_SEGW_LOCALNAME, COL_VALUE)) & """/>"
'    fWriPreInfo.AppendLine GetLeftBlankStr(2) & "<CROSSCAFILEPATH PathName=""" & Trim(GetCell(shtExport, ROW_LIC_PATH, COL_VALUE)) & """/>"
'    fWriPreInfo.AppendLine GetLeftBlankStr(2) & "<CRLFILEPATH PathName=""" & Trim(GetCell(shtExport, ROW_CRL_FILEPATH, COL_VALUE)) & """/>"
'    fWriPreInfo.AppendLine GetLeftBlankStr(2) & "<CRSERVER IpAddr=""" & Trim(GetCell(shtExport, ROW_CR_SERVIP, COL_VALUE)) & """ User=""" & _
'                         Trim(GetCell(shtExport, ROW_CR_SERVNAME, COL_VALUE)) & """ Passward=""" & Trim(GetCell(shtExport, ROW_CR_SERVPWD, COL_VALUE)) & _
'                         """ DownLoadType=""" & Replace(Trim(GetCell(shtExport, ROW_CR_DLDTYPE, COL_VALUE)), "NULL", "") & _
'                         """ Dn=""" & Trim(GetCell(shtExport, ROW_CR_DN, COL_VALUE)) & _
'                         """ Searchdn=""" & Trim(GetCell(shtExport, ROW_CR_SERACHDN, COL_VALUE)) & """/>"
'
'    Dim CaProtocol As String
'    If Trim(GetCell(shtExport, ROW_CASERVER_PROTOCOL, COL_VALUE)) = "http" Then
'        CaProtocol = "0"
'    ElseIf Trim(GetCell(shtExport, ROW_CASERVER_PROTOCOL, COL_VALUE)) = "https" Then
'        CaProtocol = "1"
'    Else
'
'    End If
'
'
'
'    fWriPreInfo.AppendLine GetLeftBlankStr(2) & "<CASERVER Name=""" & Trim(GetCell(shtExport, ROW_CASERVER_NAME, COL_VALUE)) & _
'                           """ Protocol=""" & CaProtocol & _
'                           """ IpAddr=""" & Trim(GetCell(shtExport, ROW_CASERVER_IP, COL_VALUE)) & _
'                           """ Port=""" & Trim(GetCell(shtExport, ROW_CASERVER_PORT, COL_VALUE)) & _
'                           """ Path=""" & Trim(GetCell(shtExport, ROW_CASERVER_PATH, COL_VALUE)) & _
'                           """/>"
'
'
'    M2000AddrType = GetAddrTypeStr(Trim(GetCell(shtExport, ROW_1M2000_ADDRTYPE, COL_VALUE)))
'
'
'    fWriPreInfo.AppendLine GetLeftBlankStr(2) & "<NMS1ADDR AddrType=""" & M2000AddrType & _
'                           """ IpAddr=""" & Trim(GetCell(shtExport, ROW_1M2000_IP, COL_VALUE)) & _
'                           """ URL=""" & Trim(GetCell(shtExport, ROW_1M2000_URL, COL_VALUE)) & """/>"
'
'    M2000AddrType = GetAddrTypeStr(Trim(GetCell(shtExport, ROW_2M2000_ADDRTYPE, COL_VALUE)))
'
'    fWriPreInfo.AppendLine GetLeftBlankStr(2) & "<NMS2ADDR AddrType=""" & M2000AddrType & _
'                           """ IpAddr=""" & Trim(GetCell(shtExport, ROW_2M2000_IP, COL_VALUE)) & _
'                           """ URL=""" & Trim(GetCell(shtExport, ROW_2M2000_URL, COL_VALUE)) & """/>"
'
'    M2000AddrType = GetAddrTypeStr(Trim(GetCell(shtExport, ROW_2DHCPSERVER_ADDRTYPE, COL_VALUE)))
'
'    fWriPreInfo.AppendLine GetLeftBlankStr(2) & "<DHCPSRV2ADDR AddrType=""" & M2000AddrType & _
'                           """ IpAddr=""" & Trim(GetCell(shtExport, ROW_2DHCPSERVER_IP, COL_VALUE)) & _
'                           """ URL=""" & Trim(GetCell(shtExport, ROW_2DHCPSERVER_URL, COL_VALUE)) & """/>"
'
'
'
'    fWriPreInfo.AppendLine GetLeftBlankStr(2) & "<FILES>"
'
'    If Trim(GetCell(shtExport, ROW_USB_ROOTNAME, COL_VALUE)) <> "" Then
'        fWriPreInfo.AppendLine GetLeftBlankStr(3) & "<File Name=""" & shtExport.strRootFilePath & Trim(GetCell(shtExport, ROW_USB_ROOTNAME, COL_VALUE)) & """/>"
'    End If
'
'    If Trim(GetCell(shtExport, ROW_USB_CRLNAME, COL_VALUE)) <> "" Then
'        fWriPreInfo.AppendLine GetLeftBlankStr(3) & "<File Name=""" & shtExport.strCRLFilePath & Trim(GetCell(shtExport, ROW_USB_CRLNAME, COL_VALUE)) & """/>"
'    End If
'    If Trim(GetCell(shtExport, ROW_USB_CROSSNAME, COL_VALUE)) <> "" Then
'        fWriPreInfo.AppendLine GetLeftBlankStr(3) & "<File Name=""" & shtExport.strCrossFilePath & Trim(GetCell(shtExport, ROW_USB_CROSSNAME, COL_VALUE)) & """/>"
'    End If
'
'    fWriPreInfo.AppendLine GetLeftBlankStr(2) & "</FILES>"
'
'
'    fWriPreInfo.AppendLine GetLeftBlankStr(1) & "</ConfigInfo>"
'    fWriPreInfo.AppendLine "</PreInfo>"
'
'    fWriPreInfo.CloseXmlFile
'
'    MsgBox "Export success! File path:" & path
'
'    ExportXMLFile = True
End Function



''写ditemap文件头信息
Private Sub writeHead(fwriFile As CUtf8File)

    fwriFile.AppendLine "<?xml version=""1.0"" encoding=""UTF-8""?>"
    fwriFile.AppendLine "<PreInfo Version=""1"" TimeStamp=""" & Time & """>"
    fwriFile.AppendLine GetLeftBlankStr(1) & "<ConfigInfo TimeStamp=""" & Time & """>"
  
End Sub

Private Function GetAddrTypeStr(strAddrType As String) As String
'    If strAddrType = shtExport.getIpTypeStr Then
'        GetAddrTypeStr = "1"
'    ElseIf strAddrType = shtExport.getUrlTypeStr Then
'        GetAddrTypeStr = "0"
'    ElseIf strAddrType = "" Then
'        GetAddrTypeStr = ""
'    Else
'        MsgBox "Address Type is Wrong！"
'    End If
    
End Function

Private Function GetLeftBlankStr(level As Long) As String
    GetLeftBlankStr = Space(level * 4)
End Function









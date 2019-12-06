Attribute VB_Name = "sheetVisiableController"
Public Sub HideAllSheets()
  Sheets("MPGRP").Visible = False
  Sheets("MPLNK").Visible = False
  Sheets("ETHIP").Visible = False
  Sheets("PPPLNK").Visible = False
  Sheets("BTS").Visible = False
  Sheets("ADJNODE").Visible = False

  Sheets("IPRT").Visible = False
  Sheets("BTSIP").Visible = False
  Sheets("BTSETHPORT").Visible = False
  Sheets("BTSIPCLKPARA").Visible = False
  Sheets("BTSPPPLNK").Visible = False
  Sheets("BTSMPGRP").Visible = False
  Sheets("BTSMPLNK").Visible = False
  Sheets("BTSESN").Visible = False
  Sheets("BTSBFD").Visible = False
  Sheets("BTSIPRT").Visible = False
  Sheets("BTSIPRTBIND").Visible = False
  Sheets("BTSCONNECT").Visible = False
  Sheets("BTSIDLETS").Visible = False
  Sheets("BTSMONITORTS").Visible = False
  Sheets("GCELL").Visible = False
  Sheets("PTPBVC").Visible = False
  Sheets("GCELLGPRS").Visible = False
  Sheets("GTRX").Visible = False
  Sheets("DXXTSEXGRELATION").Visible = False
  Sheets("BTSDHCPSVRIP").Visible = False
  Sheets("BTSDEVIP").Visible = False

  Sheets("IPLOGICPORT").Visible = False
  Sheets("DEVIP").Visible = False
  
  Sheets("ADJMAP").Visible = False
  Sheets("BTSVLAN").Visible = False
  Sheets("BTSVLANCLASS").Visible = False
  Sheets("BTSVLANMAP").Visible = False
  Sheets("BTSFORBIDTS").Visible = False
  Sheets("BTSOMLTS").Visible = False
  Sheets("GCELLPSBASE").Visible = False

  
  Sheets("BTSIKECFG").Visible = False
  Sheets("BTSIKEPROPOSAL").Visible = False
  Sheets("BTSIKEPEER").Visible = False
  Sheets("BTSIPSECPROPOSAL").Visible = False
  Sheets("BTSIPSECPOLICY").Visible = False
  Sheets("BTSIPSECBIND").Visible = False
  Sheets("BTSIPSECDTNL").Visible = False

    Sheets("BTSTUNNEL").Visible = False

Sheets("BTSCLK").Visible = False
Sheets("BTSDSCPMAP").Visible = False
Sheets("BTSACL").Visible = False
Sheets("BTSACLRULE").Visible = False
Sheets("BTSPACKETFILTER").Visible = False
Sheets("BTSIPGUARD").Visible = False
Sheets("BTSFLOODDEFEND").Visible = False
Sheets("BTSCERTMK").Visible = False
Sheets("BTSAPPCERT").Visible = False
Sheets("BTSIKECFG").Visible = False
Sheets("BTSIKEPROPOSAL").Visible = False
Sheets("BTSIKEPEER").Visible = False
Sheets("BTSIPSECPROPOSAL").Visible = False
Sheets("BTSIPSECPOLICY").Visible = False
Sheets("BTSIPSECBIND").Visible = False
Sheets("BTSIPSECDTNL").Visible = False
Sheets("BTSCERTREQ").Visible = False
Sheets("BTSTRUSTCERT").Visible = False
Sheets("BTSCRL").Visible = False
Sheets("BTSCRLTSK").Visible = False
Sheets("BTSCRLPOLICY").Visible = False
Sheets("BTSCERTCHKTSK").Visible = False
Sheets("BTSCA").Visible = False
Sheets("BTSCERTDEPLOY").Visible = False
Sheets("ETHTRKIP").Visible = False
Sheets("ADJPPPBIND").Visible = False
Sheets("ADJNODEDIP").Visible = False
Sheets("IPPOOL").Visible = False
Sheets("IPPOOLIP").Visible = False
Sheets("IPPOOLPM").Visible = False
Sheets("IPPOOLMUX").Visible = False
Sheets("SRCIPRT").Visible = False
Sheets("BTSSHARING").Visible = False
Sheets("BTSEXTOPIP").Visible = False
Sheets("BTSEXTOPMAP").Visible = False
Sheets("BTSEXTOPABISMUXFLOW").Visible = False
Sheets("BTSEXTOPIPPM").Visible = False
Sheets("BTSARPSESSION").Visible = False
Sheets("ALGCTRLPARA").Visible = False
Sheets("ETHTRK").Visible = False
Sheets("ETHTRKLNK").Visible = False
End Sub

Public Sub ShowTDMSheets(bShowTDM As Boolean)
  If bShowTDM Then
    Sheets("BTS").Visible = True
    Sheets("BTSCONNECT").Visible = True
    Sheets("BTSIDLETS").Visible = True
    Sheets("BTSMONITORTS").Visible = True
    Sheets("GCELL").Visible = True
    Sheets("PTPBVC").Visible = True
    Sheets("GCELLGPRS").Visible = True
    Sheets("GTRX").Visible = True
    Sheets("BTSFORBIDTS").Visible = True
    Sheets("BTSOMLTS").Visible = True
    Sheets("GCELLPSBASE").Visible = True
    Sheets("BTSSHARING").Visible = True
  End If
End Sub
Public Sub ShowcbTDMDXXSheets(bShowTDMDXX As Boolean)
    If bShowTDMDXX Then
    Sheets("BTS").Visible = True
    Sheets("BTSCONNECT").Visible = True
    Sheets("BTSIDLETS").Visible = True
    Sheets("BTSMONITORTS").Visible = True
    Sheets("GCELL").Visible = True
    Sheets("PTPBVC").Visible = True
    Sheets("GCELLGPRS").Visible = True
    Sheets("GTRX").Visible = True
    Sheets("DXXTSEXGRELATION").Visible = True
    Sheets("BTSFORBIDTS").Visible = True
    Sheets("BTSOMLTS").Visible = True
    Sheets("GCELLPSBASE").Visible = True
    Sheets("BTSSHARING").Visible = True
  End If
End Sub

Public Sub ShowIPOESheets(bShowIPOE As Boolean)
  If bShowIPOE Then
      Sheets("ETHIP").Visible = True
    Sheets("MPGRP").Visible = True
    Sheets("MPLNK").Visible = True
    Sheets("PPPLNK").Visible = True
    Sheets("BTS").Visible = True
    Sheets("ADJNODE").Visible = True

    Sheets("IPRT").Visible = True
    Sheets("BTSIP").Visible = True
    Sheets("BTSPPPLNK").Visible = True
    Sheets("BTSMPGRP").Visible = True
    Sheets("BTSMPLNK").Visible = True
    Sheets("BTSESN").Visible = True
    Sheets("BTSBFD").Visible = True
    Sheets("BTSIPRT").Visible = True
    Sheets("BTSIPRTBIND").Visible = True
    Sheets("BTSCONNECT").Visible = True
    Sheets("BTSMONITORTS").Visible = True
    Sheets("GCELL").Visible = True
    Sheets("PTPBVC").Visible = True
    Sheets("GCELLGPRS").Visible = True
    Sheets("GTRX").Visible = True
    Sheets("BTSDEVIP").Visible = True

    Sheets("IPLOGICPORT").Visible = True
    Sheets("DEVIP").Visible = True
    Sheets("BTSDHCPSVRIP").Visible = True
    Sheets("ADJMAP").Visible = True
    Sheets("BTSFORBIDTS").Visible = True
    Sheets("BTSVLAN").Visible = True
    Sheets("BTSVLANCLASS").Visible = True
    Sheets("BTSVLANMAP").Visible = True


    Sheets("BTSTUNNEL").Visible = True
    Sheets("BTSCLK").Visible = True
   Sheets("BTSDSCPMAP").Visible = True
   Sheets("ETHTRKIP").Visible = True
   Sheets("ADJPPPBIND").Visible = True
   Sheets("ADJNODEDIP").Visible = True
   Sheets("IPPOOL").Visible = True
   Sheets("IPPOOLIP").Visible = True
   Sheets("IPPOOLPM").Visible = True
   Sheets("IPPOOLMUX").Visible = True
   Sheets("GCELLPSBASE").Visible = True
   Sheets("BTSSHARING").Visible = True
Sheets("SRCIPRT").Visible = True
Sheets("BTSEXTOPIP").Visible = True
Sheets("BTSEXTOPMAP").Visible = True
Sheets("BTSEXTOPABISMUXFLOW").Visible = True
Sheets("BTSEXTOPIPPM").Visible = True
Sheets("BTSARPSESSION").Visible = True
Sheets("ALGCTRLPARA").Visible = True
Sheets("ETHTRK").Visible = True
Sheets("ETHTRKLNK").Visible = True
  End If
End Sub

Public Sub ShowIPFEandE1T1(bShowIPFE As Boolean)
  If bShowIPFE Then
  ShowIPFESheets (True)
  ShowIPOESheets (True)

  End If
End Sub
Public Sub ShowIPFESheets(bShowIPFE As Boolean)
  If bShowIPFE Then
    Sheets("ETHIP").Visible = True
    Sheets("BTS").Visible = True
    Sheets("ADJNODE").Visible = True

    Sheets("IPRT").Visible = True
    Sheets("BTSIP").Visible = True
    Sheets("BTSETHPORT").Visible = True
    Sheets("BTSIPCLKPARA").Visible = True
    Sheets("BTSESN").Visible = True
    Sheets("BTSBFD").Visible = True
    Sheets("BTSIPRT").Visible = True
    Sheets("BTSIPRTBIND").Visible = True
    Sheets("GCELL").Visible = True
    Sheets("PTPBVC").Visible = True
    Sheets("GCELLGPRS").Visible = True
    Sheets("GTRX").Visible = True
    Sheets("BTSDHCPSVRIP").Visible = True
    Sheets("BTSDEVIP").Visible = True

    Sheets("IPLOGICPORT").Visible = True
    Sheets("DEVIP").Visible = True
    Sheets("ADJMAP").Visible = True
    Sheets("BTSVLAN").Visible = True
    Sheets("BTSVLANCLASS").Visible = True
    Sheets("BTSVLANMAP").Visible = True

    Sheets("BTSCLK").Visible = True
    Sheets("BTSDSCPMAP").Visible = True

    
    Sheets("BTSIKECFG").Visible = True
    Sheets("BTSIKEPROPOSAL").Visible = True
    Sheets("BTSIKEPEER").Visible = True
    Sheets("BTSIPSECPROPOSAL").Visible = True
    Sheets("BTSIPSECPOLICY").Visible = True
    Sheets("BTSIPSECBIND").Visible = True
    Sheets("BTSIPSECDTNL").Visible = True
    Sheets("BTSTUNNEL").Visible = True
Sheets("BTSACL").Visible = True
Sheets("BTSACLRULE").Visible = True
Sheets("BTSPACKETFILTER").Visible = True
Sheets("BTSIPGUARD").Visible = True
Sheets("BTSFLOODDEFEND").Visible = True
Sheets("BTSCERTMK").Visible = True
Sheets("BTSAPPCERT").Visible = True
Sheets("BTSIKECFG").Visible = True
Sheets("BTSIKEPROPOSAL").Visible = True
Sheets("BTSIKEPEER").Visible = True
Sheets("BTSIPSECPROPOSAL").Visible = True
Sheets("BTSIPSECPOLICY").Visible = True
Sheets("BTSIPSECBIND").Visible = True
Sheets("BTSIPSECDTNL").Visible = True
Sheets("BTSCERTREQ").Visible = True
Sheets("BTSTRUSTCERT").Visible = True
Sheets("BTSCRL").Visible = True
Sheets("BTSCRLTSK").Visible = True
Sheets("BTSCRLPOLICY").Visible = True
Sheets("BTSCERTCHKTSK").Visible = True
Sheets("BTSCA").Visible = True
Sheets("BTSCERTDEPLOY").Visible = True
   Sheets("ETHTRKIP").Visible = True
   Sheets("ADJPPPBIND").Visible = True
   Sheets("ADJNODEDIP").Visible = True
   Sheets("IPPOOL").Visible = True
   Sheets("IPPOOLIP").Visible = True
   Sheets("IPPOOLPM").Visible = True
   Sheets("IPPOOLMUX").Visible = True
   Sheets("GCELLPSBASE").Visible = True
   Sheets("BTSSHARING").Visible = True
Sheets("SRCIPRT").Visible = True
Sheets("BTSEXTOPIP").Visible = True
Sheets("BTSEXTOPMAP").Visible = True
Sheets("BTSEXTOPABISMUXFLOW").Visible = True
Sheets("BTSEXTOPIPPM").Visible = True
Sheets("BTSARPSESSION").Visible = True
Sheets("ALGCTRLPARA").Visible = True
Sheets("BTSCONNECT").Visible = True
Sheets("ETHTRK").Visible = True
Sheets("ETHTRKLNK").Visible = True
  End If
End Sub





Attribute VB_Name = "userDefineController"

Public Sub ShowIPOE(bShowIPOE As Boolean)
  If bShowIPOE Then
    Sheets("ETHIP").Visible = True
    Sheets("MPGRP").Visible = True
    Sheets("MPLNK").Visible = True
    Sheets("PPPLNK").Visible = True
    Sheets("ADJPPPBIND").Visible = True
    Sheets("BTS").Visible = True
    Sheets("ADJNODE").Visible = True
    Sheets("ADJNODEDIP").Visible = True
   ' Sheets("IPPATH").Visible = True
    Sheets("IPRT").Visible = True
    Sheets("BTSIP").Visible = True
    Sheets("BTSPPPLNK").Visible = True
    Sheets("BTSMPGRP").Visible = True
    Sheets("BTSMPLNK").Visible = True
    Sheets("BTSBFD").Visible = True
    Sheets("BTSIPRT").Visible = True
    Sheets("BTSIPRTBIND").Visible = True
    Sheets("BTSCONNECT").Visible = True
    Sheets("BTSMONITORTS").Visible = True
    Sheets("BTSDEVIP").Visible = True
    'Sheets("RSCGRP").Visible = True
    Sheets("IPLOGICPORT").Visible = True
    Sheets("DEVIP").Visible = True
    Sheets("BTSDHCPSVRIP").Visible = True
    Sheets("BTSFORBIDTS").Visible = True
    Sheets("ADJMAP").Visible = True
    Sheets("BTSCLK").Visible = True
    Sheets("BTSDSCPMAP").Visible = True
    Sheets("BTSTUNNEL").Visible = True
    Sheets("BTSESN").Visible = True
      Sheets("IPPOOL").Visible = True
      Sheets("IPPOOLMUX").Visible = True
  Sheets("IPPOOLIP").Visible = True
       Sheets("IPPOOLPM").Visible = True
  Sheets("SRCIPRT").Visible = True
  
  Sheets("ETHTRKIP").Visible = True
  
  Sheets("BTSEQUIPMENT").Visible = True
  Sheets("BTSLR").Visible = True
  Sheets("BTSGTRANSPARA").Visible = True
  
  Sheets("BTSABISMUXFLOW").Visible = True
  Sheets("BTSABISPRIMAP").Visible = True
  
 Sheets("SRCIPRT").Visible = True
Sheets("BTSEXTOPIP").Visible = True
Sheets("BTSEXTOPMAP").Visible = True
Sheets("BTSEXTOPABISMUXFLOW").Visible = True
Sheets("BTSEXTOPIPPM").Visible = True
Sheets("BTSARPSESSION").Visible = True
Sheets("ALGCTRLPARA").Visible = True
  End If
End Sub

Public Sub HideAllSheets()
  Sheets("BTSDSCPMAP").Visible = False
  Sheets("MPGRP").Visible = False
  Sheets("MPLNK").Visible = False
  Sheets("ETHIP").Visible = False
  Sheets("PPPLNK").Visible = False
  Sheets("BTS").Visible = False
  Sheets("ADJNODE").Visible = False
  Sheets("ADJNODEDIP").Visible = False
 ' Sheets("IPPATH").Visible = False
  Sheets("IPRT").Visible = False
  Sheets("BTSIP").Visible = False
  Sheets("BTSETHPORT").Visible = False
  Sheets("BTSIPCLKPARA").Visible = False
  Sheets("BTSPPPLNK").Visible = False
  Sheets("BTSMPGRP").Visible = False
  Sheets("BTSMPLNK").Visible = False
  Sheets("BTSBFD").Visible = False
  Sheets("BTSIPRT").Visible = False
  Sheets("BTSIPRTBIND").Visible = False
  Sheets("BTSCONNECT").Visible = False
  Sheets("BTSMONITORTS").Visible = False
  Sheets("BTSDHCPSVRIP").Visible = False
  Sheets("BTSDEVIP").Visible = False
  'Sheets("RSCGRP").Visible = False
  Sheets("IPLOGICPORT").Visible = False
  Sheets("DEVIP").Visible = False
  Sheets("ADJMAP").Visible = False
  Sheets("BTSVLAN").Visible = False
  Sheets("BTSVLANCLASS").Visible = False
  Sheets("BTSVLANMAP").Visible = False
  Sheets("BTSFORBIDTS").Visible = False
  
  Sheets("BTSIKECFG").Visible = False
  Sheets("BTSIKEPROPOSAL").Visible = False
  Sheets("BTSIKEPEER").Visible = False
  Sheets("BTSIPSECPROPOSAL").Visible = False
  Sheets("BTSIPSECPOLICY").Visible = False
  Sheets("BTSIPSECBIND").Visible = False
  Sheets("BTSIPSECDTNL").Visible = False
  
  Sheets("BTSACL").Visible = False
  Sheets("BTSACLRULE").Visible = False
  Sheets("BTSFLOODDEFEND").Visible = False
  Sheets("BTSPACKETFILTER").Visible = False
  Sheets("BTSIPGUARD").Visible = False
  Sheets("BTSCERTMK").Visible = False
  Sheets("BTSAPPCERT").Visible = False
  
  Sheets("BTSCERTREQ").Visible = False
  Sheets("BTSTRUSTCERT").Visible = False
  Sheets("BTSCRL").Visible = False
  Sheets("BTSCRLTSK").Visible = False
  Sheets("BTSCRLPOLICY").Visible = False
  Sheets("BTSCERTCHKTSK").Visible = False
  Sheets("BTSCA").Visible = False
  Sheets("BTSCERTDEPLOY").Visible = False
  
  'Sheets("BTSLNKBKATTR").Visible = False
  Sheets("BTSTUNNEL").Visible = False
  'Sheets("BTSIPBAK").Visible = False
  Sheets("BTSCLK").Visible = False
  Sheets("IPPOOL").Visible = False
  Sheets("IPPOOLMUX").Visible = False
  Sheets("IPPOOLIP").Visible = False
  Sheets("SRCIPRT").Visible = False
  Sheets("BTSESN").Visible = False
  Sheets("IPPOOLPM").Visible = False
  Sheets("ADJPPPBIND").Visible = False
  
    Sheets("ETHTRKIP").Visible = False
  
  Sheets("BTSEQUIPMENT").Visible = False
  Sheets("BTSLR").Visible = False
  Sheets("BTSGTRANSPARA").Visible = False
  Sheets("BTSIPLGCPORT").Visible = False
  Sheets("BTSIPPM").Visible = False
  
  Sheets("BTSABISMUXFLOW").Visible = False
  Sheets("BTSABISPRIMAP").Visible = False
  Sheets("BTSIPTOLGCPORT").Visible = False
  
Sheets("BTSSHARING").Visible = False
Sheets("BTSEXTOPIP").Visible = False
Sheets("BTSEXTOPMAP").Visible = False
Sheets("BTSEXTOPABISMUXFLOW").Visible = False
Sheets("BTSEXTOPIPPM").Visible = False
Sheets("BTSARPSESSION").Visible = False
Sheets("ALGCTRLPARA").Visible = False
End Sub

Public Sub ShowIPFEandE1T1(bShowIPFE As Boolean)
  If bShowIPFE Then
 ' ShowIPOE (True)
  ShowIPFE (True)
  Sheets("BTSIPBAK").Visible = True
  Sheets("BTSLNKBKATTR").Visible = True
  End If
End Sub
Public Sub ShowIPFE(bShowIPFE As Boolean)
  If bShowIPFE Then
    Sheets("ETHIP").Visible = True
    Sheets("BTS").Visible = True
    Sheets("ADJNODE").Visible = True
    Sheets("ADJNODEDIP").Visible = True
   ' Sheets("IPPATH").Visible = True
    Sheets("IPRT").Visible = True
    Sheets("BTSIP").Visible = True
    Sheets("BTSETHPORT").Visible = True
    Sheets("BTSIPCLKPARA").Visible = True
    Sheets("BTSBFD").Visible = True
    Sheets("BTSIPRT").Visible = True
    Sheets("BTSIPRTBIND").Visible = True
    Sheets("BTSDHCPSVRIP").Visible = True
    Sheets("BTSDEVIP").Visible = True
    'Sheets("RSCGRP").Visible = True
    Sheets("IPLOGICPORT").Visible = True
    Sheets("DEVIP").Visible = True
    Sheets("ADJMAP").Visible = True
    Sheets("BTSVLAN").Visible = True
    Sheets("BTSVLANCLASS").Visible = True
    Sheets("BTSVLANMAP").Visible = True
     Sheets("BTSCLK").Visible = True
    
  Sheets("BTSIKECFG").Visible = True
  Sheets("BTSIKEPROPOSAL").Visible = True
  Sheets("BTSIKEPEER").Visible = True
  Sheets("BTSIPSECPROPOSAL").Visible = True
  Sheets("BTSIPSECPOLICY").Visible = True
  Sheets("BTSIPSECBIND").Visible = True
  Sheets("BTSIPSECDTNL").Visible = True
  
  Sheets("BTSACL").Visible = True
  Sheets("BTSACLRULE").Visible = True
  Sheets("BTSFLOODDEFEND").Visible = True
  Sheets("BTSPACKETFILTER").Visible = True
  Sheets("BTSIPGUARD").Visible = True
  Sheets("BTSCERTMK").Visible = True
  Sheets("BTSAPPCERT").Visible = True
  
  Sheets("BTSCERTREQ").Visible = True
  Sheets("BTSTRUSTCERT").Visible = True
  Sheets("BTSCRL").Visible = True
  Sheets("BTSCRLTSK").Visible = True
  Sheets("BTSCRLPOLICY").Visible = True
  Sheets("BTSCERTCHKTSK").Visible = True
  Sheets("BTSCA").Visible = True
  Sheets("BTSCERTDEPLOY").Visible = True
    
     Sheets("BTSDSCPMAP").Visible = True
    'Sheets("BTSIPBAK").Visible = False
   ' Sheets("BTSLNKBKATTR").Visible = False
     Sheets("BTSTUNNEL").Visible = True
  Sheets("IPPOOL").Visible = True
  Sheets("IPPOOLMUX").Visible = True
  
  Sheets("IPPOOLIP").Visible = True
       Sheets("IPPOOLPM").Visible = True
  Sheets("SRCIPRT").Visible = True
    Sheets("BTSESN").Visible = True

  
  Sheets("BTSEQUIPMENT").Visible = True
  Sheets("BTSLR").Visible = True
  Sheets("BTSGTRANSPARA").Visible = True
  Sheets("BTSIPLGCPORT").Visible = True
  Sheets("BTSIPPM").Visible = True
  
  Sheets("BTSABISMUXFLOW").Visible = True
  Sheets("BTSABISPRIMAP").Visible = True
  Sheets("BTSIPTOLGCPORT").Visible = True

Sheets("BTSSHARING").Visible = True
Sheets("BTSEXTOPIP").Visible = True
Sheets("BTSEXTOPMAP").Visible = True
Sheets("BTSEXTOPABISMUXFLOW").Visible = True
Sheets("BTSEXTOPIPPM").Visible = True
Sheets("BTSARPSESSION").Visible = True
Sheets("ALGCTRLPARA").Visible = True
Sheets("BTSCONNECT").Visible = True
  End If
End Sub




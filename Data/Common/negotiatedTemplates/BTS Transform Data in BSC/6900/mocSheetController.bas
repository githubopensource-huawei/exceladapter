Attribute VB_Name = "mocSheetController"

Public Sub ShowBSCIPBrd(bBSCIPBrd As Boolean)
  If bBSCIPBrd Then
    Sheets("BRD").Visible = True
    Sheets("E1T1").Visible = True
    Sheets("COPTLNK").Visible = True
    Sheets("QUEUEMAP").Visible = True
    Sheets("CLK").Visible = True
    Sheets("MSP").Visible = True
    Sheets("OPT").Visible = True
    Sheets("ETHPORT").Visible = True
    Sheets("ETHREDPORT").Visible = True
    Sheets("BFDPROTOSW").Visible = True
    Sheets("PHBMAP").Visible = True
    Sheets("EFMAH").Visible = True
    Sheets("ETHTRK").Visible = True
    Sheets("ETHTRKLNK").Visible = True
    Sheets("ALGCTRLPARA").Visible = True
    
     Sheets("IPGUARD").Visible = True
    Sheets("PORTOSCCTRLPARA").Visible = True
  Else
    Sheets("BRD").Visible = False
    Sheets("E1T1").Visible = False
    Sheets("COPTLNK").Visible = False
    Sheets("QUEUEMAP").Visible = False
    Sheets("CLK").Visible = False
    Sheets("MSP").Visible = False
    Sheets("OPT").Visible = False
    Sheets("ETHPORT").Visible = False
    Sheets("ETHREDPORT").Visible = False
    Sheets("BFDPROTOSW").Visible = False
    Sheets("PHBMAP").Visible = False
    Sheets("EFMAH").Visible = False
    Sheets("ETHTRK").Visible = False
    Sheets("ETHTRKLNK").Visible = False

    Sheets("ALGCTRLPARA").Visible = False
    
     Sheets("IPGUARD").Visible = False
    Sheets("PORTOSCCTRLPARA").Visible = False
  End If
  
  Sheets("INTBRDPARA").Visible = bBSCIPBrd
End Sub

Public Sub ShowBSCIP(bBSCIPOE As Boolean, bBSCIPFE As Boolean)
  If bBSCIPOE And bBSCIPFE Then
    Sheets("MPGRP").Visible = True
    Sheets("MPLNK").Visible = True
    Sheets("ETHIP").Visible = True
    Sheets("PPPLNK").Visible = True
    Sheets("ADJNODE").Visible = True
    Sheets("IPPATH").Visible = True
    Sheets("IPRT").Visible = True
    Sheets("VLANID").Visible = True
    Sheets("RSCGRP").Visible = True
    Sheets("IPLOGICPORT").Visible = True
    Sheets("DEVIP").Visible = True
    Sheets("ADJMAP").Visible = True
    Sheets("ETHTRKIP").Visible = True
    
    Sheets("IPMUX").Visible = True
    Sheets("IPPM").Visible = True
    Sheets("IPPATHBIND").Visible = True
    Sheets("ALGCTRLPARA").Visible = True
    
    
  End If
  
  If bBSCIPOE And Not bBSCIPFE Then
    Sheets("MPGRP").Visible = True
    Sheets("MPLNK").Visible = True
    Sheets("PPPLNK").Visible = True
    Sheets("ADJNODE").Visible = True
    Sheets("IPPATH").Visible = True
    Sheets("IPRT").Visible = True
    Sheets("VLANID").Visible = True
    Sheets("RSCGRP").Visible = True
    Sheets("IPLOGICPORT").Visible = True
    Sheets("DEVIP").Visible = True
    Sheets("ADJMAP").Visible = True
    
    Sheets("ETHTRKIP").Visible = False
    Sheets("ETHIP").Visible = False
    Sheets("IPMUX").Visible = True
    Sheets("IPPM").Visible = True
    Sheets("IPPATHBIND").Visible = True
    Sheets("ALGCTRLPARA").Visible = True
  End If
  
  If bBSCIPFE And Not bBSCIPOE Then
    Sheets("ETHIP").Visible = True
    Sheets("ETHTRKIP").Visible = True
    Sheets("ADJNODE").Visible = True
    Sheets("IPPATH").Visible = True
    Sheets("IPRT").Visible = True
    Sheets("VLANID").Visible = True
    Sheets("RSCGRP").Visible = True
    Sheets("IPLOGICPORT").Visible = True
    Sheets("DEVIP").Visible = True
    Sheets("ADJMAP").Visible = True
    
    Sheets("MPGRP").Visible = False
    Sheets("MPLNK").Visible = False
    Sheets("PPPLNK").Visible = False
    Sheets("IPMUX").Visible = True
    Sheets("IPPM").Visible = True
    Sheets("IPPATHBIND").Visible = True
  End If
  
  If Not bBSCIPOE And Not bBSCIPFE Then
    Sheets("ETHIP").Visible = False
    Sheets("ADJNODE").Visible = False
    Sheets("IPPATH").Visible = False
    Sheets("IPRT").Visible = False
    Sheets("VLANID").Visible = False
    Sheets("RSCGRP").Visible = False
    Sheets("IPLOGICPORT").Visible = False
    Sheets("DEVIP").Visible = False
    Sheets("ADJMAP").Visible = False
    Sheets("ETHTRKIP").Visible = False
    Sheets("MPGRP").Visible = False
    Sheets("MPLNK").Visible = False
    Sheets("PPPLNK").Visible = False
    Sheets("IPMUX").Visible = False
    Sheets("IPPM").Visible = False
    Sheets("IPPATHBIND").Visible = False
    Sheets("ALGCTRLPARA").Visible = False
  End If
End Sub

Public Sub ShowBTSAttr(bBTSAttr As Boolean)
  If bBTSAttr Then
    Sheets("BTS").Visible = True
    Sheets("BTSBRD").Visible = True
    Sheets("BTSESN").Visible = True
    Sheets("BTSIP").Visible = True
    Sheets("BTSCLK").Visible = True
    Sheets("BTSIPCLKPARA").Visible = True
    Sheets("BTSBFD").Visible = True
    Sheets("BTSIPRT").Visible = True
    Sheets("BTSIPRTBIND").Visible = True
    Sheets("BTSDHCPSVRIP").Visible = True
    Sheets("BTSDEVIP").Visible = True
    Sheets("BTSTUNNEL").Visible = True
    Sheets("BTSLR").Visible = True
    Sheets("BTSEQUIPMENT").Visible = True
    Sheets("BTSGTRANSPARA").Visible = True
    Sheets("BTSCRL").Visible = True
    Sheets("BTSIPLGCPORT").Visible = True
    Sheets("BTSIPPM").Visible = True
    Sheets("BTSABISMUXFLOW").Visible = True
    Sheets("BTSABISPRIMAP").Visible = True
    Sheets("BTSIPTOLGCPORT").Visible = True
    
    Sheets("BTSSHARING").Visible = True

    
    
  Else
    Sheets("BTS").Visible = False
    Sheets("BTSBRD").Visible = False
    Sheets("BTSESN").Visible = False
    Sheets("BTSIP").Visible = False
    Sheets("BTSCLK").Visible = False
    Sheets("BTSIPCLKPARA").Visible = False
    Sheets("BTSBFD").Visible = False
    Sheets("BTSIPRT").Visible = False
    Sheets("BTSIPRTBIND").Visible = False
    Sheets("BTSDHCPSVRIP").Visible = False
    Sheets("BTSDEVIP").Visible = False
    Sheets("BTSTUNNEL").Visible = False
    Sheets("BTSLR").Visible = False
    Sheets("BTSEQUIPMENT").Visible = False
    Sheets("BTSGTRANSPARA").Visible = False
    Sheets("BTSCRL").Visible = False
    Sheets("BTSIPLGCPORT").Visible = False
    Sheets("BTSIPPM").Visible = False
    Sheets("BTSABISMUXFLOW").Visible = False
    Sheets("BTSABISPRIMAP").Visible = False
    Sheets("BTSIPTOLGCPORT").Visible = False
    Sheets("BTSSHARING").Visible = False
  End If
End Sub

Public Sub ShowIPOEIPFE(bIPOE As Boolean, bIPFE As Boolean, bIPFE_E1 As Boolean)
  If bIPFE_E1 Then
    bIPOE = True
    bIPFE = True
    Sheets("BTSIPBAK").Visible = True
    Sheets("BTSLNKBKATTR").Visible = True
  Else
    Sheets("BTSIPBAK").Visible = False
    Sheets("BTSLNKBKATTR").Visible = False
  End If
  
  If bIPOE Then
    Sheets("BTSPPPLNK").Visible = True
    Sheets("BTSMPGRP").Visible = True
    Sheets("BTSMPLNK").Visible = True
    Sheets("BTSCONNECT").Visible = True
    Sheets("BTSMONITORTS").Visible = True
    Sheets("BTSFORBIDTS").Visible = True
    Sheets("BTSETHPORT").Visible = True
    Sheets("BTSVLAN").Visible = True
    Sheets("BTSVLANCLASS").Visible = True
    Sheets("BTSVLANMAP").Visible = True
    Sheets("BTSDSCPMAP").Visible = True
    Sheets("BTSEXTOPIP").Visible = True
    Sheets("BTSEXTOPMAP").Visible = True
    Sheets("BTSEXTOPABISMUXFLOW").Visible = True
    Sheets("BTSEXTOPIPPM").Visible = True
    Sheets("BTSARPSESSION").Visible = True
  Else
    Sheets("BTSPPPLNK").Visible = False
    Sheets("BTSMPGRP").Visible = False
    Sheets("BTSMPLNK").Visible = False
    Sheets("BTSCONNECT").Visible = False
    Sheets("BTSMONITORTS").Visible = False
    Sheets("BTSFORBIDTS").Visible = False
  End If
  
  If bIPFE Then
    Sheets("BTSETHPORT").Visible = True
    Sheets("BTSVLAN").Visible = True
    Sheets("BTSVLANCLASS").Visible = True
    Sheets("BTSVLANMAP").Visible = True
    Sheets("BTSDSCPMAP").Visible = True
    Sheets("BTSEXTOPIP").Visible = True
    Sheets("BTSEXTOPMAP").Visible = True
    Sheets("BTSEXTOPABISMUXFLOW").Visible = True
    Sheets("BTSEXTOPIPPM").Visible = True
    Sheets("BTSARPSESSION").Visible = True
    Sheets("BTSCONNECT").Visible = True
  End If
  
  If Not (bIPOE Or bIPFE) Then
    Sheets("BTSETHPORT").Visible = (bIPOE Or bIPFE)
    Sheets("BTSVLAN").Visible = (bIPOE Or bIPFE)
    Sheets("BTSVLANCLASS").Visible = (bIPOE Or bIPFE)
    Sheets("BTSVLANMAP").Visible = (bIPOE Or bIPFE)
    Sheets("BTSDSCPMAP").Visible = (bIPOE Or bIPFE)
     Sheets("BTSEXTOPIP").Visible = False
    Sheets("BTSEXTOPMAP").Visible = False
    Sheets("BTSEXTOPABISMUXFLOW").Visible = False
    Sheets("BTSEXTOPIPPM").Visible = False
    Sheets("BTSARPSESSION").Visible = False
  End If

End Sub

Public Sub ShowBTSIPSec(bBTSIPSec As Boolean)
  If bBTSIPSec Then
    Sheets("BTSACL").Visible = True
    Sheets("BTSACLRULE").Visible = True
    Sheets("BTSIKECFG").Visible = True
    Sheets("BTSIKEPROPOSAL").Visible = True
    Sheets("BTSIKEPEER").Visible = True
    Sheets("BTSIPSECPROPOSAL").Visible = True
    Sheets("BTSIPSECPOLICY").Visible = True
    Sheets("BTSIPSECBIND").Visible = True
    Sheets("BTSIPSECDTNL").Visible = True
    Sheets("BTSCERTREQ").Visible = True
    Sheets("BTSCRLTSK").Visible = True
    Sheets("BTSCRLPOLICY").Visible = True
    Sheets("BTSCERTCHKTSK").Visible = True
    Sheets("BTSCA").Visible = True
    Sheets("BTSCERTDEPLOY").Visible = True
    Sheets("BTSPACKETFILTER").Visible = True
    Sheets("BTSFLOODDEFEND").Visible = True
    Sheets("BTSIPGUARD").Visible = True
    Sheets("BTSCERTMK").Visible = True
    Sheets("BTSAPPCERT").Visible = True
    Sheets("BTSTRUSTCERT").Visible = True
  Else
    Sheets("BTSACL").Visible = False
    Sheets("BTSACLRULE").Visible = False
    Sheets("BTSIKECFG").Visible = False
    Sheets("BTSIKEPROPOSAL").Visible = False
    Sheets("BTSIKEPEER").Visible = False
    Sheets("BTSIPSECPROPOSAL").Visible = False
    Sheets("BTSIPSECPOLICY").Visible = False
    Sheets("BTSIPSECBIND").Visible = False
    Sheets("BTSIPSECDTNL").Visible = False
    Sheets("BTSCERTREQ").Visible = False
    Sheets("BTSCRLTSK").Visible = False
    Sheets("BTSCRLPOLICY").Visible = False
    Sheets("BTSCERTCHKTSK").Visible = False
    Sheets("BTSCA").Visible = False
    Sheets("BTSCERTDEPLOY").Visible = False
    Sheets("BTSPACKETFILTER").Visible = False
    Sheets("BTSFLOODDEFEND").Visible = False
    Sheets("BTSIPGUARD").Visible = False
    Sheets("BTSCERTMK").Visible = False
    Sheets("BTSAPPCERT").Visible = False
    Sheets("BTSTRUSTCERT").Visible = False
  End If
  
End Sub





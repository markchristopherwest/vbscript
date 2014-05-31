

'Set Screen Resolution to 800x600
'VMware Admnistrators want Same Size




'Set Hardware Resolution to Full
'VMware likes Full Hardware Resolution
Const HKEY_LOCAL_MACHINE = &amp;H80000002
Dim RegValueData
Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
 
'Get Path to Acceleration.Level
objReg.GetStringValue HKEY_LOCAL_MACHINE,"HARDWARE\DEVICEMAP\VIDEO","\Device\Video0",RegValueData
 
'Set Acceleration.Level to Full
objReg.SetDWORDValue HKEY_LOCAL_MACHINE,RIGHT(RegValueData, 82),"Acceleration.Level",0 
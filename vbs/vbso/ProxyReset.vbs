set wshshell = createobject("Wscript.shell")
WSHShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 0, "REG_DWORD"
wshshell.popup "Proxy deaktiviert"

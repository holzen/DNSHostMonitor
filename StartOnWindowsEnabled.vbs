Dim keyInicio
Set keyInicio = WScript.CreateObject("WScript.Shell")
keyInicio.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\DNSMonitorInit", "C:\DNSHostMonitor\DNSRun.bat", "REG_SZ"


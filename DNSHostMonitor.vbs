'				  DNS Host Monitor
'	 		Holzen Atocha Martinez Garcia 
'			   Versión 1.0 August 2009
'	   		   Versión 1.1 January 2020

MsgBox("DNSHostMonitor for Windows by Holzen Martinez")
'Verify the host file and path
host = "C:\\WINDOWS\\System32\\Drivers\\etc\\hosts"
set monitoredFile = CreateObject("Scripting.FileSystemObject")
'if file exists, then start monitoring'
if (monitoredFile.FileExists(host)) Then
	MsgBox("Initializing system's hosts monitoring")
	Monitoring()
else 
	MsgBox("Hosts File don't exists")
	MsgBox("Making File Hosts Default")
 	'MakeHostsDefault(host)
 	Set objWshell = Wscript.CreateObject("Wscript.Shell")
	objWshell.Run "runas /user:administrator C:\\DNSHostMonitor\\makeHosts.bat"
 	MsgBox("Default hosts file has been created, initializing system's hosts monitor")
	Monitoring()
end if

Sub Monitoring()
	'Monitorize the Hosts File
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
   	 & "{impersonationLevel=impersonate}!\\" & _
        	strComputer & "\root\cimv2")
	Set colMonitoredEvents = objWMIService.ExecNotificationQuery _
	    ("SELECT * FROM __InstanceModificationEvent WITHIN 10 WHERE " _
	        & "TargetInstance ISA 'CIM_DataFile' and " _
	            & "TargetInstance.Name='C:\\WINDOWS\\System32\\Drivers\\etc\\hosts'")
	Do
   			Set objLatestEvent = colMonitoredEvents.NextEvent
    		if objLatestEvent.PreviousInstance.FileSize<>objLatestEvent.TargetInstance.FileSize Then
				MsgBox("Hosts File has been modified")
				Wscript.Echo "Previous Size: " & objLatestEvent.PreviousInstance.FileSize & " bytes"
    			Wscript.Echo "Current Size: " & objLatestEvent.TargetInstance.FileSize	& " bytes"
				MsgBox("The system could be a victim of Pharming or a similar attack, if you are modifying the host file, ignore this warning, otherwise check your hosts file and logs or contact your Network Administrator.")
				Dim p 
				Dim c 
				p = objLatestEvent.PreviousInstance.FileSize
				c = objLatestEvent.TargetInstance.FileSize
				LogRegister p,c
			end if
	Loop
End Sub

Sub LogRegister(info1, info2)
	Dim object, path, nameFile 
	path="C:\\DNSHostMonitor\\"
	nameFile ="logs.txt"
	Set object = CreateObject("Scripting.FileSystemObject")
	if (object.FileExists(path & nameFile)) Then
		Set logFile = object.OpenTextFile(path & nameFile, 8, False, 0)
		logFile.WriteLine("------------------------")
		logFile.WriteLine("  Change on Hosts File")
		logFile.WriteLine("Previous Size: " & info1 & " bytes")
		logFile.WriteLine("Actual Size: " & info2 & " bytes")
		logFile.WriteLine("Modification in: " & Now())
		logFile.WriteLine("------------------------")
		logFile.WriteBlankLines(1)
		logFile.Close()
	else
		Set logFile = object.CreateTextFile(path & nameFile, True)
		logFile.WriteLine("------------------------")
		logFile.WriteLine("  Change on Hosts File")
		logFile.WriteLine("Previous Size: " & info1 & " bytes")
		logFile.WriteLine("Actual Size: " & info2 & " bytes")
		logFile.WriteLine("Modification in: " & Now())
		logFile.WriteLine("------------------------")
		logFile.WriteBlankLines(1)
		logFile.Close()
	end if
End Sub
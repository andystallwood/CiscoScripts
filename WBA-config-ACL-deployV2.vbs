# $language = "VBScript"
# $interface = "1.0"
Dim FSO, Shell, Windir, Runservice, oFile, oFile1, test, deployed
dim tFACL, cFPRoute, cFSRoute, tFAuth, cFPAuth, cFSAuth, tFail, tFRolled, tRolled, tUnknownT
'Dim g_clipboard, g_map
'Dim re, Matches, Match
'Dim nResult, nRow
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Set FSO = CreateObject("scripting.filesystemobject")
Set Shell = CreateObject("WScript.Shell")
Set objSc = crt.Screen
Set objD = crt.Dialog
Set objSe = crt.Session
Set objW = crt.Window
HOSTIP = objD.Prompt("Enter folder name and Path to the hosts file","Folder Name & Path","h:\script\PHOTO-ACL\Batch#\ACL-IP.txt")
Set SwitchIP = FSO.opentextfile(HOSTIP, ForReading, False)
Set objDictionary = CreateObject("Scripting.Dictionary")


User = objD.Prompt("Enter YOUR Username To Get into device"&Chr(13)&Chr(13)&_
"Same username used for all devices"," ","xxxxxxxxx")
Pass = objD.Prompt("Enter password To Get into device"&Chr(13)&Chr(13)&_
"Password must be the same for all device!"," ","xxx", TRUE)
Logfiles = objD.Prompt("Enter folder name and Path to save Log files In.","Folder Name & Path","H:\script\logs\Photo-ACL\Batch#")

missAclPhoto = 0


If FSO.FolderExists(Logfiles) Then
Else
	FSO.CreateFolder(Logfiles)
End IF
'<---- Adds a deployment start Time stamp to the summary log file.
Set Tempfile = FSO.OpenTextFile(Logfiles&"\Summary.txt",ForAppending, True)
tempfile.writeline "Deployment start - " & Now ()
tempfile.writeline "--------------"
TempFile.Close()



'<-------- Script Purpose: Deploy
'<-- 1. Update the 
'<--------

'Start loop
While Not SwitchIP.atEndOfStream
	IP = SwitchIP.Readline()
	save = 0
	'authFail = 0
	counter = counter + 1
	
	'<-------------------- Device Connect sequence ---------------------------------> 
	On Error Resume Next
	On Error Resume Next
	crt.session.connect "/SSH2 /AcceptHostKeys /L " & User & " /PASSWORD " & Pass & " " & IP & " "
	On Error Goto 0
	If ( crt.Session.Connected ) Then
		objsc.Synchronous = True
		'<-------------------- Create logfile for command/logfile changes
		objse.logfilename = Logfiles&"\%Y%M%D-%h%m-"&IP&".cfg"
		objse.Log(True) '<-- This opens the logfile and captures the SecureCRT output to the timestamped ".cfg" file
		enablepass = objSc.waitforstrings(">", "#")
		objSc.Send "enable" & vbCr
		objSc.WaitForString":"
		objSc.Send Pass & vbCr
		objSc.WaitForString"#"
		objSc.Send "term length 0" & vbCr '<-- allows
		objSc.WaitForString"#"
		
		'<-------------------- END if Connect sequence ------------------------------->
		'<-------------------- Pre-test section -------------------------------------->
		
		'<---- PHOTO-ACL Section
		objSc.Send"sh ip access-list PHOTO-ACL" & VbCr 
		aclPhoto = objSc.WaitForString("Extended",5)
		if aclPhoto = TRUE then
			'
			'<-------------------- Native command list ----------------------------------->
			'conf t
			'no ip access-list extended PHOTO-ACL
			'ip access-list extended PHOTO-ACL
			'remark Hello World
			'<-------------------- Scripted command list --------------------------------->
			'<-- "!" Notes for the config output - Ignored by device
			objSc.Send "conf t" & VbCr 
			objSc.WaitForString "(config)#" '<----------------------Command Check
			objSc.Send "no ip access-list extended PHOTO-ACL" & VbCr 
			objSc.WaitForString "(config)#" '<----------------------Command Check
			objSc.Send "ip access-list extended PHOTO-ACL" & VbCr 
			objSc.WaitForString "(config-ext-nacl)#" '<----------------------Command Check
			objSc.Send "remark Hello World" & VbCr 
			objSc.WaitForString "(config-ext-nacl)#" '<----------------------Command Check
			objSc.Send "end" & VbCr
			objSc.WaitForString"#"
			save = save + 1
		else
			Set Tempfiledata = FSO.OpenTextFile(Logfiles&"\missing-ACL-Photo.txt",ForAppending, True)
			TempFiledata.writeline IP
			TempFiledata.Close()
			missAclPhoto = missAclPhoto + 1
		end if
		
		if save > 0 then
			objSc.Send "copy run start" & VbCr
			objSc.WaitForString"]?"
			objSc.Send VbCr
			objSc.WaitForString"#"
			deployed = deployed + 1
			Set Tempfiledata = FSO.OpenTextFile(Logfiles&"\ACL-Updated-list.txt",ForAppending, True)
			TempFiledata.writeline IP
			TempFiledata.Close()
		end if
		
		objSc.Send "exit" & vbCr
		objse.Log(False)
		objSc.Synchronous = False
		objSE.Disconnect
		
	Else
		Set Tempfile = FSO.OpenTextFile(Logfiles&"\NoConnect.txt",ForAppending, True)
		TempFile.writeline IP
		TempFile.Close()
	End IF

Wend

tFail = missAclPhoto
tRolled = ""
tFRolled= ""

Set Tempfile = FSO.OpenTextFile(Logfiles&"\Summary.txt",ForAppending, True)
TempFile.writeline "Deployment Complete: " & Now ()
tempfile.writeline "--------------"
TempFile.writeline "Total Number of devices: " & counter
TempFile.writeline "Total Number of Updated: " & deployed
TempFile.writeline "Total Number of warnings: " & tFail
TempFile.writeline "Total Number of Rolled Back: N/A" & tRolled
TempFile.writeline "Rolled Back failed: N/A" & tFRolled
tempfile.writeline "--------------"
tempfile.writeline "Missing PHOTO-ACL: " & missAclPhoto
tempfile.writeline "--------------"
TempFile.Close()




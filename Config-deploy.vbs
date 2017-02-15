# $language = "VBScript"
# $interface = "1.0"
Dim FSO, Shell, deployed
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Set FSO = CreateObject("scripting.filesystemobject")
Set Shell = CreateObject("WScript.Shell")
Set objSc = crt.Screen
Set objD = crt.Dialog
Set objSe = crt.Session
Set objW = crt.Window
Set objDictionary = CreateObject("Scripting.Dictionary")

'File containing a list of Cisco Device IPs to perform the change on. One IP per line.
HOSTIP = objD.Prompt("Enter folder name and Path to the hosts file","Folder Name & Path","U:\script\PHOTO-ACL\ACL-IP.txt")
Set SwitchIP = FSO.opentextfile(HOSTIP, ForReading, False)

'File containing a list of commands to perform on each router. One command per line.
CommandsFile = objD.Prompt("Enter folder name and Path to the commands file","Folder Name & Path","U:\script\PHOTO-ACL\Commands.csv")

User = objD.Prompt("Enter YOUR Username To Get into device"&Chr(13)&Chr(13)&_
"Same username used for all devices"," ","xxxxxxxxxx")

Pass = objD.Prompt("Enter password To Get into device"&Chr(13)&Chr(13)&_
"Password must be the same for all devices!"," ","xxx", TRUE)

Logfiles = objD.Prompt("Enter folder name and Path to save Log files In.","Folder Name & Path","U:\script\PHOTO-ACL\logs\")

ErrorCount = 0

If FSO.FolderExists(Logfiles) Then
Else
	FSO.CreateFolder(Logfiles)
End IF
'<---- Adds a deployment start Time stamp to the summary log file.
Set Tempfile = FSO.OpenTextFile(Logfiles&"\Summary.txt",ForAppending, True)
DeployStart = Now ()
tempfile.writeline "Deployment start - " & DeployStart
tempfile.writeline "--------------"
TempFile.Close()

'<-------- Script Purpose: Deploy
'<-- Deploy Config on a set of Cisco devices
'<--------

'Start loop
While Not SwitchIP.atEndOfStream
	IP = SwitchIP.Readline()
	Set Config = FSO.opentextfile(CommandsFile, ForReading, False)
	save = 0
	counter = counter + 1
	
	'<-------------------- Device Connect sequence ---------------------------------> 
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
		objSc.Send "term length 0" & vbCr '<-- disables paging of screen output
		objSc.WaitForString"#"
		'<-------------------- END if Connect sequence ------------------------------->
		
			'<---- Read each config line in turn from the CSV file and send to the device
			ConfigLine = Config.ReadLine 'Read Header Row of CSV file
			
		Do While Not Config.atEndOfStream
			'Split Line Read into Category, Command, Prompt and Output and process the line
			ConfigLine = Split(Config.ReadLine,",")
			Category = ConfigLine(0)				
			Command = ConfigLine(1)
			Prompt = ConfigLine(2)
			Output = ConfigLine(3)
			WarnOrFail = ConfigLine(4)

			ProcessCommand = ProcessLine(Category, Command, Prompt, Output, Logfiles, WarnOrFail)
			If ProcessCommand  = 1 Then '<----Warning
				ErrorCount = ErrorCount + 1
			ElseIf ProcessCommand  = 2 Then '<----Failure
				ErrorCount = ErrorCount + 1
				Exit Do
			ElseIf ProcessCommand  = 3 Then '<----Input File Failure
				ErrorCount = ErrorCount + 1
				Exit Do
			Else '<----Success
			End If
		Loop
		Config.Close() 'Close Config File
		
		objSc.Send "end" & VbCr
		objSc.WaitForString"#"
		save = save + 1
					
		if save > 0 then
			deployed = deployed + SaveConfig(Logfiles)
		end if
		
		objSc.Send "exit" & vbCr
		objse.Log(False)
		objSc.Synchronous = False
		objSE.Disconnect
		
	Else
		Set NoConnectfile = FSO.OpenTextFile(Logfiles&"\NoConnect.txt",ForAppending, True)
		NoConnectfile.writeline IP
		NoConnectfile.Close()
	End IF

Wend

		SwitchIP.Close() 'Close Switch IP file File

Set Summaryfile = FSO.OpenTextFile(Logfiles&"\Summary.txt",ForAppending, True)
Summaryfile.writeline "Deployment Started: " & DeployStart
Summaryfile.writeline "Deployment Complete: " & Now ()
Summaryfile.writeline "--------------"
Summaryfile.writeline "Total Number of devices: " & counter
Summaryfile.writeline "Total Number of Updated: " & deployed
Summaryfile.writeline "Total Number of warnings: " & ErrorCount
Summaryfile.writeline "--------------"
Summaryfile.Close()

Function ProcessLine (Category, Command, Prompt, Output, Logfiles, WarnOrFail)
	if StrComp(Category,"config") = 0 then
		objSc.Send Command & VbCr
		PromptExpected = "(" & Prompt & ")#"
		objSc.WaitForString PromptExpected '<----------------------Prompt Check
		ProcessLine = 0 '<----Success
	elseif StrComp(Category,"test") = 0 then
		objSc.Send Command & VbCr
		TestSuccess = objSc.WaitForString(Output,5)
		if TestSuccess = FALSE And (StrComp(WarnOrFail,"warn") = 0) then 'Output not found, and a warning
			Set Tempfiledata = FSO.OpenTextFile(Logfiles&"\Errors.txt",ForAppending, True)
			TempFiledata.writeline IP & " Warning at " & Now() & " . Deployment Batch Started at " & DeployStart
			TempFiledata.Close()
			ProcessLine = 1 '<----Warning
		elseif TestSuccess = FALSE And (StrComp(WarnOrFail,"fail") = 0) then 'Output not found, and a failure
			Set Tempfiledata = FSO.OpenTextFile(Logfiles&"\Errors.txt",ForAppending, True)
			TempFiledata.writeline IP & " Failure. Exiting Device at " & Now() & " . Deployment Batch Started at " & DeployStart
			TempFiledata.Close()
			ProcessLine = 2 '<----Failure
		elseif TestSuccess = FALSE then
			Set Tempfiledata = FSO.OpenTextFile(Logfiles&"\Errors.txt",ForAppending, True)
			TempFiledata.writeline IP & " Command Check Failed. Exiting Device. Possible Error in Input File at " & Now() & " . Deployment Batch Started at " & DeployStart
			TempFiledata.Close()
			ProcessLine = 3 '<----Something has gone wrong with the input file
		else
			ProcessLine = 0 '<----Success
		end if
	end if
End Function

Function SaveConfig(Logfiles)
	objSc.Send "copy run start" & VbCr
	objSc.WaitForString"[startup-config]?"
	objSc.Send VbCr
	objSc.WaitForString"#"
	SaveConfig = 1
	Set Tempfiledata = FSO.OpenTextFile(Logfiles&"\Completed.txt",ForAppending, True)
	TempFiledata.writeline IP & " Deployment Started at " & DeployStart
	TempFiledata.Close()
End Function

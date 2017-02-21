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
HOSTIP = objD.Prompt("Enter folder name and Path to the hosts file","Folder Name & Path","U:\script\CreateVLAN\Hosts.csv")
Set DeviceIP = FSO.opentextfile(HOSTIP, ForReading, False)

'File containing a list of commands to perform on each router. One command per line.
CommandsFile = objD.Prompt("Enter folder name and Path to the commands file","Folder Name & Path","U:\script\CreateVLAN\Commands.csv")

User = objD.Prompt("Enter YOUR Username To Get into device"&Chr(13)&Chr(13)&_
"Same username used for all devices"," ","xxxxxxxxxx")

Pass = objD.Prompt("Enter password To Get into device"&Chr(13)&Chr(13)&_
"Password must be the same for all devices!"," ","xxx", TRUE)

Logfiles = objD.Prompt("Enter folder name and Path to save Log files In.","Folder Name & Path","U:\script\CreateVLAN\logs\")

ErrorCount = 0

If FSO.FolderExists(Logfiles) Then
Else
	FSO.CreateFolder(Logfiles)
End IF

DeployStart = Now () '<---- Used for a deployment start Time stamp.

'<-------- Script Purpose: Deploy
'<-- Deploy Config on a set of Cisco devices
'<--------

'Start loop
DeviceLine = DeviceIP.ReadLine 'Read Header Row of CSV file and ignore
While Not DeviceIP.atEndOfStream
	DeviceLine = Split(DeviceIP.Readline,",")
	IP = DeviceLine(0)
	SiteNumber = DeviceLine(1)
	Param1 = DeviceLine(2)
	Param2 = DeviceLine(3)
	Param3 = DeviceLine(4)
	Param4 = DeviceLine(5)
	Param5 = DeviceLine(6)
	Param6 = DeviceLine(7)
	Param7 = DeviceLine(8)
	Param8 = DeviceLine(9)
	Param9 = DeviceLine(10)
	Param10 = DeviceLine(11)
	
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
			ConfigLine = Config.ReadLine 'Read Header Row of CSV file and ignore
			
		Do While Not Config.atEndOfStream
			'Split Line Read into Category, Command, Prompt and Output and process the line
			ConfigLine = Split(Config.ReadLine,",")
			Category = ConfigLine(0)
			CommandStart = ConfigLine(1)
			Param = ConfigLine(2)
			CommandEnd = ConfigLine(3)
			Prompt = ConfigLine(4)
			Output = ConfigLine(5)
			WarnOrFail = ConfigLine(6)

			ProcessCommand = ProcessLine(Category, CommandStart, Prompt, Output, Logfiles, WarnOrFail)
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

DeviceIP.Close() 'Close Switch IP file File

Set Summaryfile = FSO.OpenTextFile(Logfiles&"\Summary.txt",ForAppending, True)
Summaryfile.writeline "Deployment Started: " & DeployStart
Summaryfile.writeline "Deployment Complete: " & Now ()
Summaryfile.writeline "--------------"
Summaryfile.writeline "Total Number of devices: " & counter
Summaryfile.writeline "Total Number of Updated: " & deployed
Summaryfile.writeline "Total Number of warnings: " & ErrorCount
Summaryfile.writeline "--------------"
Summaryfile.Close()

Function ProcessLine (Category, CommandStart, Prompt, Output, Logfiles, WarnOrFail)
	if StrComp(Category,"config") = 0 then
		objSc.Send CommandStart & VbCr
		PromptExpected = "(" & Prompt & ")#"
		objSc.WaitForString PromptExpected '<----------------------Prompt Check
		ProcessLine = 0 '<----Success
	elseif StrComp(Category,"test") = 0 then
		objSc.Send CommandStart & VbCr
		TestSuccess = objSc.WaitForString(Output,5)
		Set ErrorFile = FSO.OpenTextFile(Logfiles&"\Errors.txt",ForAppending, True) 'Open error File ready to be written to
		if TestSuccess = FALSE And (StrComp(WarnOrFail,"warn") = 0) then 'Output not found, and a warning
			ErrorFile.writeline IP & " Warning at " & Now() & " . Deployment Batch Started at " & DeployStart
			ProcessLine = 1 '<----Warning
		elseif TestSuccess = FALSE And (StrComp(WarnOrFail,"fail") = 0) then 'Output not found, and a failure
			ErrorFile.writeline IP & " Failure. Exiting Device at " & Now() & " . Deployment Batch Started at " & DeployStart
			ProcessLine = 2 '<----Failure
		elseif TestSuccess = FALSE then
			ErrorFile.writeline IP & " Command Check Failed. Exiting Device. Possible Error in Input File at " & Now() & " . Deployment Batch Started at " & DeployStart
			ProcessLine = 3 '<----Something has gone wrong with the input file
		else
			ProcessLine = 0 '<----Success
		end if
		ErrorFile.Close()
	end if
End Function

Function SaveConfig(Logfiles)
	objSc.Send "copy run start" & VbCr
	objSc.WaitForString"[startup-config]?"
	objSc.Send VbCr
	objSc.WaitForString"#"
	SaveConfig = 1
	Set CompletedFile = FSO.OpenTextFile(Logfiles&"\Completed.txt",ForAppending, True)
	CompletedFile.writeline IP & Now() & " . Deployment Batch Started at " & DeployStart
	CompletedFile.Close()
End Function
															

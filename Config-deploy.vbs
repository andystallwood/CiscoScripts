# $language = "VBScript"
# $interface = "1.0"
'Version 28 2019_06_25 - Removed case sensitivity for password prompt
Dim FSO, Shell
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

TotalDevices = 0
ConnectedCount = 0
WarnCount = 0
FailCount = 0
NoConnectCount = 0
SuccessCount = 0

'Input all parameters as a comma seperated string
AdvancedMode = MsgBox ("Enter all parameters as one string?" & vbCrLf & "If you don't know what this is, select No :-)" , vbYesNo, "Advanced Mode?")

If AdvancedMode = vbYes Then
	ParamString = objD.Prompt("Enter comma seperated string for input and output files", _
		"String", _
		"U:\script\Hosts.csv,U:\script\Commands.csv,U:\script\logs\,Username,Password")
	ParamInputs = Split(ParamString,",")
	HostFile = ParamInputs(0)
	CommandsFile = ParamInputs(1)
	Logfiles =  ParamInputs(2)
	User = ParamInputs(3)
	Pass =  ParamInputs(4)
	
	'Check file for invalid characters
	CheckInputFiles(HostFile)
	CheckInputFiles(CommandsFile)

Else
	'Directory for input and log files
	Directory = objD.Prompt("Enter Folder Path for the input and output files","Folder Path","U:\script\")
	
	HostFile = Directory & "Hosts.csv"
	CommandsFile = Directory & "Commands.csv"
	Logfiles = Directory & "logs\"

	'File containing a list of Cisco Device IPs to perform the change on. One IP per line.
	HostFile = objD.Prompt("Enter filename and Path to the hosts file","Hosts File Name & Path",HostFile)

	'Check file for invalid characters
	CheckInputFiles(HostFile)

	'File containing a list of commands to perform on each router. One command per line.
	CommandsFile = objD.Prompt("Enter filename and Path to the commands file","Command File Name & Path",CommandsFile)

	'Check file for invalid characters
	CheckInputFiles(CommandsFile)

	'Folder to recieve the log files
	Logfiles = objD.Prompt("Enter folder Path to save Log files In.","Log Folder Path",Logfiles)
	
	User = objD.Prompt("Enter username to get into devices","Username","xxxxxxxxxx")
	
	Pass = objD.Prompt("Enter password to get into devices","Password","xxxxxxxx", TRUE)
End If

If FSO.FolderExists(Logfiles) Then
Else
	FSO.CreateFolder(Logfiles)
End IF

DeployStart = Now () '<---- Used for a deployment start Time stamp.

Set ErrorFile = FSO.OpenTextFile(Logfiles&"\Errors.txt",ForAppending, True) 'Open error File ready to be written to

'<-------- Script Purpose: Deploy
'<-- Deploy Config on a set of Cisco devices
'<--------
Set Hosts = FSO.opentextfile(HostFile, ForReading, False)

DeviceLine = Hosts.ReadLine 'Read Header Row of CSV file and ignore

'Start loop for each device
While Not Hosts.atEndOfStream
	DeviceLine = Split(Hosts.Readline,",")
	
	IP = DeviceLine(0)
	HostName = DeviceLine(1)
	
	Set Commands = FSO.opentextfile(CommandsFile, ForReading, False)

	TotalDevices = TotalDevices + 1
	
	'<-------------------- Device Connect sequence ---------------------------------> 
	On Error Resume Next
	crt.session.connect "/SSH2 /AcceptHostKeys /L " & User & " /PASSWORD " & Pass & " " & IP & " "
		
	On Error Goto 0
	If ( crt.Session.Connected ) Then
		objsc.Synchronous = True
		'<-------------------- Create logfile for command/logfile changes
		objse.logfilename = Logfiles & "\" & IP & "-%Y%M%D-%h%m.cfg"
		objse.Log(True) '<-- This opens the logfile and captures the SecureCRT output to the timestamped ".cfg" file
		Success = TRUE
		IP = LEFT(IP & "        ", 15) 'Pad the IP with spaces on the right so the text files format nicely
		CorrectHost = objSc.WaitForString(HostName &">",5) 'Check for correct Prompt to be returned
		If CorrectHost = TRUE then 'Hostname matches IP address
			objSc.Send "enable" & vbCr
			objSc.WaitForStrings"Password:","password:"
			objSc.Send Pass & vbCr
			objSc.WaitForString"#"
			objSc.Send "term length 0" & vbCr '<-- disables paging of screen output
			objSc.WaitForString"#"
			'<-------------------- END of Connect sequence ------------------------------->

			ConfigLine = Commands.ReadLine 'Read Header Row of Commands CSV file and ignore
	
			'<-------------------- Device Configuration sequence --------------------------------->

			Do While Not Commands.atEndOfStream '<---- Read each Commands line in turn from the CSV file and send to the device
				'Split Line Read into Category, Command, Prompt and Output and process the line
				ConfigLine = Split(Commands.ReadLine,",")

				ProcessCommand = ProcessLine(ConfigLine, Logfiles, DeviceLine)
			
				If ProcessCommand  = 1 Then '<----Warning
					WarnCount = WarnCount + 1
					Success = FALSE
				ElseIf ProcessCommand  = 2 Then '<----Failure
					FailCount = FailCount + 1
					Success = FALSE
					Exit Do
				ElseIf ProcessCommand  = 3 Then '<----Input File Failure
					FailCount = FailCount + 1
					Success = FALSE
					Exit Do
				Else '<----Success
				End If
			Loop
			'<-------------------- END of Device Configuration sequence ---------------------------------> 	
			Success = TRUE
			Commands.Close() 'Close Config File
			objSc.Send "end" & VbCr
			objSc.WaitForString"#"
			ConnectedCount = ConnectedCount + SaveConfig(Logfiles, IP, HostName) ' Save Config and update Updated Count if succesful save.
			
		Else ' Hostname does not match IP address
			Success = FALSE
			ErrorFile.writeline IP &  " " & HostName & " Hostname doesn't match IP address " & Now() & ". Deployment Batch Started at " & DeployStart	
			ConnectedCount = ConnectedCount + 1
			FailCount = FailCount + 1
		End If
		
		If Success = TRUE Then
			SuccessCount = SuccessCount + 1
		End If
		
		objSc.Send "exit" & vbCr
		objse.Log(False)
		objSc.Synchronous = False
		objSE.Disconnect
		
	Else 'Device failed to connect
		ErrorFile.writeline IP &  " " & HostName & " Failed to connect " & Now() & ". Deployment Batch Started at " & DeployStart	
		NoConnectCount = NoConnectCount + 1
	End IF

Wend

'Close files
Hosts.Close()
ErrorFile.Close()

'Write Summary
Set Summaryfile = FSO.OpenTextFile(Logfiles&"\Summary.txt",ForAppending, True)
Summaryfile.writeline "------------------------------"
Summaryfile.writeline "Deployment Started        : " & DeployStart
Summaryfile.writeline "Deployment Complete       : " & Now ()
Summaryfile.writeline "------------------------------"
Summaryfile.writeline "Total Number of Devices   : " & TotalDevices
Summaryfile.writeline "Number of Connect Failures: " & NoConnectCount
Summaryfile.writeline "Number Connected          : " & Updated
Summaryfile.writeline "------------------------------"
Summaryfile.writeline "Number Successful         : " & SuccessCount
Summaryfile.writeline "Number of warnings        : " & WarnCount
Summaryfile.writeline "Number of failures        : " & FailCount
Summaryfile.writeline "-------------------------------"
Summaryfile.writeline "Hostfile Used             : " & HostFile
Summaryfile.writeline "Command File Used         : " & CommandsFile
Summaryfile.writeline "-------------------------------"
Summaryfile.writeline ""
Summaryfile.writeline ""
Summaryfile.Close()

MsgBox _
"Total Number of Devices: " & TotalDevices & vbCrLf & _
"Number of Connect Fails: " & NoConnectCount & vbCrLf & _
"Warnings: " & WarnCount & vbCrLf & _
"Failures: " & FailCount, _
vbOKOnly+vbInformation, _
"Script Completed"

'----------------------------------------------------------------------------------------------------------------------------
'Name       : ProcessLine -> Processes a line of the commands file.
'Parameters : ConfigLine  -> Comma seperated string with the config parameters present.
'           : Logfiles    -> Directory containing the log files.
'           : DeviceLine  -> Comma seperated string containing information about the device.
'Return     : SaveConfig  -> Returns 1 if the config was saved succesfully.
'----------------------------------------------------------------------------------------------------------------------------

Function ProcessLine (ConfigLine, Logfiles, DeviceLine)
	'Split the DeviceLine into the various elements
	HostName = DeviceLine(1)
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

	'Split the ConfigLine into the various elements
	Category = ConfigLine(0)
	CommandStart = ConfigLine(1)
	Param = ConfigLine(2)
	CommandEnd = ConfigLine(3)
	PromptExpected = ConfigLine(4)
	Output = ConfigLine(5)
	WarnOrFail = ConfigLine(6)

	Select Case Param
	Case "param1"  : Parameter = Param1
	Case "param2"  : Parameter = Param2
	Case "param3"  : Parameter = Param3
	Case "param4"  : Parameter = Param4
	Case "param5"  : Parameter = Param5
	Case "param6"  : Parameter = Param6
	Case "param7"  : Parameter = Param7
	Case "param8"  : Parameter = Param8
	Case "param9"  : Parameter = Param9
	Case "param10" : Parameter = Param10
	Case Else      : Parameter = ""
	End Select
	
	objSc.Send CommandStart & Parameter & CommandEnd & VbCr 'Send Command to Device

	if Category = "config" then 'Configuration Command
		ConfigSuccess = objSc.WaitForString(PromptExpected,5) 'Check for correct Prompt to be returned
		If ConfigSuccess = TRUE then
			ProcessLine = 0 'Success
		Else
			ErrorFile.writeline IP &  " " & HostName & " Config Command Failure. Exiting Device at " & Now() & ". Deployment Batch Started at " & DeployStart			
			ProcessLine = 2 'Failure
		End If
		
	elseif Category = "test" then
		TestSuccess = objSc.WaitForString(Output,5)
		
		if TestSuccess = FALSE And WarnOrFail = "warn" then 'Output not found, and a warning
			ErrorFile.writeline IP & " " & HostName & " Test Warning for " & CommandStart & Parameter & CommandEnd & " at " & Now() & ". Deployment Batch Started at " & DeployStart
			ProcessLine = 1 'Warning
			
		elseif TestSuccess = FALSE And WarnOrFail = "fail" then 'Output not found, and a failure
			ErrorFile.writeline IP &  " " & HostName & " Test Failure for " & CommandStart & Parameter & CommandEnd & ". Exiting Device at " & Now() & ". Deployment Batch Started at " & DeployStart
			ProcessLine = 2 'Failure
			
		elseif TestSuccess = FALSE then
			ErrorFile.writeline IP & " " & HostName & " Command Check failed for " & CommandStart & Parameter & CommandEnd & ". Exiting Device at " & Now() &". Error in Input File (WarnOrFail needs to be warn or fail). Deployment Batch Started at " & DeployStart
			ProcessLine = 3 'Something has gone wrong with the input file
			
		else
			ProcessLine = 0 'Success
			
		end if
	else												
		ErrorFile.writeline IP & " " & HostName & " Command Check Failed. Exiting Device at " & Now() &". Error in Input File (Category needs to be test or config). Deployment Batch Started at " & DeployStart
	end if
End Function

'----------------------------------------------------------------------------------------------------------------------------
'Name       : SaveConfig -> Saves config on the device and writes an entry into the Completed.txt file.
'Parameters : Logfiles   -> Directory containing the log files.
'           : IP         -> IP address of device
'			: HostName   -> HostName of device
'Return     : SaveConfig -> Returns 1 if the config was saved succesfully.
'----------------------------------------------------------------------------------------------------------------------------

Function SaveConfig(Logfiles, IP, HostName) 'Save the config and write to the log files
	objSc.Send "copy running-config startup-config" & VbCr
	objSc.WaitForString"[startup-config]?"
	objSc.Send VbCr
	objSc.WaitForString"#"
	SaveConfig = 1 'Return Code of 1 if successful save
	Set CompletedFile = FSO.OpenTextFile(Logfiles&"\Completed.txt",ForAppending, True)
	CompletedFile.writeline IP & " " & HostName & " " & Now() & ". Deployment Batch Started at " & DeployStart
	CompletedFile.Close()
End Function

'----------------------------------------------------------------------------------------------------------------------------
'Name       : CheckInputFiles -> Checks input files for extended ASCII (a specific problem is EN DASH).
'Parameters : Filename        -> File containing text to check for extended ASCII.
'Return     : CheckInputFiles -> Returns False if the files contains extended ASCII otherwise returns True.
'----------------------------------------------------------------------------------------------------------------------------
Function CheckInputFiles(Filename)
	CheckLineNum = 1
	
	Set CheckFile = FSO.opentextfile(Filename, ForReading, False)
	Do Until CheckFile.atEndOfStream
		CheckLine = CheckFile.ReadLine 'Read a line at a time
		For CheckCharNum = 1 to Len(CheckLine)
			Character = Mid(CheckLine,CheckCharNum,1) 'Read a character at a time
			If Asc(Character) > 126 then
				MsgBox _
				"Input File contains invalid characters." & vbCrLf & _
				"Filename " & Filename & vbCrLf & _
				"Character is at line " & CheckLineNum & vbCrLf & _
				"Character Number " & CheckCharNum
				WScript.Quit
			Else
				CheckInputFiles = TRUE  'Extended ASCII is Not Present
			End If
		Next
		CheckLineNum = CheckLineNum + 1
	Loop
	CheckFile.Close()
End Function

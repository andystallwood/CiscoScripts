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
HostFile = objD.Prompt("Enter folder name and Path to the hosts file","Folder Name & Path","U:\script\FindLANIF\Hosts.csv")

'Check file for invalid characters
If CheckInputFiles(HostFile) = FALSE then
	MsgBox("Host File contains invalid characters. Often this is Extended Dash, hidden as a dash")
	WScript.Quit
Else	
End If

'File containing a list of show commands to perform on each router. One show command per line.
CommandsFile = objD.Prompt("Enter folder name and Path to the show commands file","Folder Name & Path","U:\script\FindLANIF\Commands.csv")

'Check file for invalid characters
If CheckInputFiles(CommandsFile) = FALSE then
	MsgBox("Commands File contains invalid characters. Often this is Extended Dash, hidden as a dash")
	WScript.Quit
Else	
End If

'Folder to recieve the log files
Logfiles = objD.Prompt("Enter folder name and Path to save Log files In.","Folder Name & Path","U:\script\FindLANIF\logs\")

User = objD.Prompt("Enter YOUR Username to get into devices","Username","xxxxxxxxxx")

Pass = objD.Prompt("Enter password to get into devices","Password","xxxxxxxx", TRUE)

ErrorCount = 0

If FSO.FolderExists(Logfiles) Then
Else
	FSO.CreateFolder(Logfiles)
End IF

DeployStart = Now () '<---- Used for a deployment start Time stamp.

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

	'<-------------------- END of Connect sequence ------------------------------->
		
		ConfigLine = Commands.ReadLine 'Read Header Row of Commands CSV file and ignore
	
		'<-------------------- Device Configuration sequence ---------------------------------> 	
		Do While Not Commands.atEndOfStream '<---- Read each Commands line in turn from the CSV file and send to the device
			'Split Line Read into Category, Command, Prompt and Output and process the line
			ShowLine = Commands.ReadLine

			Call ProcessLine(ShowLine, Logfiles, IP, HostName)

		Loop
		'<-------------------- END of Device Configuration sequence ---------------------------------> 	

		Commands.Close() 'Close Command File	

		deployed = deployed + 1
		
		objSc.Send "exit" & vbCr
		objse.Log(False)
		objSc.Synchronous = False
		objSE.Disconnect
		
	Else 'Device failed to connect
		Set NoConnectfile = FSO.OpenTextFile(Logfiles&"\NoConnect.txt",ForAppending, True)
		NoConnectfile.writeline IP &  " " & HostName & " at " & Now() & " . Deployment Batch Started at " & DeployStart
		NoConnectCount = NoConnectCount + 1
		NoConnectfile.Close()
	End IF

Wend

Hosts.Close() 'Close Device IP File

'Write Summary
Set Summaryfile = FSO.OpenTextFile(Logfiles&"\Summary.txt",ForAppending, True)
Summaryfile.writeline "Scrape Started: " & DeployStart
Summaryfile.writeline "Scrape Complete: " & Now ()
Summaryfile.writeline "--------------"
Summaryfile.writeline "Total Number of devices: " & counter
Summaryfile.writeline "Total Number Scraped: " & deployed
Summaryfile.writeline "Total Number of warnings: " & ErrorCount
Summaryfile.writeline "--------------"
Summaryfile.Close()

MsgBox _
"Devices Scraped: " & deployed & vbCrLf & _
"Warnings: " &ErrorCount, _
vbOKOnly+vbInformation, _
"Script Completed"

'----------------------------------------------------------------------------------------------------------------------------
'Name       : ProcessLine -> Processes a line of the commands file.
'Parameters : ShowLine  -> Comma seperated string with the show command parameters present.
'           : Logfiles    -> Directory containing the log files.
'           : DeviceLine  -> Comma seperated string containing information about the device.
'Return     : Nothing returned.
'----------------------------------------------------------------------------------------------------------------------------

Sub ProcessLine(ShowLine, Logfiles, IP, HostName)
	
	objSc.Send ShowLine & VbCr 'Send Command to Device
	objSc.WaitForString"#"
	
	CaptureRow = objSc.CurrentRow - 1 'Subtract 1 from current cursor position, to capture command output, not prompt line
	CaptureLine = objSc.Get(CaptureRow, 1, CaptureRow, 80) 'Capture output from screen (all 80 characters)
	CaptureFirstChar = Left(CaptureLine,1)
	
	If CaptureFirstChar = " " Then
		CaptureRow = objSc.CurrentRow - 2 'Subtract 2 from current cursor position, to capture command output, not prompt line or blank line
		CaptureLine = objSc.Get(CaptureRow, 1, CaptureRow, 80) 'Capture output from screen (all 80 characters)
	Else
	End If
	
		Set OutputFile = FSO.OpenTextFile(Logfiles&"\Output.txt",ForAppending, True) 'Open error File ready to be written to
		
		OutputFile.writeline IP & "," & HostName & "," & CaptureLine

		OutputFile.Close()

End Sub


Function CheckInputFiles(Filename)
	Set File = FSO.opentextfile(Filename, ForReading, False)
	Do Until File.atEndOfStream
		Character = File.Read(1) 'Read a character
		If Asc(Character) > 126 then
			CheckInputFiles = FALSE 'Extended ASCII is Present
			Exit Do
		Else
			CheckInputFiles = TRUE  'Extended ASCII is Not Present
		End If
	Loop
	File.Close()
End Function

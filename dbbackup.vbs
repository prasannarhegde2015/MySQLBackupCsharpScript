'/*------ All Variables Declaration -----------------------------------------------
CommandDirectory="C:\Program Files\MySQL\MySQL Server 5.5\bin\"
baseDirectory="D:\Share\LocalData\LocalBackup\"
BackUpDirectory=baseDirectory&getBackupFolderName()&"\"
LogDirectory=baseDirectory&"Log"&"\"
LogFile=LogDirectory&"backuplog_"&getBackupFolderName()&".log"
squashNetworkBackupPath="U:\Quality Analysis\QA Team Backup\Squash Backup\"&getBackupFolderName()

'/---------Set up all Required Folder Structure --------------------------------------------------------------
If CreateObject("Scripting.FileSystemObject").FolderExists(LogDirectory)=False Then
	CreateObject("Scripting.FileSystemObject").CreateFolder(LogDirectory)
End If

If CreateObject("Scripting.FileSystemObject").FolderExists(BackUpDirectory)=False Then
	CreateObject("Scripting.FileSystemObject").CreateFolder(BackUpDirectory)
End If

If CreateObject("Scripting.FileSystemObject").FolderExists(squashNetworkBackupPath)=False Then
	CreateObject("Scripting.FileSystemObject").CreateFolder(squashNetworkBackupPath)
End If
'Call KillWinWord()
'-------------------------------------------------------------------------------------------------------------/
Set fso=CreateObject("Scripting.FileSystemObject")
Set f1=fso.CreateTextFile(LogFile)
'/***** Script Start
f1.WriteLine "Backup Script Started: Start time "&Now()
WScript.StdOut.WriteLine  "Created LogFile: "&LogFile
f1.WriteLine "Created LogFile: "&LogFile
readfilePath="D:\Share\squash.xls"
WScript.Sleep 3000
WScript.StdOut.WriteLine "1.)Folders Created taking backup now..."
f1.WriteLine "1.)Folders Created taking backup now..."
WScript.Sleep 3000
Dim globalTemplate

Call PerformMySqlBackUP()
Call PerformCompression()
Call PerformRoboCopy()
Call SendEmailNotification(squashNetworkBackupPath)
'Call KillWinWord()
f1.WriteLine "Backup Script Ended: End time "&Now()
'/***** Script End
Set f1=Nothing
Set fso=Nothing

'**************************************************************************'**************************************************************************
Public Function PerformMySqlBackUP()
	Call GetExpectedtData(readfilePath)
	'******************** Perofrm Database Bakcup Opearion *********************************************
	WScript.StdOut.WriteLine "2.) ******** Performing Database Backup : please Wait......"
	f1.WriteLine "2.) ******** Performing Database Backup : please Wait......"
	Do Until globalTemplate.EOF
		dbname=globalTemplate.Fields("dbname")
		tblname=globalTemplate.Fields("tablename")
		outputfilename=globalTemplate.Fields("outputfilename")
		WScript.StdOut.WriteLine "Command is: "&""""&CommandDirectory&"mysqldump.exe"""&" -u root -padmin "&dbname&"  "&tblname&" > "&""""&BackUpDirectory&outputfilename&""""
		f1.WriteLine "Command is: "&""""&CommandDirectory&"mysqldump.exe"""&" -u root -padmin "&dbname&"  "&tblname&" > "&""""&BackUpDirectory&outputfilename&""""
		cmdtext=""""&CommandDirectory&"mysqldump.exe"""&" -u root -padmin "&dbname&"  "&tblname&" > "&""""&BackUpDirectory&outputfilename&""""
		f1.WriteLine cmdtext
		Dim objShell
		Set objShell = WScript.CreateObject ("WScript.shell")
		f1.WriteLine "cmd /K"&""""&cmdtext&""""&"& Exit"
		WScript.StdOut.WriteLine "excecuting: "&cmdtext
		objShell.run "cmd /K"&""""&cmdtext&""""&"& Exit"
		Call WaitTillWindowPresent("mysqldump.exe")
		WScript.Sleep 1000
		Set objShell = Nothing
		globalTemplate.MoveNext
	Loop
	WScript.StdOut.WriteLine "3.) ******** Database Backup Opeation is complete for "&getBackupFolderName() 
	f1.WriteLine "3.) ******** Database Backup Opeation is complete for "&getBackupFolderName() 
End Function

'**************************************************************************'**************************************************************************
Public Function PerformCompression()
	DQ=""""
	targetrarFielName=BackUpDirectory&getBackupFolderName()&".rar"
	targetDirectory=BackUpDirectory
	f1.WriteLine "4.) ******** Archving Folders now "
	strRarCommand="rar a "&DQ&targetrarFielName&DQ&" -v256M -m2 "&DQ&targetDirectory&"*.sql"&DQ
	Set objShell = WScript.CreateObject ("WScript.shell")
	WScript.StdOut.WriteLine "cmd /K "&strRarCommand&"& Exit"
	objShell.run "cmd /K "&strRarCommand&"& Exit"
	Call WaitTillWindowPresent("rar")
	WScript.Sleep 1000
	f1.WriteLine "5.)******** Archving is Complete  "
	Set objShell = Nothing
	Set f2=fso.GetFolder(BackUpDirectory)
	Set allfiles=f2.Files
	For Each infile In  allfiles
		If InStr(infile.Name,".sql") Then
			fso.DeleteFile(infile)
		End If
	Next
End Function

'**************************************************************************'**************************************************************************
Public Function PerformRoboCopy()
	DQ=""""
	f1.WriteLine "6.)******** Copying to Network Backup location   "
	Set objShell = WScript.CreateObject ("WScript.shell")	
	sourcePath=baseDirectory&getBackupFolderName()
	strRobocopycmd="robocopy "&DQ&sourcePath&DQ&"  "&DQ&squashNetworkBackupPath&DQ&"  *.* /e /z "
	f1.WriteLine "7)******** Command is    "&"cmd /K "&strRobocopycmd&"& Exit"
	WScript.StdOut.WriteLine "cmd /K "&strRobocopycmd&"& Exit"
	objShell.run "cmd /K "&strRobocopycmd&"& Exit"
	Call WaitTillWindowPresent("robocopy")
	Set objShell = Nothing

End Function


'**************************************************************************'**************************************************************************
Function  GetExcelConnection(filePath)
	Dim con
	Dim connectionString
	Set con = CreateObject("ADODB.Connection")
	connectionString="Driver={Microsoft Excel Driver (*.xls)};DriverId=790;;Dbq="&filePath&";"
	con.CursorLocation=3
	con.Open connectionString
	Set GetExcelConnection= con
End Function 

'**************************************************************************'**************************************************************************
Function GetExpectedtData(dataPath)
	Dim con
	Dim sheetName
	Dim dataSheetName
	sheetName="[Sheet1$]"
	Set con = GetExcelConnection(dataPath)
	Set globalTemplate= con.Execute( "SELECT *   from " + sheetName )
	Set con= Nothing 
End Function

'**************************************************************************'**************************************************************************
Function getBackupFolderName()
	curDate = Now 
	dd=Day(curDate)
	mt=Month(curDate)
	yyyy=Year(curDate)
	WScript.StdOut.WriteLine mt
	strmonth=""
	Select Case mt
		Case 1: strmonth="Jan"
		Case 2: strmonth="Feb"
		Case 3: strmonth="Mar"
		Case 4: strmonth="Apr"
		Case 5: strmonth="May"
		Case 6: strmonth="Jun"
		Case 7: strmonth="Jul"
		Case 8: strmonth="Aug"
		Case 9: strmonth="Sep"
		Case 10: strmonth="Oct"
		Case 11: strmonth="Nov"
		Case 12: strmonth="Dec"
	End Select
	WScript.StdOut.WriteLine strmonth
	bkfolderName=CStr(dd)&"_"&strmonth&"_"&CStr(yyyy)
	WScript.StdOut.WriteLine bkfolderName
	getBackupFolderName=bkfolderName
End Function
'**************************************************************************'**************************************************************************


Function WaitTillWindowPresentOrig(windowTitleText)
	Set Word = CreateObject("Word.Application")
	Set Tasks = Word.Tasks
	mywindowFalg=True
	timeAccumalatedinMinutes=0
	While mywindowFalg=True
		tasklist=""
		For Each Task In Tasks
			tasklist=tasklist&";"&Task.name
			'If Task.name = "Administrator: Command Prompt - Start  /?" Then
		Next
		'''Get all tasks
		If InStr(tasklist,windowTitleText) Then
			mywindowFalg=True
		Else
			mywindowFalg=False
		End If
		WScript.Sleep 5000
		timeAccumalatedinMinutes=timeAccumalatedinMinutes+(1/12)
		WScript.StdOut.WriteLine "Performing Task......Cumulative time in mins: "&timeAccumalatedinMinutes
	Wend
	Word.Quit
	Set Tasks = Nothing
	Set Word =Nothing
	Call KillWinWord()
End Function


Function WaitTillWindowPresent(windowTitleText)
    DQ=""""
	f1.WriteLine "6.)******** Waiting For Window :  "&windowTitleText
	Set objShell = WScript.CreateObject ("WScript.shell")
	strWaitUtil = "D:\share\LocalData\WaitTillWindowExist.exe"	
	strUtilcmd = strWaitUtil&" "&windowTitleText
	objShell.run "cmd /K "&strUtilcmd&"& Exit"
	WScript.Sleep 3000
   chkfile =objShell.ExpandEnvironmentStrings("%userprofile%")&"\wait.txt"
	While CreateObject("Scripting.FileSystemObject").FileExists(chkfile)
	f1.WriteLine "6.)******** File Exists :  "&chkfile
	  f1.WriteLine "6.)******** Waiting For Window :  "&windowTitleText
	  WScript.Sleep 3000
	Wend 
End Function

'**************************************************************************'**************************************************************************

Function SendEmailNotification(folderPath)
	Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory. 
	Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network). 
	Const cdoAnonymous = 0 'Do not authenticate
	Const cdoBasic = 1 'basic (clear-text) authentication
	Const cdoNTLM = 2 'NTLM
	Set oCDO = CreateObject("CDO.Message") 
	oCDO.Subject = "Squash backup-"&getBackupFolderName()
	oCDO.From = "Prasanna.Hegde@me.weatherford.com"
	oCDO.To = "swati.whaval@me.weatherford.com;Ashok.Krishna@me.weatherford.com"
	oCDO.CC="Prasanna.Hegde@me.weatherford.com"
	flag=1
	Set fileSys=CreateObject("Scripting.FileSystemObject")
	Set f2=filesys.GetFolder(folderPath)
	fileCount=f2.Files.Count
	If fileCount=0 Then
		oCDO.TextBody = "Squash backup on U drive has been failed."
	Else
		If fileCount>0 Then
			oCDO.TextBody ="Total of "&fileCount&" files in Compressed Format chunks of 250 MB: are created. Please check path "&vbNewLine&folderPath
		End If
	End If
	'==This section provides the configuration information for the remote SMTP server.
	oCDO.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
	'Name or IP of Remote SMTP Server
	oCDO.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail2.weatherford.com"
	'Type of authentication, NONE, Basic (Base64 encoded), NTLM
	oCDO.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
	'Server port (typically 25)
	oCDO.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserverport") =25
	'Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP server)
	oCDO.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
	oCDO.Configuration.Fields.Update
	'==End remote SMTP server configuration section==
	oCDO.Send
	
	Set f2 =Nothing
	Set fileSys=Nothing
	Set oCDO=Nothing
	Set objFSO=Nothing
End Function


Function KillWinWord()

Dim objWMIService, objProcess, colProcess
	Dim strComputer, strProcessKill 
	strComputer = "."
	strProcessKill = "'WINWORD.exe'" 
	
	Set objWMIService = GetObject("winmgmts:" _
	& "{impersonationLevel=impersonate}!\\" _ 
	& strComputer & "\root\cimv2") 
	
	Set colProcess = objWMIService.ExecQuery _
	("Select * from Win32_Process Where Name = " & strProcessKill )
	For Each objProcess In colProcess
		objProcess.Terminate()
	Next 
	
End Function

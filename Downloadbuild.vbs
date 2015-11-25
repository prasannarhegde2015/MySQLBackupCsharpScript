'Option Explicit
'------------  All Variables are declared on the Top -------------------------------------------------------
dst="U:\Quality Analysis\QA Team Backup\WellFlo\1.Wellflo Builds\6.0"
flpath = "E:\DailyBackups\BuildCopy\WellFlo6\latestfilename.txt"
lgpath="E:\DailyBackups\BuildCopy\TaskLog.txt"

'------------  All Variables are declared on the Top -------------------------------------------------------
Set logfile = New TextFileOp
logfile.FilePath = lgpath
Dim fileSystem, folder, file ,objMessage
Dim path ,DestinationFile,sourcepath,destinationpath,sourcepath2,destinationpath2
Dim cmdrobo,src,dest,oShell,strfnamearr,strfnamearr1,strLargedate,i

flname = Trim( Replace(GetLatestFileName(flpath),vbLf,""))
logfile.WriteLog "Latest FileName:"&flname
WScript.StdOut.WriteLine  "Latest FileName:"&flname
lastbkslpos=InStrRev(flname,"\")
src = Mid(flname,1, (Len(flname) - (Len(flname)- lastbkslpos +1 )) )
WScript.StdOut.WriteLine "Source foloder:"&src
lastbkslpos1=InStrRev(src,"\")

flname1=Mid(flname,lastbkslpos+1, (Len(flname)- lastbkslpos + 1 ))
WScript.StdOut.WriteLine "src modified "&flname1

dstname1=Mid(src,lastbkslpos1+1, (Len(src)- lastbkslpos1 + 1 ))

If Not CreateObject("Scripting.FileSystemObject").FolderExists(dst&"\"&dstname1) Then
     CreateObject("Scripting.FileSystemObject").CreateFolder(dst&"\"&dstname1)
End If
Call DoRobocoy(src,dst&"\"&dstname1,flname1)
'Call SendEmail ()


Public Function DoRobocoy(sourcepath,destinationpath,strLargedate)
	Set oShell=WScript.CreateObject("WScript.shell")
	Const Q=""""
	cmdrobo="robocopy "&Q&sourcepath&Q&" "&Q&destinationpath&Q&" "&Q&strLargedate&Q&"  /z"
	WScript.StdOut.WriteLine "Command is " & cmdrobo
	logfile.WriteLog "Command is " & cmdrobo
	oshell.Run"cmd.exe /C"&cmdrobo,1,True
	Set oShell=Nothing
End Function

Function SendEmail ()
	Set objMessage = CreateObject("CDO.Message") 
	objMessage.Subject = "Download completed" 
	objMessage.From = "prasanna.hegde@me.weatherford.com" 
	'objMessage.To = "moumeeta.naskar@me.weatherford.com;Prasanna.Hegde@me.weatherford.com;Deepankar.Bandopadhyay@me.weatherford.com;swati.whaval@me.weatherford.com"
	objMessage.To = "prasanna.hegde@me.weatherford.com;Ritu.Sah@ME.Weatherford.com;Gaurav.Kakad@me.weatherford.com"
	objMessage.TextBody = "Downloaded the latest build of WellFlo successfully at location: "&destinationpath
	objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail2.weatherford.com" 'Modify to your SMTP Server Address
	objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	objMessage.Configuration.Fields.Update
	objMessage.Send
End Function


Function GetLatestFileName(strpath)
GetLatestFileName=""
Set fso1 = CreateObject("Scripting.FileSystemObject")
Set f11= fso1.OpenTextFile(strpath)
GetLatestFileName=f11.ReadAll
f11.Close
 fso1.DeleteFile(strpath)
Set f11=Nothing
Set fs011= nothing

End Function


Class TextFileOp
	Private fso,f1
	Private pfilepath
	
	Private Sub Class_Initialize()
		Set fso = CreateObject("Scripting.FileSystemObject")
	End Sub
	
	Public Property Let FilePath(value)
	pfilepath = value
	End Property
	
	Public Sub WriteLog(stxt)
		If Not fso.FileExists(pfilepath) Then
			Set f1 = fso.CreateTextFile(pfilepath)
			f1.WriteLine( "["&Now&"]: "&stxt)
			f1.Close
		Else
			Set f1 = fso.OpenTextFile(pfilepath,8,True)
			'   
			f1.WriteLine( "["&Now&"]: "&stxt)
			f1.Close
		End If
	End Sub
	
	Private Sub Class_Terminate()
		Set f1 =Nothing
		Set fso = Nothing
	End Sub
	
End Class

CommandDirectory="C:\Program Files\MySQL\MySQL Server 5.5\bin\"
BackUpDirectory="U:\Quality Analysis\QA Team Backup\Squash Backup\6_Apr_2013\"

LogFile="e:\test\log1.txt"
Set fso=CreateObject("Scripting.FileSystemObject")
Set f1=fso.CreateTextFile(LogFile)
WScript.StdOut.WriteLine  "Created LogFile: "&LogFile
f1.WriteLine "Created LogFile: "&LogFile
readfilePath="E:\SquashDatabaseBackup\squash.xls"

WScript.Sleep 3000
WScript.StdOut.WriteLine "1.)Restoring DB.."
WScript.Sleep 3000
Dim globalTemplate

Call GetExpectedtData(readfilePath)
'******************** Perofrm Database Bakcup Opearion *********************************************
WScript.StdOut.WriteLine "2.)******** Performing Database Restore : please Wait......"
f1.WriteLine "2.)******** Performing Database Restore :  please Wait......"

Do Until globalTemplate.EOF
	dbname=globalTemplate.Fields("dbname")
	tblname=""
	outputfilename=globalTemplate.Fields("outputfilename")
	WScript.StdOut.WriteLine "Command is: "&""""&CommandDirectory&"mysql"""&" -u root -padmin "&dbname&"  "&tblname&" < "&""""&BackUpDirectory&outputfilename&""""
    f1.WriteLine "Command is: "&""""&CommandDirectory&"mysql"""&" -u root -padmin "&dbname&"  "&tblname&" < "&""""&BackUpDirectory&outputfilename&""""
	cmdtext=""""&CommandDirectory&"mysql"""&" -u root -padmin "&dbname&"  "&tblname&" < "&""""&BackUpDirectory&outputfilename&""""
	f1.WriteLine cmdtext
	Dim objShell
	Set objShell = WScript.CreateObject ("WScript.shell")
	f1.WriteLine "cmd /K"&""""&cmdtext&""""&"& Exit"
	objShell.run "cmd /K"&""""&cmdtext&""""&"& Exit"
	WScript.Sleep 5000
	Set objShell = Nothing
	globalTemplate.MoveNext
Loop
WScript.StdOut.WriteLine "4.)******** Database Restore Opeation is complete !!!!! "
f1.Close
Set f1=Nothing
Set fso=Nothing
'**************************************************************************
Function  GetExcelConnection(filePath)
	Dim con
	Dim connectionString
	Set con = CreateObject("ADODB.Connection")
	connectionString="Driver={Microsoft Excel Driver (*.xls)};DriverId=790;;Dbq="&filePath&";"
	con.CursorLocation=3
	con.Open connectionString
	Set GetExcelConnection= con
End Function 

Function GetExpectedtData(dataPath)
	Dim con
	Dim sheetName
	Dim dataSheetName
	sheetName="[Sheet1$]"
	Set con = GetExcelConnection(dataPath)
	Set globalTemplate= con.Execute( "SELECT *   from " + sheetName )
	Set con= Nothing 
End Function

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
		Case 10: strmonth="Nov"
		Case 12: strmonth="Dec"
	End Select
	WScript.StdOut.WriteLine strmonth
	bkfolderName=CStr(dd)&"_"&strmonth&"_"&CStr(yyyy)
	WScript.StdOut.WriteLine bkfolderName
	getBackupFolderName=bkfolderName
End Function

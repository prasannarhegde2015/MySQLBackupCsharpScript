Dim bullet
Dim response
Dim globalTemplate
bullet = Chr(10) & "   " & Chr(149) & " "
codepfilepath = "c:\RTUEMU\code.txt"
excelFilepath = "c:\RTUEMU\RTEMUData.xls"
displayText=""
Set fso = CreateObject("Scripting.FileSystemObject")
Set f1 = fso.OpenTextFile(codepfilepath)
 displayText = f1.ReadAll()
 
 f1.Close
 Set f1=Nothing
 Set fso= nothing
'MsgBox displayText
Do
    response = InputBox("Please enter the number that corresponds to your selection:" & Chr(10) & bullet & displayText)
    If response = "" Then WScript.Quit  'Detect Cancel
    If IsNumeric(response) Then Exit Do 'Detect value response.
    MsgBox "You must enter a numeric value.", 48, "Invalid Entry"
Loop
'MsgBox "The user chose :" & response, 64, "Yay!"
Call GetExpectedtData(excelFilepath)

reqpath=""
Do Until globalTemplate.EOF
 scode = globalTemplate.Fields("Code").value
 spath = globalTemplate.Fields("FileFullpath").value
  If CStr(scode) = CStr(response)Then
       reqpath =  spath
  End If
globalTemplate.movenext
Loop 


'MsgBox "Full Path: "&reqpath

Call UpdateConfigXmlFileValue(reqpath)

Function UpdateConfigXmlFileValue(strConfigPathValue)

Set xmlDoc = CreateObject("Msxml2.DOMDocument")
xmlDoc.async = False
xmlDoc.load "C:\RTUEMU\RtuEmu.exe.config"
 'WScript.Echo xmlDoc.parseError
If xmlDoc.parseError = 0 Then
  For Each x In xmlDoc.selectNodes("//setting/*")
    'WScript.Echo x.parentNode.getAttribute("name") & ": " _
    '  & x.getAttribute("name")
	'WScript.Echo  x.text
	''or  use x.item(0).text
	if x.parentNode.getAttribute("name") = "configFilePath"  then
	     x.text = strConfigPathValue
	end if
  Next
  
End If
xmlDoc.Save "C:\RTUEMU\RtuEmu.exe.config"
End Function


Function  GetExcelConnection(filePath)
	Dim con
	Dim connectionString
	Set con = CreateObject("ADODB.Connection")
	connectionString="   Driver={Microsoft Excel Driver (*.xls)};DriverId=790;;Dbq=" &  filePath  & ";ReadOnly=0;"
	con.CursorLocation=3
	'MsgBox filePath 
	'MsgBox connectionString
	con.Open connectionString
	Set GetExcelConnection= con
End Function 


Function GetExpectedtData(dataPath)
	Dim con
	Dim sheetName
	Dim dataSheetName
	sheetName="[Template$] "
	Set con = GetExcelConnection(dataPath)
	Set globalTemplate= con.Execute( "SELECT *   from " + sheetName )
	Set con= Nothing 
End Function

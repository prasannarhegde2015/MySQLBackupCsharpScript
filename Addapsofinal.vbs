'/****************************************************************

'/*************************************** Initial Call ***************************************
MsgBox "Instructions:"&_
vbNewLine&"Enter 1 : To do action for single Excel file"&vbNewLine& _
"Enter 2 : To do action for Batch ( All *.xls files inside a Folder)", _
64,"Function Add Apos to numeric Excel cells"
action=InputBox("Please Enter your choise: 1 or 2 ","Function Add Apos to numeric Excel cells")
'/*********************** Some Validation *******************************************************
If Not IsNumeric(action) Then
	MsgBox "Invalid Option was selected !! Script Will Quit!",16,"Invalid Selection"
	WScript.Quit
End If
If CDbl(action) <> 1 And CDbl(action) <> 2 Then
	MsgBox "Invalid Option was selected !! Script Will Quit!",16,"Invalid Selection"
	WScript.Quit
End If


Select Case action
	Case "1": '/************************ Sinlge File Operation **********************************/
	
	If Open_Win_Dialog(fname) =False Then
		fname=InputBox("Please Enter File Name Manually By typing (as invoking MScomDlg Failed  !!! ")
		If InStr(1,fname,".xls",1) <> 0 Then
			If CreateObject("Scripting.FileSystemObject").FileExists(fname) Then
				MsgBox "File Name Succesfully accepted",64,"File Selection Success"
			Else
				MsgBox "Given xls file was Not found !! Script Will Quit!",48,"Script Termination"
				WScript.Quit
			End If 
		Else
			MsgBox "This Script Takes only .xls files which exist on your machine !! Script Will Quit!",48,"Script Termination"
			WScript.Quit
		End If
	End If
	
	Set prn = New IELoader
	
	str1="Please Wait<br><img src='http://home.gopal.com/im2.gif' height='100' width='100'>"	
	str2=str2&"<font color='green' size='8' >Action has been Completed </font>"
	str2=str2&"*******************************************************************************"
	prn.createdlg()
	prn.showcontent str1
	Set excelObj=CreateObject("Excel.Application")
	failvalue=""
	nCellAddress=""
	ncount=0
	Set workbookObj=excelObj.Workbooks.Open(fname)
	For nCount=1 To workbookObj.Worksheets.Count
		npcount=0
		Set worksheetObj=workbookObj.Sheets(nCount)
		sdetails= sdetails&"Worksheet Name ="&worksheetObj.Name&vbNewLine
		sdetails= sdetails& "rows "&worksheetObj.UsedRange.Rows.count&vbNewLine
		sdetails= sdetails&	 "Columns "&worksheetObj.UsedRange.Columns.count&vbNewLine
		npdt=worksheetObj.UsedRange.Rows.count*worksheetObj.UsedRange.Columns.count
		counter=0
		strstatustext="<br>Adding apos to sheet :"&worksheetObj.Name
		For Each indCell In worksheetObj.UsedRange
			npcount=CInt((counter/npdt)*100)
			sdetails= sdetails& "Value of Cell = "&indCell.Value&vbNewLine
			If indCell.value <> Null Or indCell.value<>"" Then
				If IsNumeric(indCell.value) Then
					indCell.value = "'" & indCell.value
				End If
			End If
			prn.updateprogress strstatustext,npcount
			'WScript.Sleep 100
			counter=counter+1
		Next 
	Next
	prn.showcontent str2
	WScript.Sleep 2000
	prn.closedlg()
	WScript.Sleep 100
	workbookObj.Save
	workbookObj.Close
	excelObj.Quit
	MsgBox "Done with Action ",64,"Script Complete"
	Set worksheetObj=Nothing
	Set workbookObj=Nothing
	Set excelObj=Nothing
	
	Case "2" :'/************************ All *.xls files inside a folder Operation **********************************/
	fname=InputBox("Please Enter Folder Path by typing ...Valid Folder Path  ","Function Add Apos to numeric Excel cells")
	If CreateObject("Scripting.FileSystemObject").FolderExists(fname) Then
		MsgBox "Foldeer Name Succesfully accepted",64,"Folder Correct"
	Else
		MsgBox "Given Foldeer Name !! Script Will Quit!",16,"Quittting.."
		WScript.Quit
	End If 
	Set prn = New IELoader
	str1="Please Wait <img src='http://phegde2009.freevar.com/images/loading.gif' height='100' width='100'>"
	str2="*******************************************************************************"
	str2=str2&"<font color='green' size='7' ><strong>All Action has been Completed.</strong> </font>"
	str2=str2&"*******************************************************************************"
	prn.createdlg()
	
	Set folderobj=CreateObject("Scripting.FileSystemObject").Getfolder(fname)
	Set allFiles=folderobj.Files
	For Each eachFile In allFiles
		prn.showcontent str1
		status1="<b>Doing action wait.....File Name : "&eachFile.name&"<b>"
		prn.updateprogress status1,""
		If InStr(1,eachFile.name,".xls",1) <> 0 Then
			Set excelObj=CreateObject("Excel.Application")
			failvalue=""
			nCellAddress=""
			ncount=0
			Set workbookObj=excelObj.Workbooks.Open(eachFile)
			'Iterrate Through all worksheets in the workbook
			For nCount=1 To workbookObj.Worksheets.Count
				npcount=0
				Set worksheetObj=workbookObj.Sheets(nCount)
				sdetails= sdetails&"Worksheet Name ="&worksheetObj.Name&vbNewLine
				sdetails= sdetails& "rows "&worksheetObj.UsedRange.Rows.count&vbNewLine
				sdetails= sdetails&	 "Columns "&worksheetObj.UsedRange.Columns.count&vbNewLine
				npdt=worksheetObj.UsedRange.Rows.count*worksheetObj.UsedRange.Columns.count
				unit_incr_size= CDbl(npdt/100)
				prsize=CDbl(unit_incr_size/npdt)* 100
				counter=0
				strstatustext=status1&"<br>Adding apos to sheet :"&worksheetObj.Name
				For Each indCell In worksheetObj.UsedRange
					npcount=CInt((counter/npdt)*100)
					sdetails= sdetails& "Value of Cell = "&indCell.Value&vbNewLine
					If indCell.value <> Null Or indCell.value<>"" Then
						If IsNumeric(indCell.value) Then
							indCell.value = "'" & indCell.value
						End If
					End If
					prn.updateprogress strstatustext,npcount
					counter=counter+1
				Next 
			Next
			WScript.Sleep 1000
			workbookObj.Save
			workbookObj.Close
			excelObj.Quit
			Set worksheetObj=Nothing
			Set workbookObj=Nothing
			Set excelObj=Nothing
		End If
		
	Next
	prn.showcontent str2
	prn.closedlg()
	WScript.Sleep 20
	MsgBox "Done with Action !!!!",64,"Complete"
	Case Else : MsgBox "Cant come here"
End Select
'/***************************** Try to use CommonDialog Object***********************
Function Open_Win_Dialog(fname)
	Open_Win_Dialog=True
	On Error Resume Next 
	Dim CDObj 
	Set CDObj = CreateObject("UserAccounts.CommonDialog") 
	If Err Then 
		Open_Win_Dialog=False 
		Exit Function 
	End If 
	CDObj.MaxFileSize = 260 ' Init buffer (NECESSARY!) 
	CDObj.Filter="Excel 2003|*.xls|Excel 2007|*.xlsx"
	CDObj.FilterIndex=1
	CDObj.ShowOpen 
	fname = CDObj.filename 
	' Cancel or no file name? 
	If fname <> vbNullString Then 
		MsgBox "The file you choose is " & fname 
	Else
		Open_Win_Dialog=False
		MsgBox "Wrong filters bwre set script would Exit"
		WScript.Quit
	End If 
End Function 
'/***************************** LOADER OBJECT ***********************

'/***************************** CREATE IE OBJECT ***********************

Class IELoader
	'/***************************** CREATE ***********************
	Private obj_IE	
	Public Sub createdlg()
		Set obj_IE = CreateObject("InternetExplorer.Application")
		obj_IE.Navigate "about:blank" 
		obj_IE.ToolBar = 0
		obj_IE.StatusBar = 0
		obj_IE.Width=400
		obj_IE.Height = 200 
		obj_IE.Left = 0
		obj_IE.Top = 0
		obj_IE.Resizable=False
		obj_IE.Document.Title="Progress Bar"
		Do While (obj_IE.Busy)
			WScript.Sleep 200
		Loop 
		obj_IE.Visible = 1 
	End Sub
	
	'/***************************** SHOW ***********************
	Public Sub showcontent(stxt)
		strdiv="<div id='textstatus' name='textstatus' ></div>"
		strdiv=strdiv&"<br><div id='progressbar' name='progressbar' ></div>"
		obj_IE.Document.Body.InnerHTML= strdiv&stxt
		WScript.Sleep 50
	End Sub
	
	'	/***************************** UPDATE DIV Sections ***********************
	Public Sub updateprogress(prgtext,intpercent)
		obj_IE.Document.GetElementById("textstatus").InnerHTML=prgtext
		obj_IE.Document.GetElementById("progressbar").InnerHTML= "Progress Status : "&intpercent&"%"
	End Sub
	'	/***************************** Gracefuly Close ***********************
	Public Sub closedlg()
		WScript.Sleep 20
		obj_IE.Quit
		Set obj_IE=Nothing
	End Sub
	
End Class
'/***************************************Use this Class that amkes use of IE ****************************

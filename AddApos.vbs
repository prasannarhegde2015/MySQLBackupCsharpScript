Set odlg=CreateObject("UserAccounts.CommonDialog")
odlg.Filter="Excel 2003|*.xls|Excel 2007|*.xlsx"
odlg.FilterIndex=1
odlg.InitialDir="C:\Project\WellFlo\Test_input_data\Structured_Data"
odlg.ShowOpen()
sexcelfile1=odlg.FileName


Call loader("ON")
Call Trim_Excel_file(sexcelfile1)
Call loader("OFF")

'/*****************************************************************
Sub Trim_Excel_file(sexcelfile)
	Set oxl=CreateObject("excel.application")
	oxl.workbooks.open(sexcelfile)
	For isht= 1 To oxl.ActiveWorkbook.Worksheets.Count
		Set osht=oxl.activeworkbook.worksheets(isht)
		For i=1 To osht.UsedRange.rows.count
			For j=1 To osht.UsedRange.columns.count
				If IsNumeric (osht.cells(i,j).value) Then
					osht.cells(i,j).value="'"&(osht.cells(i,j).value)
				End If
				
			Next
		Next
	Next
	oxl.displayalerts=False
	oxl.save
	oxl.quit
	Set osht=Nothing
	Set oxl=Nothing
End Sub
'/*****************************************************************


Public Function Loader(arg)
Set newiedlg= New Iedialog
stxt1="Please Wait <br> <img src='http://phegde2009.freevar.com/images/loading.gif'  > "
stxt2="<font color='red' size='6' > Test Script Execution : Complete</font>"
      Select Case arg
      Case "ON" : newiedlg.CreateDlg()
      			  newiedlg.showIEdialog sxt1  

      Case "OFF" :newiedlg.showIEdialog sxt2
      			  newiedlg.Close()
      Case Else : MsgBox "Invalid Argument"
      End Select
End Function


'/*****************************************************************
Class Iedialog
	Private objIE
	''Private Sub Class_Initialize()
	'	Set objIE=CreateObject("InternetExplorer.Application")
	'	objIe.Visible=True
	'	objIe.Navigate "about:Blank"
	'	objIE.StatusBar=0
	'	objIE.ToolBar=0
	'	objIe.Left=0
	'	objIE.Top=0
	'	objie.Width=200
	'	objIE.Height=200
	'	objIE.Resizable=False
'	End Sub
	
	Public Sub CreateDlg()
		Set objIE=CreateObject("InternetExplorer.Application")
		objIe.Visible=True
		objIe.Navigate "about:Blank"
		objIE.StatusBar=0
		objIE.ToolBar=0
		objIe.Left=0
		objIE.Top=0
		objie.Width=200
		objIE.Height=200
		objIE.Resizable=False
	End Sub
	
	Public Sub showIEdialog(str)
		objIE.Document.body.innerHTML= str
	End Sub
	
'/	Private Sub Class_Terminate()
 '/ /      WScript.Sleep 5000
'		objIE.Quit
'		Set objIE=Nothing
'	End Sub
	
	 Public Function Close()
        objIE.quit
        objIE = Nothing
    End Function
	
End Class
'/*****************************************************************

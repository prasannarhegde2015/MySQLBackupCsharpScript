
strMasterTCFile="C:\LicenseCheck\Machine Details_Production_Optimization.xlsx"
Set objMxls = CreateObject("Excel.Application")
Set owbM = objMxls.Workbooks.Open(strMasterTCFile)
Set objS= owbM.Worksheets("Sheet1")
colCount = objS.usedrange.columns.count
rowCount = objS.usedrange.rows.count
Dim machineName(50)
Dim ownerName(50)
Dim expireDate(50)
Dim status(50)
Int k=0
 set fso = CreateObject("Scripting.FileSystemObject")
 
 set f1 = fso.createTextfile("C:\lic.txt")
 f1.writeline("Executing lice check")
 
For i= 2 To rowcount

sofwaresExpiry = ""
k = i-2
	For j = 10 To colCount
	f1.writeline(objS.cells(i,j)& " and  "& objS.cells(i,1))
		If objS.cells(i,j)<> "NA" and objS.cells(i,1) = "Y" Then
				expDate= objS.cells(i,j) - 7
				curDate = Now
				f1.writeline("Expiry date week before:  "&expDate&"  and Condition:=  "& (expDate <= curDate))
				If(expDate <= curDate)Then
						
					machineName(k) = objS.cells(i,4)
					ownerName(k) = objS.cells(i,2)
					expireDate(k) = objS.cells(i,10)
					If objS.cells(i,16) < curDate Then
						status(k) = "Expired"
					Else
						status(k) = "Will Expire"
					End If					
					sofwaresExpiry = vbCrLf  & objS.cells(1,j)	& vbCrLf  & sofwaresExpiry	
					test = 	sofwaresExpiry	
					f1.writeline("test is "&test)
                Else
				    f1.writeline("No License Expiry for "&k)
				End If
				
		End If
		
	Next

Next

objMxls.Workbooks.Close()
Set objMxls = Nothing

				If test <> "" Then
					Set oCDO = CreateObject("CDO.Message") 
					oCDO.Subject = "License "
					oCDO.From = "ashok.krishna@me.weatherford.com"
					'oCDO.To = "Deepankar.Bandopadhyay@me.weatherford.com;ashok.krishna@me.weatherford.com"
					oCDO.To = "Ramya.Adhikari@ME.Weatherford.com;Prasanna.Hegde@me.weatherford.com"
					oCDO.CC="ashok.krishna@me.weatherford.com"
					oCDO.TextBody = "Hi All, " &vbCrLf&vbTab &"Please find License Expiry details for PO team in the attachment. "
'					 & vbCrLf&vbCrLf&vbCrLf&vbCrLf & "MachineOwner		MachineName		Status		ExpiryDate" 
					Set oWord = CreateObject("Word.Application")
					oWord.visible = False
					Set oDoc = oWord.Documents.Add()
					Set oRange = oDoc.Range()
					oDoc.Tables.Add oRange,30,4
					Set oTable = oDoc.Tables(1)
					oTable.Cell(1,1).Range.Text="MachineOwner"
					oTable.Cell(1,2).Range.Text="MachineName"
					oTable.Cell(1,3).Range.Text="Status"
					oTable.Cell(1,4).Range.Text="ExpireDate"
					o =0
					p=2	
					For n=2 To 50				
					If 	ownerName(o) <> "" Then												
					oTable.Cell(p,1).Range.Text=ownerName(o)
					oTable.Cell(p,2).Range.Text=machineName(o)
					oTable.Cell(p,3).Range.Text=status(o)
					oTable.Cell(p,4).Range.Text=expireDate(o)
					p = p+1					
					End If
					o = o+1
					Next
				'	Make the First Row Data to Bold
					oTable.Rows.Item(1).Range.Font.Bold = True
					'Makes the First Row Data to Italic
					oTable.Rows.Item(1).Range.Font.Italic = True
				'	Changes the font size
					oTable.Rows.Item(1).Range.Font.Size = 12
					oTable.Rows.Item(1).Range.Font.Underline = True
				'	Sets the width of the Column
					oTable.Columns.Item(1).SetWidth 100,0
					
					oDoc.SaveAs("C:\Licensing.doc")
					oWord.Quit
					Set oWord = Nothing					
					WScript.Sleep 30000
												
					oCDO.TextBody =oCDO.TextBody& "Please renew and  update new license details in 'Machine Details_Production_Optimization.xlsx' Location : \\MEINWESSVMQA06\LicenseCheck" & vbCrLf&vbCrLf&vbCrLf&vbCrLf&vbCrLf&vbCrLf&vbCrLf& " Thanks and Regards"& vbCrLf&"Ashok Krishna K"
					
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
					oCDO.AddAttachment "C:\Licensing.doc"
					oCDO.Send					
					
					set oCDO=Nothing 
					Else
					Set oCDO = CreateObject("CDO.Message") 
					oCDO.Subject = "License "
					oCDO.From = "ashok.krishna@me.weatherford.com"
					oCDO.To = "Prasanna.Hegde@me.weatherford.com;ashok.krishna@me.weatherford.com"
					'oCDO.To = "Deepankar.Bandopadhyay@me.weatherford.com"
					oCDO.CC="ashok.krishna@me.weatherford.com"
					oCDO.TextBody ="Hello ,"& vbCrLf&vbCrLf&vbCrLf&"There are no machines for which License will expire with in a week" & vbCrLf&vbCrLf&vbCrLf&vbCrLf&vbCrLf&vbCrLf&vbCrLf& " Thanks and Regards"& vbCrLf&"Ashok Krishna K"
					
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
					
				End If
				
				f1.close
				
				set f1=nothing
				set fso= nothing


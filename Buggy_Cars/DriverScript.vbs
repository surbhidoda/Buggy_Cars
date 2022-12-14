'----------------------'
'-----Declaration------'
'----------------------'
Public testDir , testSuitControllerFile ,configSheet , controllerSheet , environmentSheet , testScript ,TCStatus , qtApp, test
Public iTCExecuted , iStartTime , iEndTime , strMailTO , strMailCC , MyFile
Public reportingMode , attachmentMode , strSummaryReportFile , attachmentData , strSharepointLink 
Dim strURL , strZipWithDateTimeStamp

On error resume next
'---------------------------------------------------'
'------- Finding Path and Variable declaration------'
'---------------------------------------------------'
testDir = createobject("Scripting.Filesystemobject").GetAbsolutePathName(".")
'msgbox testDir
'msgbox "Hello"
testSuitControllerFile = testDir & "\" & "Controller.xlsx"
TCStatus = testDir & "\" & "Library" & "\" & "TCStatus.txt"
iPTCExecuted = 0
iPTCFailCount = 0
iPTCPassCount = 0

iTCExecuted = 0
iTCFailCount = 0
iTCPassCount = 0
'Dim qtApp			 'As QuickTest.Application Declare the Application object variable
Dim qtTest			 'As QuickTest.Test Declare a Test object variable
Dim qtResultsOpt  		' Declare a Run Results Options object variable	
Set qtApp = CreateObject("QuickTest.Application") 			' Create the Application object
qtApp.Launch 
'msgbox "QTP Launched"												' Start QuickTest
qtApp.Visible = True 										' Make the QuickTest application visible
Set qtTest = qtApp.Test	
'msgbox "qtTest object created"

'-------------------------------------'
'------- Accessing Excel Sheet -------'
'-------------------------------------'
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True 
Set controllerFile = objExcel.Workbooks.Open( testSuitControllerFile ) 
Set controllerSheet = controllerFile.Sheets("Controller") 
rowCountController = controllerSheet.UsedRange.Rows.Count 
Set configSheet = controllerFile.Sheets("Config")
rowCountConfig = configSheet.UsedRange.Rows.Count

reportingMode = configSheet.cells(2,3).value
attachmentMode = configSheet.cells(2,4).value
strSharepointLink = configSheet.cells(2,5).value


'-----------------'
'---Timer Starts---'
'-----------------'
iStartTime = Now

'--------------------------------------------------'
'------- Running the Scripts From Controller-------'
'--------------------------------------------------'
For intRow = 2 to rowCountController
	If ucase(controllerSheet.cells(intRow,4).value) = "YES" Then
		testScript = controllerSheet.cells(intRow,3).value	
		'msgbox "the test script is "&testScript

		'---------------------------'
		'----Execute Test Scripts---'
		'---------------------------'
		'msgbox "starting execute script"
		ExecuteTestScripts(testScript)	
''msgbox "ending execute script"		

	End If
	
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}" & "!\\.\root\cimv2")
	Set colProcess = objWMIService.ExecQuery ("Select * From Win32_Process")
	For Each objProcess in colProcess
		If LCase(objProcess.Name) = LCase("EXCEL.EXE") OR LCase(objProcess.Name) = LCase("EXCEL.EXE *32") Then
        objProcess.Terminate()
        ' MsgBox "- ACTION: " & objProcess.Name & " terminated"
		End If
	Next
	For Each objProcess in colProcess
		If LCase(objProcess.Name) = LCase("WerFault.exe") OR LCase(objProcess.Name) = LCase("WerFault.exe *32") Then
        objProcess.Terminate()
        ' MsgBox "- ACTION: " & objProcess.Name & " terminated"
		End If
	Next
	
	
	
	
	'-------To Avoid Server Unreachable error message--------'

	Set objExcel = CreateObject("Excel.Application") 
	Set controllerFile = objExcel.Workbooks.Open( testSuitControllerFile ) 
	Set controllerSheet = controllerFile.Sheets("Controller") 

Next

'----------------'
'---Timer Ends---'
'----------------'
iEndTime = now


'--------------------------'
'--------Email Result------'
'--------------------------'
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True 
Set controllerFile = objExcel.Workbooks.Open( testSuitControllerFile ) 
Set configSheet = controllerFile.Sheets("Config")
strURLtoLog = configSheet.cells(2,1).value
rowCountConfig = configSheet.UsedRange.Rows.Count

For intRow = 2 to rowCountConfig
	If intRow  = 2	Then
		strMailTO = configSheet.cells(intRow,1).Value
		
	Else
		strMailTO = strMailTO & ";" & configSheet.cells(intRow,1).Value
		
	End If
Next

For intRow = 2 to rowCountConfig
	If intRow  = 2	Then
		strMailCC = configSheet.cells(intRow,2).Value
		
	Else
		strMailCC = strMailCC & ";" & configSheet.cells(intRow,2).Value
		
	End If
Next


'------------------------------------------------------------------'
'-------Call Function to Create Test Suite Summary html------------'
'------------------------------------------------------------------'
CreateTestSuiteSummary




'---------------------------------------------'
'------------Attachment Mode Choose-----------'
'---------------------------------------------'
IF attachmentMode="Executive Summary Report" Then
	attachmentData = strSummaryReportFile
ElseIf attachmentMode="Zip Report (Summary + Individual Test Case Report)" Then
	ExecuteZip("ZipCode.vbs")
	attachmentData = testDir & "\" & strZipWithDateTimeStamp
Else
	attachmentData = strSummaryReportFile
End If

'---------------------------------------------'
'------------Reporting Mode Choose------------'
'---------------------------------------------'
IF UCASE(reportingMode)="MAIL" Then
	''Send Test Summary Report as Email attachment
	SendMail (attachmentData)
ElseIf UCASE(reportingMode)="SHAREPOINT" Then
	''Sharepoint upload
	SharepointUpload(strSharepointLink)
ElseIf UCASE(reportingMode)="BOTH" Then
	''Send Test Summary Report as Email attachment
	SendMail (attachmentData)
	''Sharepoint upload
	SharepointUpload(strSharepointLink)
ElseIf UCASE(reportingMode)="NONE" Then
	''Do nothing
Else
	''DEFAULT Send Test Summary Report as Email attachment
	SendMail (attachmentData)
End If	

'----------------------------------------'
'------------Releasing Objects-----------'
'----------------------------------------'

	Set controllerFile  		= Nothing 	' Release testSuitControllerFile Object
	Set controllerSheet 		= Nothing 	' Release the controller Sheet  object
	Set configSheet 			= Nothing 	' Release the config Sheet object
	Set objExcel				= Nothing
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}" & "!\\.\root\cimv2")
	Set colProcess = objWMIService.ExecQuery ("Select * From Win32_Process")
	For Each objProcess in colProcess
		If LCase(objProcess.Name) = LCase("EXCEL.EXE") OR LCase(objProcess.Name) = LCase("EXCEL.EXE *32") Then
        objProcess.Terminate()
        ' MsgBox "- ACTION: " & objProcess.Name & " terminated"
		End If
	Next
	For Each objProcess in colProcess
		If LCase(objProcess.Name) = LCase("WerFault.exe") OR LCase(objProcess.Name) = LCase("WerFault.exe *32") Then
        objProcess.Terminate()
        ' MsgBox "- ACTION: " & objProcess.Name & " terminated"
		End If
	Next
'-----------------------------------------'
'------------Script END Notification------'
'-----------------------------------------'
Msgbox("Your Test Suite has been EXECUTED")
WScript.Quit





' ================================================================================================
'  NAME			: ExecuteTestScripts
'  DESCRIPTION 	  	: This invokes QTP, opens & executes the test. This would call functions to write the summary report & email the samel
'  PARAMETERS		: testScript - Test script name to be executed
'
' ================================================================================================

Public Sub ExecuteTestScripts(testScript)

	test = testDir  & "\" & "Scripts" & "\" & testScript
	qtApp.Open test, False 							' Open the test in read-only mode
	
	' If Test does not exist then log error and execute next script
	If (Err.Description <> "Cannot open test.") Then
		On Error Goto 0
		qtApp.Test.Settings.Launchers("Web").CloseOnExit = False
		
		Set qtLibraries = qtApp.Test.Settings.Resources.Libraries ' Get the libraries collection object
		qtLibraries.Add testDir & "\Library\CommonLib.vbs", 1 ' Add the library to the collection
		qtLibraries.Add testDir & "\Library\AppLib.vbs", 2 ' Add the library to the collection
		qtLibraries.Add testDir & "\Library\ReportLib.vbs", 3 ' Add the library to the collection
	    Wscript.sleep 1000		
		On error resume next
		
		qtApp.Test.Run, True
		Wscript.sleep 1000
		
		If Err.Number Then   
			WScript.Sleep 200   
			Err.Clear   
			qtApp.Test.Run , True    
			Counter = 0
			FLag = 0
			Do while Flag = 0 AND Counter < 5
				If Err.Number then   
					Err.Clear   
					qtApp.Test.Run , True  
					Counter = Counter + 1
					WScript.Sleep 200 					
				Else
				    Flag = 1
				End if			
			Loop

			WScript.Sleep 200
		End If
		
		iTCExecuted = iTCExecuted + 1
		iPTCExecuted = iPTCExecuted + 1
		'sTCRunStatus = qtApp.Test.LastRunResults.Status		' Get Test Script Result Status
		
		Set fso=createobject("Scripting.FileSystemObject")
		'Open the file "TCStatus.txt" in reading mode.
		Set qfile=fso.OpenTextFile(TCStatus,1,True)
		'Read  characters from the file
		sTCRunStatus = qfile.Read(6)                 
		
		Set objContExcel = CreateObject("Excel.Application")
		Set objWorkbook = objContExcel.Workbooks.Open(testSuitControllerFile)
		If (sTCRunStatus = "Failed") Then
			iTCFailCount = iTCFailCount + 1
			objContExcel.cells(intRow,5).value = objContExcel.cells(intRow,6).value 'prev status 
			objContExcel.cells(intRow,6).value = "FAIL"
			If UCASE(objContExcel.cells(intRow,5).value) = "PASS" Then
				iPTCPassCount = iPTCPassCount + 1
			ElseIf UCASE(objContExcel.cells(intRow,5).value) = "FAIL" Then
				iPTCFailCount = iPTCFailCount + 1
			End If
						
		ElseIf (sTCRunStatus = "Passed") Then
			iTCPassCount = iTCPassCount + 1
			objContExcel.cells(intRow,5).value = objContExcel.cells(intRow,6).value 'prev status 
			objContExcel.cells(intRow,6).value = "PASS"
 			If UCASE(objContExcel.cells(intRow,5).value) = "PASS" Then
				iPTCPassCount = iPTCPassCount + 1
			ElseIf UCASE(objContExcel.cells(intRow,5).value) = "FAIL" Then
				iPTCFailCount = iPTCFailCount + 1
			End If
		Else
			objContExcel.cells(intRow,5).value = objContExcel.cells(intRow,6).value 'prev status 
			objContExcel.Cells(intRow,6).Value = sTCRunStatus
		End If
		qtApp.Test.Close	' Close Test Script
	Else
		Set objContExcel = CreateObject("Excel.Application")
		Set objWorkbook = objContExcel.Workbooks.Open(testSuitControllerFile)
		iTCExecuted = iTCExecuted + 1
		iTCFailCount = iTCFailCount + 1
		objContExcel.Cells(intRow, 6).Value = Err.Description & " " & test
		On Error Goto 0
		qtApp.Test.Close	' Close Test Script		
	End If
	'Release the allocated objects
	'Close the files
	qfile.Close
	Set qfile=nothing
	Set fso=nothing
	objWorkbook.Save	' Save Controller File
	objWorkbook.Close '  Close the excel report
	objContExcel.Quit					' Quit Excel Object
	Set objContExcel = Nothing
	Set objWorkbook = Nothing
	qtApp.Close
	qtApp.Quit
End Sub


' ================================================================================================
'  NAME			: CreateTestSuiteSummary
'  DESCRIPTION 	  	: This function is to create a test Suite summary report
'  PARAMETERS		: sControllerFile - test controller file
' ================================================================================================

Public Function CreateTestSuiteSummary ()


	Set App = CreateObject("QuickTest.Application")
	'strGblTestName = App.Test.Name
	Dim objNet
	' create a network object
	Set objNet = CreateObject("WScript.NetWork")
	' show the user name
	strGlbUserName = objNet.UserName 
	' show the computer name
	strGlbComputerName = objNet.ComputerName
	' show the domain name
	'strGlbUserDomain = objNet.UserDomain		
	' destroy the object
	Set objNet = Nothing 
	
	
	
	
	Set fso1 = CreateObject("Scripting.FileSystemObject")
	strSummaryReportFile = testDir & "\" & "TestResults" & "\" &  "TestSuiteExecutionSummary" & "_" & Month(Now) & "-" & Day(Now) & "-" & Year(Now) & "_" & Hour(Now) & "-" & Minute(Now) & "-" & Second(Now) & ".html"
	Set MyFile = fso1.CreateTextFile(strSummaryReportFile,True)
	Set objRepExcel = CreateObject("Excel.Application")
	Set objRepWorkbook = objRepExcel.Workbooks.Open(testSuitControllerFile)
    Set objRepControllerSheet = objRepExcel.Sheets("Controller")
	iRowNum = 2
	
	' WRITE HEADER
	'Pie Chart'
	MyFile.WriteLine("<html><head>  <script type='text/javascript' src='https://www.gstatic.com/charts/loader.js'></script>")
    MyFile.WriteLine("<script type='text/javascript'>")
    MyFile.WriteLine("  google.charts.load('current', {packages:['corechart']}); ")
    MyFile.WriteLine("  google.charts.setOnLoadCallback(drawChart); ")
    MyFile.WriteLine("  function drawChart() { ")
    MyFile.WriteLine("    var data = google.visualization.arrayToDataTable([ ")
    MyFile.WriteLine("      ['Status', 'No.'], ")
    MyFile.WriteLine("      ['Pass',     " & iTCPassCount & "],")
    MyFile.WriteLine("      ['Fail',      " & iTCFailCount & "], ")
   ' MyFile.WriteLine("      ['Un-Executed',  2], ")
   ' MyFile.WriteLine("      ['Blocked', 2],")
    MyFile.WriteLine("    ]); ")
    MyFile.WriteLine("    var options = { ")
    MyFile.WriteLine("      title: 'Test Summary', ")
    MyFile.WriteLine("      is3D: true, ")
    MyFile.WriteLine("    }; ")
    MyFile.WriteLine("    var chart = new google.visualization.PieChart(document.getElementById('piechart_3d')); ")
    MyFile.WriteLine("    chart.draw(data, options); }")
 
    MyFile.WriteLine("</script>  <title>Test Suite Summary Report</title></head>")
	MyFile.WriteLine("<body><div align='center'><center>")

	
	'Outer table'
	MyFile.WriteLine("<table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' bordercolor='#111111' width='100%' height='116'><tr><td width='70%'>")
	
	'Header
	MyFile.WriteLine("<table border='1' cellpadding='0' cellspacing='0' style='border-collapse: collapse' bordercolor='#111111' width='100%' height='116'>")
	MyFile.WriteLine("<tr><td width='100%' colspan='2' bgcolor='#828385' height='36'>")
	MyFile.WriteLine("<p align='center'><font face='Verdana' size='5' color='#FFFFFF'>Test Suite Summary Report</font></td></tr>")
	'Date and Time of Execution
	MyFile.WriteLine("<tr><td width='30%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>Execution Date &amp; Time</font></b></td>")
	MyFile.WriteLine("<td width='70%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>" & FormatDateTime (Date, 1) & " " & FormatDateTime (Time, 0) & "</font></td></tr>")

	''''URL
	''MyFile.WriteLine("<tr><td width='19%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>Test URL</font></b></td>")
	''MyFile.WriteLine("<td width='81%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>"& strURLtoLog &"</font></td></tr>")
	
	'User Name
	MyFile.WriteLine("<tr><td width='30%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>Executed By</font></b></td>")
	MyFile.WriteLine("<td width='70%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>" & ucase(strGlbUserName) & "</font></td></tr>")
	'Machine Name
	MyFile.WriteLine("<tr><td width='30%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>Executed Machine</font></b></td>")
	MyFile.WriteLine("<td width='70%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>" & strGlbComputerName & "</font></td></tr>")
	
	''Test Case Executed
	MyFile.WriteLine("<tr><td width='30%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>Test Case (s) Executed</font></b></td>")	
	MyFile.WriteLine("<td width='70%' height='25'><p style='margin-left: 5'>" & iTCExecuted & "</td></tr>")
	
	''Test Case PASS
	MyFile.WriteLine("<tr><td width='30%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>Test Case (s)<font color='#05A251'>PASS</font></font></b></td>")
	MyFile.WriteLine("<td width='70%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>" & iTCPassCount & "</font></td></tr>")
	
	
	''Test Case FAIL
	MyFile.WriteLine("<tr><td width='30%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>Test Case (s)<font color='#FF0000'>FAIL</font></font></b></td>")
	MyFile.WriteLine("<td width='70%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>" & iTCFailCount & "</font></td></tr>")

		''Test Case PASS
	MyFile.WriteLine("<tr><td width='30%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>Last Run Test Case (s)<font color='#05A251'>PASS</font></font></b></td>")
	MyFile.WriteLine("<td width='70%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>" & iPTCPassCount & "</font></td></tr>")
	
	
	''Test Case FAIL
	MyFile.WriteLine("<tr><td width='30%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>Last Run Test Case (s)<font color='#FF0000'>FAIL</font></font></b></td>")
	MyFile.WriteLine("<td width='70%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>" & iPTCFailCount & "</font></td></tr>")

	MyFile.WriteLine("</table></center></div>")
	MyFile.WriteLine("<p style='margin-left: 5'>&nbsp;</p> </td>")
	
	'Pie Chart'
	MyFile.Writeline("<td width='30%'><div id='piechart_3d' style='width: 500px; height: 300px;' align='Right'></div> </td></tr></table>")
	
	''WRITE TEST SCRIPT STATUS
	MyFile.WriteLine("<table border='1' cellpadding='0' cellspacing='0' style='border-collapse: collapse' bordercolor='#111111' width='100%' height='89'>")
	MyFile.WriteLine("<td width='7%' height='25'align='center' bgcolor='#828385'><b><font face='Verdana' size='2'>Sl. No.</font></b></td>")
	MyFile.WriteLine("<td width='50%' height='29' align='center' bgcolor='#828385'><b><font face='Verdana' size='2'>Test Script Name</font></b></td>")
	MyFile.WriteLine("<td width='10%' height='29' align='center' bgcolor='#828385'><b><font face='Verdana' size='2'>Previous Status</font></b></td>")
	MyFile.WriteLine("<td width='10%' height='29' align='center' bgcolor='#828385'><b><font face='Verdana' size='2'>Present Status</font></b></td></tr>")
		
	Do Until objRepControllerSheet.Cells(iRowNum,1).Value = ""
		If UCase(TRIM(objRepControllerSheet.Cells(iRowNum,4).Value)) = "YES" Then
			 testScript     = objRepExcel.Cells(iRowNum, 3).Value
			 sPTCStatus      = objRepExcel.Cells(iRowNum, 5).Value
			 sTCStatus      = objRepExcel.Cells(iRowNum, 6).Value
			'Sl No
			
			MyFile.WriteLine("<td width='7%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>" & iRowNum - 1 & "</font></td>")
			' TEST SCRIPT NAME
			MyFile.WriteLine("<td width='50%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>" & testScript & "</font></td>")
			' PREVIOUS STATUS
			If UCase(sPTCStatus) = "" Then
				MyFile.WriteLine("<td width='10%' height='25'><p style='margin-center: 5' align='center'><font face='Verdana' size='2'> &nbsp; Not Executed </font></td>")
			ElseIf (UCase(sPTCStatus) = "FAIL") Then 
				MyFile.WriteLine("<td width='10%' height='25'><p style='margin-center: 5' align='center'><font face='Verdana' size='2' color='#FF0000'>&nbsp;" & sPTCStatus & "</font></td>")
			ElseIf (UCase(sPTCStatus) = "PASS") Then 
				MyFile.WriteLine("<td width='10%' height='25'><p style='margin-center: 5' align='center'><font face='Verdana' size='2' color='#05A251'>&nbsp;" & sPTCStatus & "</font></td>")
			Else 
				MyFile.WriteLine("<td width='10%' height='25'><p style='margin-center: 5' align='center'><font face='Verdana' size='2'>&nbsp;" & sPTCStatus & "</font></td>")
			End If
			' STATUS
			If UCase(sTCStatus) = "" Then
				MyFile.WriteLine("<td width='10%' height='25'><p style='margin-center: 5' align='center'><font face='Verdana' size='2'> &nbsp; Not Executed </font></td></tr>")
			ElseIf (UCase(sTCStatus) = "FAIL") Then 
				MyFile.WriteLine("<td width='10%' height='25'><p style='margin-center: 5' align='center'><font face='Verdana' size='2' color='#FF0000'>&nbsp;" & sTCStatus & "</font></td></tr>")
			ElseIf (UCase(sTCStatus) = "PASS") Then 
				MyFile.WriteLine("<td width='10%' height='25'><p style='margin-center: 5' align='center'><font face='Verdana' size='2' color='#05A251'>&nbsp;" & sTCStatus & "</font></td></tr>")
			Else 
				MyFile.WriteLine("<td width='10%' height='25'><p style='margin-center: 5' align='center'><font face='Verdana' size='2'>&nbsp;" & sTCStatus & "</font></td></tr>")
			End If
		End If
		iRowNum = iRowNum + 1	' Increment Excel Row count	
	Loop
		
	MyFile.WriteLine("</table>")
	MyFile.WriteLine("<p>&nbsp;</p>")
	MyFile.WriteLine("<table border='1' cellpadding='0' cellspacing='0' style='border-collapse: collapse' bordercolor='#111111' width='36%' align='Left'>")
	''WRITE START / END TIMES
	''START TIME	
	MyFile.WriteLine("<tr><td width='25%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>Start Time</font></b></td>")
	MyFile.WriteLine("<td width='75%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>&nbsp;" & iStartTime & "</font></td></tr>")
	
	''END TIME
	MyFile.WriteLine("<tr><td width='25%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>End Time</font></b></td>")
	MyFile.WriteLine("<td width='75%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>&nbsp;" & iEndTime &"</font></td></tr>")
	MyFile.WriteLine("</table></body></html>")
	objRepWorkbook.Close
	objRepExcel.Quit      
	Set objRepWorkbook = Nothing	' Quit Excel Object
	Set objRepExcel = Nothing
	
End Function



' ================================================================================================
'  NAME			: SendMail
'  DESCRIPTION 	  	: Send mail with the Summary report attached
'  PARAMETERS		: strSummaryReportFile - 
' ================================================================================================

Public Function SendMail (strSummaryReportFile)
	Dim OutlookApplication
	Dim objMail
	'strMailSub  = "Test Execution Results Report: " & RepDate & "  " & RepTime & " for Scenario: " & strScenario
	'strMailBody = "Please find enclosed the result report of automated execution of the script for the Scenario '" & Ucase(strScenario) & "'  "
	strMailSubject 	= "Test Automation Execution Report : " & FormatDateTime (Date, 1) & " " & FormatDateTime (Time, 0)
	strMailBody		= "Please find attached Test Automation Execution Report"
	Set OutlookApplication = CreateObject("Outlook.Application")
	Set objMail = OutlookApplication.CreateItem(olMailItem)

	
	With objMail
			.To 		= strMailTO
			.CC 		= strMailCC 
			.Attachments.add  strSummaryReportFile
           	.Subject 	= strMailSubject
			.Body 		= strMailBody			
			.Send
	End With

	Set objMail 			= Nothing
	Set OutlookApplication 	= Nothing
End Function

' ================================================================================================
'  NAME			    : SharepointUpload
'  DESCRIPTION 	  	: Uploads the Test Results in Sharepoint
'  PARAMETERS		: strSharepointLink
' ================================================================================================

public Function SharepointUpload(strSharepointLink)
	dim filesys
	set filesys=CreateObject("Scripting.FileSystemObject")
	'msgbox attachmentData
	If filesys.FileExists(attachmentData) Then
		newFolderName = "Test"
		'msgbox strSharepointLink & "\" & newFolderName 
		'filesys.CreateFolder( strSharepointLink & "\" & newFolderName )
		'msgbox strSharepointLink & "\" & newFolderName & "\" 
		filesys.CopyFile attachmentData , strSharepointLink & "\" '& newFolderName & "\" 
	End If
	set flesys = nothing
End Function

' ================================================================================================
'  NAME			    : ExecuteZip
'  DESCRIPTION 	  	: Zipps the Result folder
'  PARAMETERS		: strFileName
' ================================================================================================

Public Function ExecuteZip(strFileName)
	Dim objShell
	Set objShell = Wscript.CreateObject("WScript.Shell")
	objShell.Run strFileName
	Set objShell = Nothing
	d=now()
	Dim arr(6)
	arr(0)="ZippedTestResult"
	arr(1)=DatePart("d",d)
	arr(2)=DatePart("m",d)
	arr(3)=DatePart("yyyy",d)
	arr(4)=DatePart("h",d) & "h"
	arr(5)=DatePart("n",d) & "m"
	arr(6)=DatePart("s",d)	& "s"
Select case arr(2)
   case 1
	arr(2) = "Jan"
   case 2
	arr(2) = "Feb"
   case 3
	arr(2) = "Mar"
   case 4
	arr(2) = "Apr"
   case 5
	arr(2) = "May"
   case 6 
	arr(2) = "Jun"
   case 7
	arr(2) = "Jul"
   case 8
	arr(2) = "Aug"
   case 9
	arr(2) = "Sept"
   case 10
	arr(2) = "Oct"
   case 11
	arr(2) = "Nov"
   case 12
	arr(2) = "Dec"
   case else
End select
	strZipWithDateTimeStamp = Join(arr,"_") & ".zip"
	Dim Fso
	Wscript.Sleep(1600)
	Set Fso = WScript.CreateObject("Scripting.FileSystemObject")
	Fso.MoveFile  testDir & "\" & "Zipped_Test_Result.zip" , strZipWithDateTimeStamp
	' Using Set is mandatory
	Set Fso = nothing
End Function

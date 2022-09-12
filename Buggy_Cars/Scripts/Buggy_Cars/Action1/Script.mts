'===========================================================================================================
'Description			: Launch
'Created By			: Surbhi Kaushik
'Created On			: 09/09/2022
'Last Updated On	: 
'Notes				:
'Parameters		: Refer to input / Output Properties
'===========================================================================================================

'******************************************************************************************

sTestDir = Environment.Value ("TestDir")  'Finding Script UFT file directory
arrPath = Split(sTestDir, "\")
arrPath(UBound(arrPath)-1) = "Library"
For I=0 to UBound(arrPath) -1
	If (I=0) Then
		strLibPath = arrPath(I)
	Else
		strLibPath = strLibPath + "\" + arrPath(I)
	End If
Next

strCommLibPath = strLibPath & "\" & "CommonLib.vbs"
strAppLibPath = strLibPath & "\" & "AppLib.vbs"
strReportLibPath = strLibPath & "\" & "ReportLib.vbs"

Set qtApp = CreateObject("QuickTest.Application") 			' Create the Application object
Set qtLibraries = qtApp.Test.Settings.Resources.Libraries   ' Get the libraries collection object

'If qtLibraries.Find(strCommLibPath) = -1 Then               ' returns 1 if lib is associated else -1
'	ExecuteFile strCommLibPath 
'End If
'If qtLibraries.Find(strAppLibPath) = -1 Then
'	ExecuteFile strAppLibPath
'End If
'If qtLibraries.Find(strReportLibPath) = -1 Then
'	ExecuteFile strReportLibPath
'End If
Set qtApp = Nothing
Set qtLibraries = Nothing
''''''''''''''''''create global variables & Loads common repository
Initialization ()

''''''''''''''''''Call function to create HTML test script execution result.
CreateResultFile()

'*************************************************************************************************



WebUtil.DeleteCookies 								'Delete cookies
SystemUtil.CloseProcessByName("iexplore.exe") 	'Close any existing IE Browser
SystemUtil.CloseProcessByName("chrome.exe") 		'Close any existing Chrome Browser
SystemUtil.CloseProcessByName("firefox.exe")		'Close any existing firefox Browser
SystemUtil.CloseProcessByName("msedge.exe")		'Close any existing Edge Browser

'Dim mode_Maximized, mode_Minimized
'mode_Maximized = 3 'Open in maximized mode
'mode_Minimized = 2 'Open in minimized mode

SELECT_BROWSER  = parameter.Item("Input_Browser")

Select Case SELECT_BROWSER
	
	Case "CHROME"
		SystemUtil.Run "chrome.exe", Parameter("Input_Link") , , ,mode_Maximized 'Launch Chrome
		Reporter.ReportEvent micDone, "Launch Browser","Google Chrome is now open"

	Case "IE"
		SystemUtil.Run "iexplore.exe", Parameter("Input_Link") , , ,mode_Maximized 'Launch Internet Explorer
		Reporter.ReportEvent micDone, "Launch Browser","Internet Explorer is now open"
	
	Case "FIREFOX"
		SystemUtil.Run "firefox.exe", Parameter("Input_Link") , , ,mode_Maximized 'Launch Firefox
		Reporter.ReportEvent micDone, "Launch Browser","Firefox is now open"
		
	Case "EDGE"
		SystemUtil.Run "msedge.exe", Parameter("Input_Link") , , ,mode_Maximized 'Launch Firefox
		Reporter.ReportEvent micDone, "Launch Browser","Firefox is now open"
		
End Select

Wait(10)

Dim obj
Dim BrowserName
Dim PageTitle
Dim SAPPortal
Call BROWSERPROPERTIES(BrowserName)
Call PAGEPROPERTIES(PageTitle,SAPPortal)

If PageTitle <> "" Then
	Set obj=Browser("name:="&BrowserName).Page("title:="&PageTitle) 
End If


If Obj.Exist(30) Then
	Reporter.ReportEvent micPass, "Launch Browser", "Browser is launch successfully"
	LogResult micPass , "Browser Successfully invoked" , "Passed"
	else
	LogResult micFail, "Browser not launched" , "Fail"
End If



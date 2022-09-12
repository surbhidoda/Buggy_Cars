'=======================================================================================================
' Test Script Name	: Register
' Description		: This script will Register into the buggy cars appication
' Created by		: Surbhi Kaushik
' Last modified		:  10/09/2022
' Last modified by	: 
'=======================================================================================================
'-------------------------------------------------------------------------'
'----------------------- TEST SCRIPT STARTS-------------------------------'
'-------------------------------------------------------------------------'



 'create global variables & Loads common repository
Initialization ()
'Click Register Button
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebButton("name:=Register").Click
wait(10)
'Enter the Login text value
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebEdit("name:=username").Click
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebEdit("name:=username").Set "test_1"
'Enter Firstname
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebEdit("name:=firstName").Click
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebEdit("name:=firstName").Set "test_1"
'Enter lastname
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebEdit("name:=lastName").Click
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebEdit("name:=lastName").Set "test_1"
'Enter password
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebEdit("name:=password","index:=1").Click

Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebEdit("name:=password","index:=1").SetSecure "631be64a3ca8695948f05fcecb0fa68eccd8ebf468b22939c04cb2ca"
'Enter confirm password
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebEdit("name:=confirmPassword").Click
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebEdit("name:=confirmPassword").SetSecure "631be64a3ca8695948f05fcecb0fa68eccd8ebf468b22939c04cb2ca"
'Click Register
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebButton("name:=Register","index:=1").Highlight
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebButton("name:=Register","index:=1").Click
wait(20)
If Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebElement("innertext:=Registration is successful ","index:=0").Exist(5) Then
	'reporter.ReportEvent micPass, "Actual order number is not null and is" , "Passed"
	LogResult micPass , "Registration is successful" , "Passed"	
else 
LogResult micFail, "Registration is not successful" , "Failed"
End If









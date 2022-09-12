'=======================================================================================================
' Test Script Name	: Login
' Description		: This script will login into the buggy cars application
' Created by		: Surbhi Kaushik
' Last modified		: 10/09/2022
' Last modified by	: 
'=======================================================================================================
'-------------------------------------------------------------------------'
'----------------------- TEST SCRIPT STARTS-------------------------------'
'-------------------------------------------------------------------------'



 'create global variables & Loads common repository
Initialization ()
'Enter value in login text
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebEdit("placeholder:=Login").Highlight
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebEdit("placeholder:=Login").click
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebEdit("placeholder:=Login").set "test_1"
'Enter value in password text field
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebEdit("name:=password","index:=0").Highlight
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebEdit("name:=password","index:=0").click
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebEdit("name:=password","index:=0").setSecure"631be64a3ca8695948f05fcecb0fa68eccd8ebf468b22939c04cb2ca"
'Click Login button
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebButton("name:=Login").Highlight
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebButton("name:=Login").Click
wait(10)



If Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").Link("name:=Profile").Exist(5) Then
	LogResult micPass , "User has logged in successfully" , "Passed"	
else 
LogResult micFail, "User login is not successful" , "Failed"
End If









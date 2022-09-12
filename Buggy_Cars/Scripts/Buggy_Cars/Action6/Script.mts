'=======================================================================================================
' Test Script Name	: Logout
' Description		: This script will logout the buggy cars application
' Created by		: Surbhi Kaushik
' Last modified		: 10/09/2022
' Last modified by	: 
'=======================================================================================================
'-------------------------------------------------------------------------'
'----------------------- TEST SCRIPT STARTS-------------------------------'
'-------------------------------------------------------------------------'



 'create global variables & Loads common repository
Initialization ()
'Click Logout
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").Link("name:=Logout").Highlight
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").Link("name:=Logout").Click

'Verify that the user has logged out successfully
If Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebButton("name:=Register").Exist(5) Then
	LogResult micPass , "User has logged out successfully" , "Passed"	
else 
LogResult micFail, "User logout is not successful" , "Failed"
End If
testCleanup()





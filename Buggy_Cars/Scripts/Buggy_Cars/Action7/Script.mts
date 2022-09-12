'=======================================================================================================
' Test Script Name	: Add_Vote
' Description		: This script will add a vote to the car and check if the total vote count has increased
' Created by		: Surbhi Kaushik
' Last modified		: 12/09/2022
' Last modified by	: 
'=======================================================================================================
'-------------------------------------------------------------------------'
'----------------------- TEST SCRIPT STARTS-------------------------------'
'-------------------------------------------------------------------------'



 'create global variables & Loads common repository
Initialization ()
'Go to Home
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").Link("name:=Buggy Rating").Highlight
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").Link("name:=Buggy Rating").Click
wait(5)
'Click Alpha romeo Guilia Quadrifoglio
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").Image("title:=Guilia Quadrifoglio").Highlight
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").Image("title:=Guilia Quadrifoglio").Click
wait(10)
'Get the current number of votes for this model
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebElement("innertext:=Votes:.*","html tag:=H4").Highlight
VoteNum=Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebElement("innertext:=Votes:.*","html tag:=H4").GetROProperty("outertext")
SubVot=Right(VoteNum,4)
'msgbox SubVot
wait(5)
oldvotval=SubVot
'Click the vote button
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebButton("name:=Vote.*").Highlight
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebButton("name:=Vote.*").Click
wait(5)
'Validate the increase in the number of votes
Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebElement("innertext:=Votes:.*","html tag:=H4").Highlight
VoteNum=Browser("name:=Buggy Cars Rating").Page("title:=Buggy Cars Rating").WebElement("innertext:=Votes:.*","html tag:=H4").GetROProperty("outertext")
SubVot=Right(VoteNum,4)
currvoteval=SubVot
oldvotval=oldvotval+1
'msgbox oldvotval
If strcomp(oldvotval,currvoteval)=0 Then
	'Reporter.ReportEvent micPass, "Number of votes increased", "Voting functionality worked"
	LogResult micPass , "Number of votes increased" , "Passed"
	LogResult micPass , "The total number of votes is "&currvoteval , "Passed"
	else
	LogResult micFail, "Number of votes did not increase" , "Fail"
End If





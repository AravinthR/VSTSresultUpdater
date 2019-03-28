Dim mTCid
Dim mProjName
Dim mRunId
Dim mResultId
Dim mTestStatus
Dim mReleaseName
Dim mTestCaseName
Dim mTestPlanId
Dim mTestSuiteId
Dim mTestPointId
Dim mUsername
Dim mPassword
Dim mResponse
Dim mRequest
Dim scriptName
Dim Lvar1
Dim result
Dim mAttachmentPath
Dim mSplitTestData
Dim TempPlanID
Dim TempSuiteID
Dim bFlag
Dim suiteplandet
Dim FinalFileName

bFlag = 0
mReleaseName = "APITriggeredResult"

mTestStatus =     wscript.arguments(0) '"FAILED"
mTestCaseName =  wscript.arguments(1)'"TC02.04.01" 
mTestData =    wscript.arguments(2)'"NA$$$NA$$$143" '"PlanID$$$SuiteID$$$TestCaseID"   
mUsername =    wscript.arguments(3) 
mPassword =  wscript.arguments(4) 
FilePath = wscript.arguments(5) '"D:\TestAutomation\"
mSplitTestData = split(mTestData,"$$$")
mTestPlanId = mSplitTestData(0)
mTestSuiteId = mSplitTestData(1)
mTCid = mSplitTestData(2)

Title = mTestCaseName
StepDescription = "One or more Steps may have  Fail"
Priority = "2"
Severity = "3 - Medium"
AssignTo = Replace(mUsername,"\","\\")

LogIt "=-=-=-=-=-=-=-=-=-=-=-=-=-=- NEW TRANSACTION BEGINS =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-"
' =-=-=-=-=-=-=-=- API CALLS URLS=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
mURLgetTestData ="<YOUR URL>/_apis/test/suites?testCaseId=" & mTCid & "&api-version=2.0-preview"
'urlTestPoint = "<YOUR URL>/"& mProjName &"/_apis/test/plans/" & mTestPlanId & "/suites/" & mTestSuiteId & "/points?testcaseid=" & mTCid & "&api-version=1.0"
'urlGetResultId = "<YOUR URL>/"& mProjName &"/_apis/test/runs/"& mRunId &"/results?api-version=1.0"
'urlUpdateTcResult = "<YOUR URL>/"& mProjName &"/_apis/test/runs/"& mRunId &"/results?api-version=2.0-preview"
'urlUpdateTestRun = "<YOUR URL>/"& mProjName &"/_apis/test/runs/"& mRunId &"?api-version=1.0"
' =-=-=-=-=-=-=-=- API CALLS URLS ENDS =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

if mTestPlanId <> "NA" or mTestSuiteId <> "NA" then
	LogIt "all information provided. Hence retrieving only project Name"
	LogIt "Calling SUITE ID NORMAL"
	suiteplandet = getPlanSuiteDataNormal(mURLgetTestData)
	Lvar1 = split(suiteplandet,"|")
	mProjName = Lvar1(2)
	bFlag = 1
else
	LogIt "no information provided. Hence retrieving ALL using TCid"
	LogIt "Calling SUITE ID"
	suiteplandet = getPlanSuiteData(mURLgetTestData)
	if suiteplandet <> "NA" then
		Lvar1 = split(suiteplandet,"|")
		mTestSuiteId= Lvar1(0)
		mTestPlanId= Lvar1(1)
		mProjName = Lvar1(2)
		bFlag = 1
	else 
		LogIt "more than 1 occurrences detected for this test case :" & mTCid &" in the PLAN ID. Hence, skipping this test case from updating to VSTS"
		bFlag = 0
	end if
end if

if bFlag = 1 then
	'LogIt "testPlan ID = " & mTestPlanId 
	'LogIt "Test Suite ID = " & mTestSuiteId
	'LogIt "Project name = " &mProjName
	' =-=-=-=-=-=-=-=- API CALLS URLS=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
	'mURLgetTestData ="<YOUR URL>/_apis/test/suites?testCaseId=" & mTCid & "&api-version=2.0-preview"
	
	
	
	
	' =-=-=-=-=-=-=-=- API CALLS URLS ENDS =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

	LogIt "======== Sending TC ID and URL to getrunID function ========"
	urlTestPoint = "<YOUR URL>/"& mProjName &"/_apis/test/plans/" & mTestPlanId & "/suites/" & mTestSuiteId & "/points?testcaseid=" & mTCid & "&api-version=1.0"
	mRunId = getrunid(urlTestPoint,mReleaseName & "$" & mTestCaseName)
	LogIt "run ID received from the function is = " &mRunId
	LogIt "======== Sending run ID to gettcresultid function ========"
	urlGetResultId = "<YOUR URL>/"& mProjName &"/_apis/test/runs/"& mRunId &"/results?api-version=1.0"
	mResultId = gettcresultid(urlGetResultId)
	LogIt "response received = " &mResultId
	LogIt "======== Sending run ID to updatetestresult function ========"
	urlUpdateTcResult = "<YOUR URL>/"& mProjName &"/_apis/test/runs/"& mRunId &"/results?api-version=2.0-preview"
	updateResponse = updatetestresult(mTestStatus, urlUpdateTcResult, mResultId)
	LogIt "response received = " &updateResponse 
	
	LogIt "======== Sending run ID to updaterun function ========"
	urlUpdateTestRun = "<YOUR URL>/"& mProjName &"/_apis/test/runs/"& mRunId &"?api-version=1.0"
	updaterun urlUpdateTestRun
	'FinalUpload =uploadFile(FinalFileName , Filename,mRunId)
	ResultsFolder = LatestFolder(FilePath &"\Results")
	SuiteFolder =LatestFolder(ResultsFolder)
	ResultPath = SuiteFolder & "\HTML Results\"
	Filename =  LatestFile(ResultPath)
	Stream = Base64Conv(Filename)
	Test = RunUpload(FinalFileName ,Stream,mRunId)
	Set filesys = CreateObject("Scripting.FileSystemObject") 
	filesys.DeleteFile Stream 
	LogIt "===================== THE END =================================" 
else
	LogIt "The API response returned more than 1 occurrences of the test case"
	LogIt "JSON resp received := " & mResponse
	LogIt "===================== THE END =================================" 
end if

function getPlanSuiteData(Lurl)
	'LogIt "======= Inside getPlanSuiteData function ========="

	dim Lvar1
	dim Lvar2
	dim Lvar3
	dim aFlag
	
	aFlag = 1
	
	mResponse = getDataAPIreq(Lurl)
	'LogIt "Response Received from get call is: " &mResponse
	'LogIt "extracting required parameters from response"
	if instr(mResponse , "},{") > 0 then
		aFlag = 0
	else 
		resparr =split(mResponse ,",")
		Lvar1=split(resparr(0),":")
		LsuiteId = Lvar1(2)
		Lvar2=split(resparr(4),":")
		LprojName = Lvar2(1)
		LprojName = right(LprojName,len(LprojName)-1)
		LprojName = left(LprojName,len(LprojName)-1)
		Lvar3=split(resparr(6),":")
		LplanId = Lvar3(2)
		LplanId = right(LplanId,len(LplanId)-1)
		LplanId = left(LplanId,len(LplanId)-1)
	end if
	if aFlag = 1 then
		'LogIt "====== Exiting getPlanSuiteData function======"
		getPlanSuiteData = LsuiteId & "|" & LplanId & "|" & LprojName
	else 
		getPlanSuiteData = "NA" 
	end if
	
end function
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
function getPlanSuiteDataNormal(Lurl)
	'LogIt "======= Inside getPlanSuiteData function ========="
	dim Lvar1
	dim Lvar2
	dim Lvar3
	
	mResponse = getDataAPIreq(Lurl)
	'LogIt "Response Received from get call is: " &mResponse
	'LogIt "extracting required parameters from response"
	resparr =split(mResponse ,",")
	Lvar1=split(resparr(0),":")
	LsuiteId = Lvar1(2)
	Lvar2=split(resparr(4),":")
	LprojName = Lvar2(1)
	LprojName = right(LprojName,len(LprojName)-1)
	LprojName = left(LprojName,len(LprojName)-1)
	Lvar3=split(resparr(6),":")
	LplanId = Lvar3(2)
	LplanId = right(LplanId,len(LplanId)-1)
	LplanId = left(LplanId,len(LplanId)-1)
	'LogIt "====== Exiting getPlanSuiteData function======"
	getPlanSuiteDataNormal = LsuiteId & "|" & LplanId & "|" & LprojName
end function
'=========================================================================================================
'Getting runid function below
function getrunid(Lurl,Lrelname)

	Dim LjData
	Dim LrunId

		'LogIt "==== INSIDE getrunid FUNCTION===="
		'LogIt "Test run name received is:" & Lrelname
		'LogIt "URL received is : " &Lurl

		'LogIt "========== Sending data to getpointData FUNCTION==========="
	mTestPointId = getPointData(Lurl)
		'LogIt "received response test point ID: " &mtestpointid

	LjData = "{ ""name"": """ & Lrelname & """, ""plan"": { ""id"": """ & mTestPlanId & """ }, ""pointIds"":[ " & mTestPointId & " ] }"
		'LogIt "json Data: " &LjData
	LpostURL = "<YOUR URL>/"& mProjName &"/_apis/test/runs?api-version=2.0-preview"
		'LogIt "=========Sending Data to postData function ==============="
	LrunId =postData(LpostURL, LjData)
		'LogIt "response received = " &mRunId

		'LogIt "============= EXITING RUNID FUNCTION ========="
	getrunid = LrunId
end function
'=========================================================================================================
function getPointData(Lurl)

	'LogIt "====== Inside getpointData function ======="
	dim startPoint, stopPoint
	dim testpoint
	mResponse = getDataAPIreq(Lurl)

	startPoint=instr(mResponse,"{""value"":[{""id"":") + 16
		' 						{"value":[{"id":
	stopPoint=instr(startPoint, mResponse,",""url"":""http:") 
	'LogIt startPoint & "|" & stoppoint & "|"

	'LogIt "getpointdata" & mResponse 
	'LogIt "startpoint" & startPoint
	'LogIt "stoppoint" & stopPoint

	testpoint = mid(mResponse, startPoint, stopPoint-startPoint)
	'testpoint = extractData(mResponse,"{ ""value"": [ {""id"":", ",""url"":""http:", 1)

	'LogIt testpoint & "|"
	'LogIt "====== Exiting getpointdata function ======"

	getPointData = testpoint
end function
'========================================================================================================
function postData(Ldata, Ljsondata)
	dim oWinHttp

	'LogIt "In PostData with url :" & Ldata
	'LogIt "In PostData with json:" & Ljsondata

	contentType ="application/json; charset=UTF-8"
	Set oWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
	oWinHttp.Open "POST", Ldata
	oWinHttp.SetCredentials mUsername, mPassword, HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
	oWinHttp.setRequestHeader "Accept", "application/json"
	oWinHttp.setRequestHeader "Content-Type", contentType
	oWinHttp.Send Ljsondata

	'LogIt oWinHttp.Status & "|" & oWinHttp.StatusText
	'LogIt  oWinHttp.responseText

	resparr =split(oWinHttp.responseText,",")
	runid=split(resparr(0),":")

	'LogIt "runid is: " &runid(1)
	postData =runid(1)
	'LogIt "====== Exiting postData function ======"
end function
'=========================================================================================================
'get testresult id
function gettcresultid(Lurl)
	'LogIt "====Inside getTCresultID function ========"

	Dim Lvar1
	Dim Lvar2
	Dim LresultId

	mResponse = getDataAPIreq(Lurl)
	'LogIt "Notice The Value"
	'LogIt mResponse 
	Lvar1 = split(mResponse,",")
	Lvar2 = split(Lvar1(0),":")
	LresultId = Lvar2(2)

	gettcresultid = LresultId
	'LogIt "=========== Exiting gettcresultID function========"
end function
'========================================================================================================
function getDataAPIreq(Lurl)
	'LogIt "====== inside getData function ======"

	Dim restReq
	Dim response
	
	'LogIt "In GetData with :" & Lurl
	Set restReq = CreateObject("MSXML2.ServerXMLHTTP")
	restReq.open "GET", Lurl, false, mUsername, mPassword
	restReq.send 
	response = restReq.responseText

	'LogIt restReq.status & "|" & restReq.statustext
	'LogIt "Response for get call is :" & response
	getDataAPIreq = response

	'LogIt "====== Exiting getData function ======"
end function
'========================================================================================================
'updaterun results
Function updaterun(Lurl)
	'LogIt "========inside updaterun===============" & runid

	restRequest = "{""logEntries"": [ { ""entryId"": 1, ""dateCreated"": ""2015-05-17 05:00:00"", ""message"": ""Test run started"" }],""state"": ""Completed""}"
	contentType ="application/json; charset=UTF-8"

	Set oWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
	oWinHttp.Open "PATCH", Lurl
	oWinHttp.SetCredentials mUsername,mPassword,HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
	oWinHttp.setRequestHeader "Accept", "application/json"
	oWinHttp.setRequestHeader "Content-Type", contentType
	'LogIt "Sent patch request to update run status"
	oWinHttp.Send restRequest
	response = oWinHttp.StatusText
	'LogIt oWinHttp.Status
	'LogIt response
	'LogIt oWinHttp.ResponseText
	
	'LogIt "======Outside updaterun ========="
end Function
'===============================================================================================================
'updating tc results into run
Function updatetestresult(LtcStatus, Lurl, LtcResultId)
	'LogIt "======= Inside updatetestresult======="

	Dim restRequest
	Dim contentType
	Dim response

	LtcStatus = LtcStatus '&"ed"
	'LogIt "TC Status = " &LtcStatus
	restRequest = "[{""testResult"": { ""id"":"& LtcResultId &"}, ""state"": ""Completed"", ""outcome"":""" & LtcStatus & """, ""comment"": """ & LtcStatus & """}]"
	contentType ="application/json; charset=UTF-8"

	Set oWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
	oWinHttp.Open "PATCH", Lurl
	oWinHttp.SetCredentials mUsername,mPassword,HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
	oWinHttp.setRequestHeader "Accept", "application/json"
	oWinHttp.setRequestHeader "Content-Type", contentType
	oWinHttp.Send restRequest
	response = oWinHttp.StatusText

	'LogIt "updating testresult status and response following"
	'LogIt oWinHttp.Status
	'LogIt response
	'LogIt oWinHttp.ResponseText

	updatetestresult = oWinHttp.Status
if  InStr(LCase(LtcStatus),"fail") then 
BugID = CreateBug(Title,StepDescription,Priority,Severity,AssignTo,mTCid)
ResultsFolder = LatestFolder(FilePath &"\Results")
SuiteFolder =LatestFolder(ResultsFolder)
ResultPath = SuiteFolder & "\HTML Results\"
Filename =  LatestFile(ResultPath)
Upload = uploadFile(FinalFileName , Filename,BugID)
end if 
	
	'LogIt "=======Outside updatetestresult function ========="
end Function
'===============================================================================================================
function extractData(strData, strBeginText, strEndText, iBegin)
dim startPoint, stopPoint
dim strValue
iExtractData = 0
startPoint=instr(iBegin, strData, strBeginText) + len(strBeginText)
stopPoint=instr(startPoint, strData, strEndText) 
strValue= mid(strData, startPoint, stopPoint-startPoint)
iExtractData = stopPoint
'LogIt startPoint & "|" & stoppoint & "|" & strValue
extractData = strValue
end function
'=====================================================================================
sub LogIt(strlog) 
	dim objFSOlog 
	dim outFile 
	Set objShell = CreateObject("Wscript.Shell")
	strPath = objShell.CurrentDirectory
	'On Error Resume Next
	Set objFSOlog = CreateObject("Scripting.FileSystemObject") 
	set outFile = objFSOlog.OpenTextFile(strPath &"\VSTS_resUpdater.log", 8, true)
	outFile.WriteLine Now() & ": " & strlog
	outFile.Close
	set outFile = nothing
	Set objFSOlog = nothing
end sub
'================================================================================================
function CreateBug(Title,StepDescription,Priority,Severity,AssignTo,tcID)
Lurl = "<YOUR URL>/" &mProjName &"/_apis/wit/workitems/$Bug?api-version=1.0"

restRequest = "[{ ""op"":""add"", ""path"": ""/fields/System.Title"", ""value"":""" & Title & """} , {""op"":""add"", ""path"": ""/fields/Microsoft.VSTS.TCM.ReproSteps"", ""value"":""" & StepDescription & """} , {""op"":""add"", ""path"": ""/fields/Microsoft.VSTS.Common.Priority"", ""value"":""" & Priority & """} ,{""op"":""add"", ""path"": ""/fields/Microsoft.VSTS.Common.Severity"", ""value"":""" & Severity & """} ,{""op"":""add"", ""path"": ""/fields/System.AssignedTo"", ""value"":""" & AssignTo & """} ]"			
contentType ="application/json-patch+json; charset=UTF-8"
Set oWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
oWinHttp.Open "PATCH", Lurl
oWinHttp.SetCredentials mUsername,mPassword,HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
oWinHttp.setRequestHeader "Accept", "application/json-patch+json"
oWinHttp.setRequestHeader "Content-Type", contentType
oWinHttp.Send restRequest
response = oWinHttp.responseText
BugIDTrim=instr(response,"{""id"":") +6
BugIDTrimEndPont=instr(BugIDTrim, response,",""rev") 
BugID = mid(response, BugIDTrim, BugIDTrimEndPont - BugIDTrim )
CreateBug =BugID
'======================Below Code will Link the Bug with TestCase===============================

LinkURL = "<YOUR URL>/_apis/wit/workitems/"&BugID &"?api-version=1.0"
LinkRestRequest = "[{""op"":""add"", ""path"": ""/relations/-"", ""value"": { ""rel"":""Microsoft.VSTS.Common.TestedBy-Forward"", ""url"":""<YOUR URL>/_apis/wit/workitems/" & tcID & """, ""attributes"": { ""comment"": ""Making a new link for the dependency""}}}]"   
contentType ="application/json-patch+json; charset=UTF-8"   
Set oWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
oWinHttp.Open "PATCH", LinkURL
oWinHttp.SetCredentials mUsername,mPassword,HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
oWinHttp.setRequestHeader "Accept", "application/json-patch+json"
oWinHttp.setRequestHeader "Content-Type", contentType
oWinHttp.Send LinkRestRequest
status = oWinHttp.status
response = oWinHttp.StatusText
end function 
'================================================================================================
function LatestFolder(FilePath)
Set fs = CreateObject("Scripting.FileSystemObject")
Set MainFolder = fs.GetFolder(FilePath)
For Each fldr In MainFolder.SubFolders
''As per comment
If fldr.DateLastModified > LastDate Or IsEmpty(LastDate) Then
LastFolder = fldr.Name
LastDate = fldr.DateLastModified
End If
Next
LatestFolder =FilePath &"\"&LastFolder
end function
'================================================================================================
function LatestFile (Filepath)
searchFileName = mTestCaseName
Set fso = CreateObject("Scripting.FileSystemObject")  
Set folder = fso.GetFolder(Filepath)  
' Loop over all files in the folder until the searchFileName is found
For each file In folder.Files   
If instr(file.name, searchFileName) Then
FinalFileName = file.name
Exit For
End If
Next
LatestFile = Filepath & FinalFileName
end function 
'=========================Below Code will Upload the Attachment=================
function uploadFile(FinalFileName , Filename,WorkItemID)

LinkURL = "<YOUR URL>/_apis/wit/attachments?fileName="&   FinalFileName &"&api-version=1.0"
Set objStream = CreateObject("ADODB.Stream")
objStream.CharSet = "utf-8"
objStream.Open
objStream.LoadFromFile(Filename)
restRequest = objStream.ReadText()
restRequest = "" & restRequest & ""
'LogIt restRequest
objStream.Close
Set objStream = Nothing
contentType ="application/json"
Set oWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
'LogIt "after url"
oWinHttp.Open "POST", LinkURL
oWinHttp.SetCredentials mUsername,mPassword,HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
oWinHttp.setRequestHeader "Accept", "application/json"
oWinHttp.setRequestHeader "Content-Type", contentType
oWinhttp.Send restRequest
status = oWinHttp.status
response = oWinHttp.responseText
URLTrim=instr(response,",""url"":") +7
URLEndPont= instr(URLTrim, response,"}") 
URL = mid(response, URLTrim, URLEndPont-URLTrim)

'=========================Below Code will link the Attachment to a workItem=================
LinkURL = "<YOUR URL>/_apis/wit/workitems/"&WorkItemID &"?api-version=1.0"
LinkRestRequest = "[{""op"":""add"", ""path"": ""/relations/-"", ""value"": { ""rel"":""AttachedFile"", ""url"":" & URL & "}}]" 
contentType ="application/json-patch+json; charset=UTF-8"   
Set oWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
oWinHttp.Open "PATCH", LinkURL
oWinHttp.SetCredentials mUsername,mPassword,HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
oWinHttp.setRequestHeader "Accept", "application/json-patch+json"
oWinHttp.setRequestHeader "Content-Type", contentType
oWinHttp.Send LinkRestRequest
status = oWinHttp.status
response = oWinHttp.StatusText
end function 
'=====================================================================================
function RunUpload(FinalFileName ,mStream, mRunID)
	Set objStream = CreateObject("ADODB.Stream")
	objStream.CharSet = "utf-8"
	objStream.Open
	objStream.LoadFromFile(mStream)
	streamData = objStream.ReadText()
	streamData = Replace(streamData, vbCr, "")
	streamData = Replace(streamData, vbLf, "")
	restRequest =  "{  ""stream"" : """& streamData &""" ,  ""filename"" : """& FinalFileName &""" ,  ""comment"" : "" Test attachment upload "" ,  ""attachmentType"" :  ""GeneralAttachment"" }"
	restRequest = "" & restRequest & ""
	objStream.Close
	Set objStream = Nothing
	contentType ="application/json; charset=UTF-8"
	Set oWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
	oWinHttp.Open "POST", "<YOUR URL>//_apis/test/runs/"& mRunID &"/attachments?api-version=2.0-preview.1"
	'LogIt "after url"
	oWinHttp.SetCredentials mUsername,mPassword,HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
	oWinHttp.setRequestHeader "Accept", "application/json"
	oWinHttp.setRequestHeader "Content-Type", contentType
	oWinhttp.Send restRequest
	response = oWinHttp.responseText
	LogIt "RUN Attachement URL and ID" &response
end function 
'=====================================================================================
Function Base64Conv(FileName)
Dim inByteArray
dim  base64Encoded
TypeBinary = 1
inByteArray = readBytes(FileName)
Set objShell = CreateObject("Wscript.Shell")
strPath = objShell.CurrentDirectory
base64Encoded = encodeBase64(inByteArray)
Set objFSOlog = CreateObject("Scripting.FileSystemObject") 
	set outFile = objFSOlog.OpenTextFile(strPath &"\ByteStream.txt", 8, true)
	outFile.WriteLine base64Encoded
	Base64Conv = strPath &"\ByteStream.txt"
End Function
'=====================================================================================
private function readBytes(file)
dim inStream
' ADODB stream object used
'LogIt "inside read"
TypeBinary = 1
set inStream = WScript.CreateObject("ADODB.Stream")
' open with no arguments makes the stream an empty container 
inStream.Open
inStream.type= TypeBinary
inStream.LoadFromFile(file)
readBytes = inStream.Read()
'LogIt readBytes
end function
'=====================================================================================
private function encodeBase64(bytes)
set objXML = createobject("MSXML2.DOMDocument.3.0")
Set objNode = objXML.createElement("b64")
objNode.dataType = "bin.base64"
objNode.nodeTypedValue = bytes
encodeBase64 = objNode.Text
end function

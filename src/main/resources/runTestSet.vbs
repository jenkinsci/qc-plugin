' The MIT License
'
' Copyright (c) 2010-2012, Manufacture Française des Pneumatiques Michelin,
' Thomas Maurel, CollabNet, Johannes Nicolai, Shane Smart, Mickael Beluet,
' Romain Seguy, Bhanu Pratap
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
' THE SOFTWARE.

Function addBlankSpaces(str, length)
  Dim strlen
  strlen = length - Len(str)
  If strlen < 0 Then
    addBlankSpaces = Left(str, length - 3) & "..."
  ElseIf strlen > 0 Then
    addBlankSpaces = str & Space(strlen)
  Else
    addBlankSpaces = str
  End If
End Function

Function generateLine(length)
  Dim line
  Dim i
  line = "+"
  For i = 3 To length
    line = line & "-"
  Next
  generateLine = line & "+"
End Function

Sub logMessage(p_szMessage)
  WScript.StdOut.WriteLine Date & " " & Time & " : " & p_szMessage
End Sub

Function prefixWithZero(str, length)
  Dim pre
  If length > len(str) then
    pre = String(length - len(str), "0")
  End If
  prefixWithZero = pre & str
End Function

Function stripTags(str)
  Dim regex
  Set regex = New RegExp
  With regex
    .Pattern = "<|>|&"
    .IgnoreCase = True
    .Global = True
  End With
  stripTags = regex.Replace(str, "")
End Function

' ------------------------------------------------------------------------------

Class QCFailure

  Private fName
  Private fDesc

  Public Property Get Name
    Name = fName
  End Property

  Public Property Get Desc
    Desc = fDesc
  End Property

  Public Property Let Name(iName)
    fName = iName
  End Property

  Public Property Let Desc(iDesc)
    fDesc = iDesc
  End Property

End Class

' ------------------------------------------------------------------------------

Class QCTest

  Private tName         ' test case name
  Private tDuration     ' duration
  Private tStatus       ' current status
  Private tFailure      ' object failure
  Private tFailureDesc  ' failure description

  Public Property Get Duration
    Duration = tDuration
  End Property

  Public Property Get Failure
    Set Failure = tFailure
  End Property

  Public Property Get FailureDesc
    FailureDesc = tFailureDesc
  End Property

  Public Property Get Name
    Name = tName
  End Property

  Public Property Get Status
    Status = tStatus
  End Property

  Public Property Let Name(iName)
    tName = iName
  End Property

  Public Property Let Duration(iDuration)
    tDuration = iDuration
  End Property

  Public Property Set Failure(iFailure)
    Set tFailure = iFailure
  End Property

  Public Property Let FailureDesc(iFailureDesc)
    tFailureDesc = iFailureDesc
  End Property

  Public Property Let Status(iStatus)
    tStatus = iStatus
  End Property

End Class

' ------------------------------------------------------------------------------

Class QCTestRunner

  Private tdConnection
  Private errorMsg
  Private timestamp
  Private tests()
  Private hostName
  Private folder
  Private name
  Private domain
  Private project

  Sub Class_Initialize
    errors = 0
    Set tdConnection = CreateObject("TDApiOle80.TDConnection")
    timestamp = CStr(Now)
  End Sub

  Public Property Get Connected
    On Error Resume Next
    If tdConnection Is Nothing Then
      Connected = False
    Else
      Connected = tdConnection.ProjectConnected
    End If
  End Property

  Public Property Get ErrorMessage
    ErrorMessage = errorMsg
  End Property

  Public Sub ConnectToProject(QCServerURL, QCLogin, QCPass, QCDomain, QCProject)
    On Error Resume Next
    hostName = QCServerURL
    domain = QCDomain
    project = QCProject

    If tdConnection Is Nothing Then
      errorMsg = "Can't create TDConnection Object"
    Else
      tdConnection.InitConnectionEx QCServerURL
      If tdConnection.Connected = False Then
        errorMsg = "Can't connect to server"
      Else
        WScript.StdOut.WriteLine "Connected to server " & QCServerURL
        tdConnection.Login QCLogin, QCPass
        If tdConnection.LoggedIn = False Then
          errorMsg = "Can't login on QC server"
        Else
          WScript.StdOut.WriteLine "Logged in with user " & QCLogin
          tdConnection.Connect QCDomain, QCProject
          If tdConnection.ProjectConnected = False Then
            errorMsg = "Can't open Domain/Project"
          Else
            WScript.StdOut.WriteLine "Opened project " & QCDomain & "\" & QCProject
          End If
        End If
      End If
    End If
  End Sub

  ' runMode: RUN_LOCAL, RUN_REMOTE or RUN_PLANNED_HOST
  Public Sub RunTestSet(tsFolderName, tsName, timeout, runMode, runHost)
    On Error Resume Next
    Dim tsFactory
    Dim tsTestFactory
    Dim tsTreeManager
    Dim tsPath
    Dim tsFolder
    Dim tsList
    Dim tList
    Dim targetTestSet
    Dim tdFilter
    Dim test
    Dim executionStatus
    Dim tsExecutionFinished
    Dim iter
    Dim eventsList
    Dim testExecStatusObj
    Dim qTest
    Dim currentTest
    Dim qFailure

    folder = tsFolderName
    name = tsName

    Set tsFactory = tdConnection.TestSetFactory
    Set tsTreeManager = tdConnection.TestSetTreeManager

    tsPath = "Root\" & tsFolderName

    Set tsFolder = tsTreeManager.NodeByPath(tsPath)
    If tsFolder Is Nothing Then
      errorMsg = "Could not find folder " & tsPath
    Else
      Set tsList = tsFolder.FindTestSets(tsName)
      If tsList.Count < 1 Then
        errorMsg = "Could not find TestSet " & tsName
      Else
        Set targetTestSet = tsList.Item(1)
        WScript.StdOut.WriteBlankLines(1)
        WScript.StdOut.WriteLine generateLine(100)
        WScript.StdOut.WriteLine "| " & addBlankSpaces("TestSet Name", 88) &  " | " & addBlankSpaces("ID", 6) & "|"
        WScript.StdOut.WriteLine generateLine(100)
        WScript.StdOut.WriteLine "| " & addBlankSpaces(tsName, 88) & " | " & addBlankSpaces(targetTestSet.ID, 6) &  "|"
        WScript.StdOut.WriteLine generateLine(100)

        ' start the scheduler
        Set Scheduler = targetTestSet.StartExecution("")

        If Scheduler Is Nothing Then
          errorMsg = "Could not instantiate test set scheduler"
        Else
          Set tsTestFactory = targetTestSet.TSTestFactory
          Set tdFilter = tsTestFactory.Filter
          tdFilter.Filter("TC_CYCLE_ID") = targetTestSet.ID
          Set tList = tsTestFactory.NewList(tdFilter.Text)

          ' set up for the run depending on where the test instances are to execute
          Select Case runMode
            Case "RUN_LOCAL"
              ' run all tests on the local machine
              Scheduler.RunAllLocally = True

            Case "RUN_REMOTE"
              ' run tests on a specified remote machine
              Scheduler.TdHostName = RunHost
              ' RunAllLocally must not be set for remote invocation of tests. As
              ' such, do not do this: Scheduler.RunAllLocally = False

            Case "RUN_PLANNED_HOST"
              ' run on the hosts as planned in the test set
              Scheduler.RunAllLocally = False
          End Select

          WScript.StdOut.WriteLine "| " & addBlankSpaces("Number of tests: " & tList.Count, 97) &  "|"
          WScript.StdOut.WriteLine generateLine(100)
          WScript.StdOut.WriteLine "| " & addBlankSpaces("Test Name", 65) &  " | " & addBlankSpaces("ID", 6) & " | " & addBlankSpaces("Host", 20) &  "|"
          WScript.StdOut.WriteLine generateLine(100)

          ReDim tests(tList.Count - 1)

          i = 1
          For Each test In tList
            Select Case runMode
              Case "RUN_LOCAL"
                WScript.StdOut.WriteLine "| " & addBlankSpaces(test.Name, 65) &  " | " & addBlankSpaces(test.ID, 6) & " | " & addBlankSpaces(RunHost, 20) &  "|"
                WScript.StdOut.WriteLine generateLine(100)
                Scheduler.RunOnHost(test.ID) = runHost

              Case "RUN_REMOTE"
                WScript.StdOut.WriteLine "| " & addBlankSpaces(test.Name, 65) &  " | " & addBlankSpaces(test.ID, 6) & " | " & addBlankSpaces(RunHost, 20) &  "|"
                WScript.StdOut.WriteLine generateLine(100)
                Scheduler.RunOnHost(test.ID) = runHost

              Case "RUN_PLANNED_HOST"
                WScript.StdOut.WriteLine "| " & addBlankSpaces(test.Name, 65) &  " | " & addBlankSpaces(test.ID, 6) & " | " & addBlankSpaces(test.HostName, 20) &  "|"
                WScript.StdOut.WriteLine generateLine(100)
                Scheduler.RunOnHost(test.ID) = test.HostName
            End Select

            ' initialization of the test's default values which is No Run (in
            ' order to handle specific cases in which test are not run)

            Set qTest = New QCTest
            qTest.Name = test.Name
            qTest.Status = "No Run"
            qTest.Duration = 0

            Set qFailure = New QCFailure
            qFailure.Name = "No Run"
            qFailure.Desc = "No Run"

            Set qTest.Failure = qFailure

            Set tests(i - 1) = qTest

            i = i + 1
          Next

          ' tests are actually run
          Scheduler.run
          WScript.StdOut.WriteBlankLines(1)
          WScript.StdOut.WriteLine "Running-Tests..."
          WScript.StdOut.WriteLine "Scheduler started around " & CStr(Now)
          WScript.StdOut.WriteBlankLines(1)
          Set executionStatus = Scheduler.ExecutionStatus

          ' let's wait for the tests to end ("normally" or because of the timeout)
          While ((tsExecutionFinished = False) And (iter < timeout))
            iter = iter + 5
            executionStatus.RefreshExecStatusInfo "all", True
            tsExecutionFinished = executionStatus.Finished

            WScript.StdOut.WriteLine generateLine(100)
            WScript.StdOut.WriteLine "| " & addBlankSpaces(CStr(Now), 97) & "|"
            For i = 1 To executionStatus.Count
              Set testExecStatusObj = executionStatus.Item(i)
              Set currentTest = targetTestSet.TSTestFactory.Item(testExecStatusObj.TSTestId)

              WScript.StdOut.WriteLine "| " & addBlankSpaces(testExecStatusObj.TSTestId, 8) & _
                      addBlankSpaces(currentTest.Name, 70) & _
                      addBlankSpaces(testExecStatusObj.Status, 19) & "|"
            Next
            WScript.StdOut.WriteLine generateLine(100)

            WScript.Sleep(10000)
          Wend

          If iter < timeout Then
            WScript.StdOut.WriteBlankLines(1)
            WScript.StdOut.WriteLine generateLine(100)
            WScript.StdOut.WriteLine "| " & addBlankSpaces("Tests results", 97) &  "|"
            WScript.StdOut.WriteLine generateLine(100)
            WScript.StdOut.WriteLine "| " & addBlankSpaces("Test", 72) &  " | " & addBlankSpaces("Result", 22) & "|"
            WScript.StdOut.WriteLine generateLine(100)

            For i = 1 To executionStatus.Count
              Set testExecStatusObj = executionStatus.Item(i)
              Set currentTest = targetTestSet.TSTestFactory.Item(testExecStatusObj.TSTestId)

              ' we search the id of the test in the tests list in order to update it
              l_id = GetIdTestName(currentTest.Name)
              Set qTest = tests(l_id)

              ' duration and status are updated according to the run
              qTest.Duration = currentTest.LastRun.Field("RN_DURATION")
              qTest.Status = testExecStatusObj.Status

              If instr(1, testExecStatusObj.Status, "Passed") Then
                Set qTest.Failure = Nothing
              Else
                Set qFailure = New QCFailure
                qFailure.Name = testExecStatusObj.Status
                qFailure.Desc = testExecStatusObj.Message
                Set qTest.Failure = qFailure

                ' let's get some more info for addition in the result XML file
                If testExecStatusObj.Status = "FinishedFailed" Then
                  qTest.FailureDesc = GenerateFailedLog(currentTest.LastRun)
                Else
                  qTest.FailureDesc = testExecStatusObj.Status & " : " & testExecStatusObj.Message
                End if
              End If

              Set tests(l_id) = qTest

              WScript.StdOut.WriteLine "| " & addBlankSpaces(currentTest.Name, 72) &  " | " & addBlankSpaces(testExecStatusObj.Status, 22) & "|"
              WScript.StdOut.WriteLine generateLine(100)
            Next

            WScript.StdOut.WriteBlankLines(1)
            WScript.StdOut.WriteLine "Scheduler finished around " & CStr(Now)
          Else
            errorMsg = "Timed out"
          End If

          GenerateDetailedReport(tList)
        End If ' endif scheduler
      End If ' endif test set
    End If ' endif test set folder
  End Sub

  Sub GenerateDetailedReport(objTSTestList)
    WScript.StdOut.WriteLine "Generating detailed report..."
    WScript.StdOut.WriteBlankLines(1)
    WScript.StdOut.WriteLine generateLine(100)

    For i = 1 To objTSTestList.Count
        Set objTest = objTSTestList.Item(i)
        vTestCase = objTest.name
        vStatus = objTest.Status

        WScript.StdOut.WriteLine "| " & addBlankSpaces(vTestCase, 97) & "|"
        WScript.StdOut.WriteLine generateLine(100)

        Set objRun = objTest.LastRun
        iStepCnt = 1
        Set objStepFactory = objRun.StepFactory
        Set objTSTestStepsList = objStepFactory.NewList("")
' cf. detailed comments below
'        vActual = ""

        For Each objStep In objTSTestStepsList
            vDesc = Trim(objStep.Field("ST_DESCRIPTION"))
            vDesc = Replace(vDesc, "<html><body>", "")
            vDesc = Replace(vDesc, "</body></html>", "")
            remain = iStepCnt & ". " & vDesc
            ' if length is longer than 100 chars, split it into multiple lines
            Do While 1
                If Len(remain) > 97 Then
                    line = Left(remain, 97)
                    WScript.StdOut.WriteLine "| " & addBlankSpaces(line, 97) & "|"
                    remain = Right(remain, Len(remain) - 97)
                Else
                    WScript.StdOut.WriteLine "| " & addBlankSpaces(remain, 97) & "|"
                    Exit Do
                End If
            Loop
' cf. detailed comments below
'            ' :> is a pattern that we're nearly sure to not find in the steps name; we'll use it later for splitting the string
'            vActual = vActual & ":>" & objStep.Field("ST_ACTUAL")
            iStepCnt = iStepCnt + 1
        Next

' From Romain:
' The following code comes from https://groups.google.com/d/topic/jenkinsci-dev/xUE-qoL1F2M/discussion
' I'm unsure about what vActual refers to: None of my samples produce an output which could explain
' what we're talking about; If someone knows, please go ahead and uncomment (and fix)!
'        WScript.StdOut.WriteLine "| " & addBlankSpaces("", 97) & "|"
'        WScript.StdOut.WriteLine "| " & addBlankSpaces("Actual:", 97) & "|"
'
'        actualsArray = Split(vActual, ":>", -1, 1)
'        For k = 0 To UBound(actualsArray)
'            If Trim(actualsArray(k)) <> "" Then
'                remain = Trim(actualsArray(k))
'                remain = Replace(remain, "<html><body>", "")
'                remain = Replace(remain, "</body></html>", "")
'                ' if length is longer than 100 chars, split it into multiple lines
'                Do While 1
'                    If Len(remain) > 97 Then
'                        line = Left(remain, 97)
'                        WScript.StdOut.WriteLine "| " & addBlankSpaces(line, 97) & "|"
'                        remain = Right(remain, Len(remain) - 97)
'                    Else
'                        WScript.StdOut.WriteLine "| " & addBlankSpaces(remain, 97) & "|"
'                        Exit Do
'                    End If
'                Loop
'            End If
'        Next

' From Romain: I'm really unsure about what status (take a look at how vStatus is defined) we're
' talking about: It doesn't necessarily match with the actual automated test status...
'        WScript.StdOut.WriteLine "| " & addBlankSpaces("", 97) & "|"
'        WScript.StdOut.WriteLine "| " & addBlankSpaces("Test Result", 47) & " :: " & addBlankSpaces(vStatus, 46) & "|"
        WScript.StdOut.WriteLine generateLine(100)
    Next
  End Sub

  Function GenerateFailedLog(p_Test)
    Set stList = p_Test.StepFactory.NewList("")

    l_szReturn = ""
    l_szFailedMessage = ""

    ' loop on each step in the steps
    For Each Step In stList
      Select Case Step.Status
        Case "Failed"
          l_szFailedMessage = l_szFailedMessage & Step.Field("ST_DESCRIPTION") & vbcrlf
        Case Else
      End Select
    Next

    GenerateFailedLog = l_szFailedMessage
  End Function

  Public Function GetIdTestName(p_szName)
    GetIdTestName = -1

    For i = 0 to Ubound(tests)
      Set qTest = tests(i)

      If qTest Is Nothing Then
        ' do nothing
      Else
        If qTest.Name = p_szName then
          GetIdTestName = i
          Exit For
        End if
      End if
    Next
  End function

  Public Sub Disconnect
    On Error Resume Next
    If tdConnection.ProjectConnected Then
      tdConnection.Disconnect
      If tdConnection.LoggedIn Then
        tdConnection.Logout
        If tdConnection.Connected Then
          tdConnection.ReleaseConnection
          WScript.StdOut.WriteLine "Connection released"
        End If
      End If
    End If
  End Sub

  Public Sub WriteToXML(fileName)
    Dim logShell
    Dim currentDate
    Dim numError
    Dim numFailure
    Dim numTest
    Dim totalTime
    Dim body
    Dim header

    WScript.StdOut.WriteBlankLines(1)
    WScript.StdOut.WriteLine "Generating report..."
    currentDate = YEAR(Date()) & _
            "-" & prefixWithZero(Month(Date()),2) & _
            "-" & prefixWithZero(Day(Date()),2) & _
            "T" & prefixWithZero(Hour(Now()),2) & _
            ":" & prefixWithZero(Minute(Now()),2) & _
            ":" & prefixWithZero(Second(Now()),2)

    WScript.StdOut.WriteLine "Report file path: " & fileName
    WScript.StdOut.WriteBlankLines(1)

    Set logShell = Nothing
    Set objStream = CreateObject("ADODB.Stream" )
    objStream.Open
    objStream.Position = 0
    objStream.Charset = "UTF-8"

    WScript.StdOut.WriteLine generateLine(100)
    WScript.StdOut.WriteLine "| " & addBlankSpaces("Test Name", 72) & " | " & addBlankSpaces("Status", 22) & "|"
    WScript.StdOut.WriteLine generateLine(100)

    totalTime = 0
    numFailure = 0
    numTest = 0
    header = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf
    If Not (errorMsg = "") Then
      numError = 1
      body = vbTab & "<error message=""" & errorMsg & """ type=""fatal"">" & vbCrLf & _
                  vbTab & vbTab & errorMsg & vbCrLf & _
                  vbTab & "</error>"
    Else
      numError = 0
      body = ""
      For Each test In tests
        numTest = numTest + 1

        If test Is Nothing Then
          ' do nothing
        Else
          totalTime = totalTime + test.Duration
          body =  body & vbTab & "<testcase classname=""" & domain & "." & project & "." & stripTags(folder) & "." & stripTags(name) & """ " & _
                  "name=""" & stripTags(test.Name) & """ " & _
                  "time=""" & test.Duration  & ".0"">" & vbCrLf
          lStatus = ""

          If test.Failure Is Nothing Then
            body = body & vbTab & "</testcase>" & vbCrLf
            lStatus = "Passed"

          Elseif test.Status = "No Run" or test.Status = "Condition Failed" Then  ' the test didn't run
            body = body & vbTab & vbTab & "<failure message=""" & test.Status & """ type=""" & test.Status & """>" & vbCrLf & _
                    "<![CDATA[" & test.Status & "]]>" & vbCrLf & "</failure>" & vbCrLf

            body = body & vbTab & "</testcase>" & vbCrLf
            numFailure = numFailure + 1
            lStatus = test.Status
          Else

            body = body & vbTab & vbTab & "<failure message=""" & stripTags(test.Failure.Desc) & """ type=""" & test.Failure.name & """>" & vbCrLf & _
                    "<![CDATA[" & test.Failure.Name & " : " & test.FailureDesc & "]]>" & vbCrLf & _
                    "</failure>" & vbCrLf

            body = body & vbTab & "</testcase>" & vbCrLf
            numFailure = numFailure + 1
            lStatus = test.Failure.Name
          End If

          WScript.StdOut.WriteLine "| " & addBlankSpaces(test.Name, 72) &  " | " & addBlankSpaces(lStatus, 22) & "|"
          WScript.StdOut.WriteLine generateLine(100)

        End if
      Next
    End If

    header = header & "<testsuite errors=""" & numError & """ " & _
            "failures=""" & numFailure & """  " & _
            "hostname=""" & hostName &  """  " & _
            "name=""" & domain & "." & project & "." & stripTags(folder) & "." & stripTags(name) & """  " & _
            "tests=""" & numTest & """ " & _
            "time=""" & totalTime & ".0"" " & _
            "timestamp=""" & currentDate & """>"
    header = header & body
    header = header & "</testsuite>"
    objStream.WriteText header
    objStream.SaveToFile fileName
    objStream.Close
    WScript.StdOut.WriteLine "Report Created"

  End Sub

End Class

' ------------------------------------ Main ------------------------------------

Dim args
Dim test
Dim qcTimeout
Set args = WScript.Arguments
Set test = New QCTestRunner

If args.Count<9 Or args.Count>11 Then

  lszMessage = "Required arguments:" + vbcrlf
  lszMessage = lszMessage + "Arg1 : QC Server" + vbcrlf
  lszMessage = lszMessage + "Arg2 : QC UserName" + vbcrlf
  lszMessage = lszMessage + "Arg3 : QC Password" + vbcrlf
  lszMessage = lszMessage + "Arg4 : QC Domain" + vbcrlf
  lszMessage = lszMessage + "Arg5 : QC Project" + vbcrlf
  lszMessage = lszMessage + "Arg6 : QC TestSetFolder" + vbcrlf
  lszMessage = lszMessage + "Arg7 : QC TestSetName" + vbcrlf
  lszMessage = lszMessage + "Arg8 : XML Junit File" + vbcrlf
  lszMessage = lszMessage + "Arg9 : Timeout" + vbcrlf
  lszMessage = lszMessage + "Arg10: RunMode (RUN_PLANNED_HOST or RUN_REMOTE or RUN_LOCAL -- RUN_PLANNED_HOST if not specified)" + vbcrlf
  lszMessage = lszMessage + "Arg11: RunHost (to be specified when in RUN_REMOTE mode)" + vbcrlf

  WScript.Echo lszMessage
  WScript.Quit 1

Else

  qcServer = args.Item(0)
  qcUser = args.Item(1)
  qcPassword = args.Item(2)
  qcDomain = args.Item(3)
  qcProject = args.Item(4)
  qcTestSetFolder = args.Item(5)
  qcTestSetName = args.Item(6)
  strXmlFile = args.Item(7)
  qcTimeout = args.Item(8)

  logMessage("Script parameters:")
  logMessage("*************************************************")
  logMessage("QC Server       : " & qcServer)
  logMessage("QC UserName     : " & qcUser)
  logMessage("QC Password     : ********")
  logMessage("QC Domain       : " & qcDomain)
  logMessage("QC Project      : " & qcProject)
  logMessage("QC TestSetFolder: " & qcTestSetFolder)
  logMessage("QC TestSetName  : " & qcTestSetName)
  logMessage("XML Junit File  : " & strXmlFile)
  logMessage("Timeout         : " & qcTimeout)
  logMessage("*************************************************")

  ' default execution environment: the planned one
  runMode = "RUN_PLANNED_HOST"
  runHost = ""

  If args.Count >= 10 Then
    runMode = args.Item(9)
    logMessage("RunMode         : " & runMode)

    If runMode = "RUN_PLANNED_HOST" or runMode = "RUN_REMOTE" or runMode = "RUN_LOCAL" then
      If runMode = "RUN_REMOTE" then
        If args.Count > 10 Then
          runHost = args.Item(10)
          logMessage("RunHost         : " & runHost)
        Else
          WScript.StdOut.WriteLine "When RunMode is set to RUN_REMOTE, you must specify the name of the host which will run the tests."
          WScript.Quit 1
        End if
      ElseIf runMode = "RUN_LOCAL" then
        Set WshNetwork = WScript.CreateObject("WScript.Network")
        runHost = WshNetwork.ComputerName
        logMessage("RunHost         : " & runHost)
      End if
    Else
      WScript.StdOut.WriteLine "The RunMode parameter must be RUN_PLANNED_HOST, RUN_REMOTE or RUN_LOCAL."
      WScript.Quit 1
    End if

  End If

  logMessage("*************************************************")

End if

test.ConnectToProject qcServer, qcUser, qcPassword, qcDomain, qcProject
If test.Connected Then
  test.RunTestSet qcTestSetFolder, qcTestSetName, qcTimeout, runMode, runHost
End If

If Not (test.ErrorMessage = "") Then
  WScript.StdOut.WriteLine test.ErrorMessage
  test.WriteToXML strXmlFile
  WScript.Quit 1
End If

test.Disconnect
test.WriteToXML strXmlFile
WScript.Quit 0

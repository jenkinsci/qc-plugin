' The MIT License
'
' Copyright (c) 2010, Manufacture Française des Pneumatiques Michelin, Thomas Maurel,
' CollabNet, Johannes Nicolai
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

Function stripTags(str)
  Dim regex
  Set regex = New RegExp
  With regex
    .Pattern = "<|>"
    .IgnoreCase = True
    .Global = True
  End With
  stripTags = regex.Replace(str, "")
End Function

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

Function prefixWithZero(str, length) 
  Dim pre
  If length > len(str) then
    pre = String(length - len(str), "0")
  End If
  prefixWithZero = pre & str
End Function 

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

Class QCTest 

  Private tName
  Private tDuration
  Private tFailure

  Public Property Get Name
    Name = tName
  End Property

  Public Property Get Duration
    Duration = tDuration
  End Property

  Public Property Get Failure
    Set Failure = tFailure
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
	
End Class

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
      errorMsg = "Cant create TDConnection Object"
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

  Public Sub RunTestSet(tsFolderName, tsName, timeout)
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
        WScript.StdOut.WriteBlankLines(2)
        WScript.StdOut.WriteLine generateLine(50)
        WScript.StdOut.WriteLine "| " & addBlankSpaces("TestSet Name", 38) &  " | " & addBlankSpaces("ID", 6) & "|"
        WScript.StdOut.WriteLine generateLine(50)
        WScript.StdOut.WriteLine "| " & addBlankSpaces(tsName, 38) &          " | " & addBlankSpaces(targetTestSet.ID, 6) &  "|"
        WScript.StdOut.WriteLine generateLine(50)

        Set Scheduler = targetTestSet.StartExecution("")
	    If Scheduler Is Nothing Then
	      errorMsg = "Could not instantiate test set scheduler"
   	    Else
          Set tsTestFactory = targetTestSet.TSTestFactory
          Set tdFilter = tsTestFactory.Filter
          tdFilter.Filter("TC_CYCLE_ID") = targetTestSet.ID
          Set tList = tsTestFactory.NewList(tdFilter.Text)
          WScript.StdOut.WriteBlankLines(2)
          WScript.StdOut.WriteLine generateLine(50)
          WScript.StdOut.WriteLine "| " & addBlankSpaces("Test Name", 15) &  " | " & addBlankSpaces("ID", 6) & " | " & addBlankSpaces("Host", 20) &  "|"
          WScript.StdOut.WriteLine generateLine(50)

          ReDim tests(tList.Count - 1)

          For Each test In tList
            WScript.StdOut.WriteLine "| " & addBlankSpaces(test.Name, 15) &  " | " & addBlankSpaces(test.ID, 6) & " | " & addBlankSpaces(test.HostName, 20) &  "|"
            WScript.StdOut.WriteLine generateLine(50)
            Scheduler.RunOnHost(test.ID) = test.HostName
          Next

          Scheduler.RunAllLocally = False
          Scheduler.run
          WScript.StdOut.WriteBlankLines(2)
          WScript.StdOut.WriteLine "Running-Tests..."
          WScript.StdOut.WriteBlankLines(2)
          Set executionStatus = Scheduler.ExecutionStatus

          While ((tsExecutionFinished = False) And (iter < timeout))
            iter = iter + 5
            executionStatus.RefreshExecStatusInfo "all", True
            tsExecutionFinished = executionStatus.Finished
            WScript.Sleep( 5000 )
          Wend

          If iter < timeout Then

            WScript.StdOut.WriteLine generateLine(50)
            WScript.StdOut.WriteLine "| " & addBlankSpaces("Test", 22) &  " | " & addBlankSpaces("Result", 22) & "|"
            WScript.StdOut.WriteLine generateLine(50)

            For i = 1 To executionStatus.Count
              Set qTest = New QCTest
              Set testExecStatusObj = executionStatus.Item(i)

              Set currentTest = targetTestSet.TSTestFactory.Item(testExecStatusObj.TSTestId)

              qTest.Name = currentTest.Name
              qTest.Duration = currentTest.LastRun.Field("RN_DURATION")

              If Not (currentTest.LastRun.Status = "Passed") Then
                Set qFailure = New QCFailure
                qFailure.Desc = testExecStatusObj.Message
                qFailure.Name = currentTest.LastRun.Status
                Set qTest.Failure = qFailure
              Else
                Set qTest.Failure = Nothing
              End If

              Set tests(i - 1) = qTest
                WScript.StdOut.WriteLine "| " & addBlankSpaces(currentTest.Name, 22) &  " | " & addBlankSpaces(currentTest.LastRun.Status, 22) & "|"
                WScript.StdOut.WriteLine generateLine(50)
            Next

            WScript.StdOut.WriteBlankLines(2)
            WScript.StdOut.WriteLine "Scheduler finished around " & CStr(Now)
          Else
            errorMsg = "Timed out"
          End If
	    End If
      End If
    End If
  End Sub

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
    Dim logFilePath
    Dim currentDate
    Dim numError
    Dim numFailure
    Dim numTest
    Dim totalTime
    Dim body
    Dim header

    WScript.StdOut.WriteLine "Creating report..."
    currentDate = YEAR(Date()) & _
                                    "-" & prefixWithZero(Month(Date()),2) & _
                                    "-" & prefixWithZero(Day(Date()),2) & _
                                    "T" & prefixWithZero(Hour(Now()),2) & _
                                    ":" & prefixWithZero(Minute(Now()),2) & _
                                    ":" & prefixWithZero(Second(Now()),2)

    Set logShell = CreateObject("Wscript.Shell")
    logFilePath = logShell.CurrentDirectory & "\" & fileName
    Set logShell = Nothing
    Set objStream = CreateObject("ADODB.Stream" )
    objStream.Open
    objStream.Position = 0
    objStream.Charset = "UTF-8"

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
        totalTime = totalTime + test.Duration
        body = 	body & vbTab & "<testcase classname=""" & domain & "." & project & "." & folder & "." & name & """ " & _
                                                        "name=""" & test.Name & """ " & _
                                                        "time=""" & test.Duration  & ".0"" "

        If test.Failure Is Nothing Then
          body = body & "/>" & vbCrLf
        Else
          body = body & ">" & vbCrLf

          body = body & vbTab & vbTab & "<failure message=""" & stripTags(test.Failure.Desc) & """ type=""" & test.Failure.name & """>" & vbCrLf & _
                                    vbTab & vbTab & vbTab & test.Failure.Name & " : " & stripTags(test.Failure.Desc) & vbCrLf & _
                                    vbTab & vbTab & "</failure>" & vbCrLf

          body = body & vbTab & "</testcase>" & vbCrLf
          numFailure = numFailure + 1
        End If
      Next
    End If

    header = header & "<testsuite errors=""" & numError & """ " & _
                                                    "failures=""" & numFailure & """  " & _
                                                    "hostname=""" & hostName &  """  " & _
                                                    "name=""" & domain & "." & project & "." & folder & "." & name & """  " & _
                                                    "tests=""" & numTest & """ " & _
                                                    "time=""" & totalTime & ".0"" " & _
                                                    "timestamp=""" & currentDate & """>"
    header = header & body
    header = header & "</testsuite>"
    objStream.WriteText header
    objStream.SaveToFile logFilePath
    objStream.Close
    WScript.StdOut.WriteLine "Report Created"
    
  End Sub
	
End Class	

' --- Main ---

Dim args
Dim test
Dim qcTimeout
Set args = WScript.Arguments
Set test = New QCTestRunner

If args.Count > 8 Then
  qcTimeout = CInt(args.Item(8))
Else
  qcTimeout = 600
End If	


test.ConnectToProject args.Item(0), args.Item(1), args.Item(2), args.Item(3), args.Item(4)
If test.Connected Then
  test.RunTestSet args.Item(5), args.Item(6), qcTimeout
End If
If Not (test.ErrorMessage = "") Then
  WScript.StdOut.WriteLine test.ErrorMessage
  test.WriteToXML args.Item(7)
  WScript.Quit 1
End If
test.Disconnect
test.WriteToXML args.Item(7)
WScript.Quit 0

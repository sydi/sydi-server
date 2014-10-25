Option Explicit
'==========================================================
' LANG : VBScript
' NAME : sydi-wrapper.vbs
' AUTHOR : Patrick Ogenstad (patrick.ogenstad@netsafe.se)
' VERSION : 0.2.3
' DATE : 2008-12-05
' Description : Used as a wrapper for SYDI-Server
'
' Running the script:
' Edit the settings below and run with command line options
'
' For Options: cscript.exe sydi-wrapper.vbs -h
' Feedback: Please send feedback to patrick.ogenstad@netsafe.se
'
' LICENSE :
' Copyright (c) 2004-2006, Patrick Ogenstad
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'
'  * Redistributions of source code must retain the above copyright notice,
'    this list of conditions and the following disclaimer.
'  * Redistributions in binary form must reproduce the above copyright notice,
'    this list of conditions and the following disclaimer in the documentation
'    and/or other materials provided with the distribution.
'  * Neither the name SYDI nor the names of its contributors may be used
'    to endorse or promote products derived from this software without
'    specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
' IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
' ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE
' LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
' CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
' SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
' INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
' CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
' ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
' POSSIBILITY OF SUCH DAMAGE.
'==========================================================
'==========================================================
' Settings
Dim strScriptVersion
strScriptVersion = "0.2.3"
Dim WMIOPTIONS, REGISTRYOPTIONS
Dim SYDISERVER, OUTPUTDIRECTORY, LOGDIRECTORY
Dim EXPORTFORMAT, OTHEROPTIONS
Dim TIMEOUT

' Gathering Options
' WMI - Everything enabled by default
WMIOPTIONS="-wabefghipPqrsSu"
' Registry - Everything enabled by default
REGISTRYOPTIONS="-racdklp"

' Export Options
'EXPORTFORMAT="word" ' For Microsoft Word
EXPORTFORMAT="xml" ' For XML

' Location Options
' Location of SYDI-Server.vbs
SYDISERVER="C:\scripts\sydi-server.vbs"
OUTPUTDIRECTORY="C:\scripts\Output files\"
LOGDIRECTORY="C:\Scripts\Log Files\"
TIMEOUT="600" ' How many seconds to wait until a computer-scan is aborted, this hasn't been tested

' Other options, check sydi-server.vbs -h for help
' Uncoment/Change One of the below
OTHEROPTIONS="-sh" ' For HTML Stylesheet on XML output
'OTHEROPTIONS="-b10" ' Base Font size of 12
' End Of Settings


Dim bDisplayHelp, bInvalidArgument, bAlternateCredentials
Dim bGotComputerList
Dim strUserName, strPassword

' Export
Dim strExportstring, strExportExt, strWrapperLogFile
Dim objFSOW, objFSWrapperLog

' Computers
Dim objDbrComputers

' Sources
Dim strSourceType
Dim strSourceTextfile, strSourceDomain, strSourceAD


' Constants
Const adVarChar = 200
Const MaxCharacters = 255
Const ForReading = 1
Const ForWriting = 2
'==========================================================
'==========================================================
' Main Body

If LCase (Right (WScript.FullName, 11)) <> "cscript.exe" Then
    MsgBox "This script should be run from a command line (eg ""cscript.exe sydi-wrapper.vbs"")", vbCritical, "Error"
    WScript.Quit
End If

Init
GetOptions
Main

'==========================================================
'==========================================================
' Procedures

Function CreateReportFile()
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If (objFSO.FolderExists(LOGDIRECTORY)) Then
		strWrapperLogFile = "SydiWrapper.log"
		Set objFSOW = CreateObject("Scripting.FileSystemObject")
		Set objFSWrapperLog = objFSOW.OpenTextFile(LOGDIRECTORY & strWrapperLogFile, ForWriting, True)
		WriteStars
		ReportProgress "* Running SYDI-Wrapper v." & strScriptVersion
		ReportProgress "* SYDISERVER: " & SYDISERVER
		ReportProgress "* OUTPUTDIRECTORY: " & OUTPUTDIRECTORY
		ReportProgress "* LOGDIRECTORY: " & LOGDIRECTORY
		ReportProgress "* TIMEOUT: " & TIMEOUT & " seconds"
		ReportProgress "* Time of Scan: " & Now
		ReportProgress "* Current Options: " & SYDISERVER & " " & WMIOPTIONS & " " & REGISTRYOPTIONS & " " & strExportstring
		ReportProgress "*   -t%Computer% -o" & OUTPUTDIRECTORY & "%Computer%" & strExportExt & " " & OTHEROPTIONS
		If (bAlternateCredentials) Then
			ReportProgress "* Using Alternate Credentials"
		End If
	Else
		Wscript.Echo "Folder Doesn't Exist: " & LOGDIRECTORY
		CreateReportFile = False
		Exit Function
	End If
	
	CreateReportFile = True
End Function ' CreateReportFile()

Sub DisplayHelp
	WScript.Echo "SYDI-Wrapper v." & strScriptVersion
	WScript.Echo "Usage: cscript.exe sydi-wrapper.vbs [options]"
	WScript.Echo "Example: cscript.exe sydi-wrapper.vbs -tComputers.csv"
	WScript.Echo ""
	WScript.Echo "Target options"
	WScript.Echo " -d	- Get All Computers from Domain ie WinNT:// (-dYOURDOMAIN)"
	WScript.Echo " -a	- Get Computers from active directory"
 	WScript.Echo "   	  -a (Find Domain From RootDSE)"
 	WScript.Echo "   	  -aDC=yourdomain,DC=com (Specify domain to search)"
 	WScript.Echo "   	  -a""OU=Corp Members,DC=yourdomain,DC=com"" (Target specific OU)"
	WScript.Echo " -t	- Get Computers from textfile (-tc:\lists\computers.txt)"
 	WScript.Echo "   	  (Computers by line and/or comma separated)"
	WScript.Echo " -u	- Username (To run with different credentials)"
	WScript.Echo " -p	- Password (To run with different credentials, must be used with -u)"
 	WScript.Echo VbCrLf
	WScript.Echo "   	  Please also note that you have to change the settings"
	WScript.Echo "   	  inside sydi-wrapper.vbs for the script to work!"
 	WScript.Echo VbCrLf
	WScript.Echo " -h	- Display help"
 	WScript.Echo VbCrLf
End Sub ' DisplayHelp

Function GetComputersFromActiveDirectory()
	Dim objADODBConn, objComm, objDbrAD, objRootDSE

	Const ADS_SCOPE_SUBTREE = 2
	Const Disabled = 2

	If (strSourceAD = "") Then
		Set objRootDSE = GetObject("LDAP://RootDSE")
		strSourceAD = objRootDSE.Get("DefaultNamingContext")
	End If

	ReportProgress "* Targeting computers from Active Directory"
	ReportProgress "* Using list from: " & strSourceAD
	
	Set objADODBConn = CreateObject("ADODB.Connection")
	Set objComm =   CreateObject("ADODB.Command")
	objADODBConn.Provider = "ADsDSOObject"
	objADODBConn.Open "Active Directory Provider"
	
	Set objComm.ActiveConnection = objADODBConn
	objComm.CommandText = "Select Name, userAccountControl from 'LDAP://" & strSourceAD & "' " _
			& "Where objectClass='computer' Order by Name"  
	objComm.Properties("Page Size") = 1000
	objComm.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
	Set objDbrAD = objComm.Execute
	objDbrAD.MoveFirst
	
	Do Until objDbrAD.EOF
		If (objDbrAD.Fields("userAccountControl") and disabled) Then
			ReportProgress "Ignoring disabled account: " & objDbrAD.Fields("Name").Value
		Else
			objDbrComputers.AddNew
			objDbrComputers("Name") = objDbrAD.Fields("Name").Value
		End If
		objDbrAD.MoveNext
	Loop
	GetComputersFromActiveDirectory = True
End Function ' GetComputersFromActiveDirectory()

Function GetComputersFromDomain()
	Dim objDomain, objComputer
	ReportProgress "* Targeting computers from Domain"
	ReportProgress "* Using list from: " & Ucase(strSourceDomain)
	Set objDomain = GetObject("WinNT://" & strSourceDomain)
	objDomain.Filter = Array("computer")
	For Each objComputer In objDomain
		objDbrComputers.AddNew
		objDbrComputers("Name") = objComputer.Name
	Next
	GetComputersFromDomain = True
End Function ' GetComputersFromDomain()

Function GetComputersFromFile()
	Dim objFSO, objTextFile, strTextLine, arrComputers, i
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If (objFSO.FileExists(strSourceTextfile)) Then
		ReportProgress "* Targeting computers from file"
		ReportProgress "* Using list from: " & Ucase(strSourceTextfile)
	
		Set objTextFile = objFSO.OpenTextFile (strSourceTextfile, ForReading)
		Do Until objTextFile.AtEndOfStream
			strTextLine = objTextFile.Readline
			arrComputers = Split(strTextLine , ",")
			For i = 0 to Ubound(arrComputers)
				objDbrComputers.AddNew
				objDbrComputers("Name") = arrComputers(i)
			Next
		Loop
	Else
		Wscript.Echo "File Doesn't Exist: " & strSourceTextfile
		GetComputersFromFile = False
	End If
	GetComputersFromFile = True
End Function ' GetComputersFromFile()


Sub GetOptions()
	Dim objArgs, nArgs
	
	Set objArgs = WScript.Arguments
	If (objArgs.Count > 0) Then
		For nArgs = 0 To objArgs.Count - 1
			SetOptions objArgs(nArgs)
		Next
	Else
		bDisplayHelp = True
	End If

End Sub ' GetOptions

Sub Init()
	' Default settings
	Dim strExportformat
	
	bDisplayHelp = False
	bGotComputerList = False
	bAlternateCredentials = False
	strSourceType = ""
	strExportformat = Lcase(EXPORTFORMAT)
	Select Case strExportformat
		Case "word"
			strExportstring = "-ew -d"
			strExportExt = ".doc"
		Case "xml"
			strExportstring = "-ex"
			strExportExt = ".xml"
		Case Else
			Wscript.Echo "Wrong EXPORTFORMAT: " & EXPORTFORMAT & " use word or xml"
			bInvalidArgument = True
	End Select
	
	Set objDbrComputers = CreateObject("ADOR.Recordset")
	objDbrComputers.Fields.Append "Name", adVarChar, MaxCharacters
	objDbrComputers.Fields.Append "ErrorNumber", adVarChar, MaxCharacters
	objDbrComputers.Fields.Append "ErrorDescription", adVarChar, MaxCharacters
	objDbrComputers.Fields.Append "PassedTest", adVarChar, MaxCharacters
	objDbrComputers.Fields.Append "DomainRole", adVarChar, MaxCharacters
	objDbrComputers.Fields.Append "WMIName", adVarChar, MaxCharacters
	objDbrComputers.Open
	
End Sub ' Init

Sub Main()
	If (bInvalidArgument) Then
		WScript.Echo "Invalid Arguments" & VbCrLf
		bDisplayHelp = True
	End If
	
	If (bDisplayHelp) Then
		DisplayHelp
	Else
		If Not (CreateReportFile) Then
			Exit Sub
		End If
		Select Case strSourceType
			Case "text"
				bGotComputerList = GetComputersFromFile()
			Case "domain"
				bGotComputerList = GetComputersFromDomain()
			Case "ad"
				bGotComputerList = GetComputersFromActiveDirectory()
			Case Else
				Wscript.Echo "Error"
		End Select
		If (bGotComputerList) Then
			ScanTargets
		End If
		objFSWrapperLog.Close
	End If
End Sub ' Main

Sub ReportProgress(strMessage)
	WScript.Echo strMessage
	objFSWrapperlog.WriteLine strMessage
End Sub ' ReportProgress



Sub RunSYDI(strComputer)
	Dim strSYDIString, objShell, objExec
	Dim nRuns, nAcceptedRuns
	
	nAcceptedRuns = TIMEOUT * 1000
	nRuns = 0
	strSYDIString = "cscript.exe //nologo """ & SYDISERVER & """ " & WMIOPTIONS & " " & REGISTRYOPTIONS & " " & strExportstring & _
		" -t" & strComputer & " -o""" & OUTPUTDIRECTORY & strComputer & strExportExt & """ " & OTHEROPTIONS 
	If (bAlternateCredentials) Then
		strSYDIString = strSYDIString & " -u" & strUserName & " -p" & strPassword
	End If
		
	Set objShell = CreateObject("WScript.Shell")
	Set objExec = objShell.Exec(strSYDIstring)
	Do While objExec.Status = 0
		If (nRuns = nAcceptedRuns) Then
			WScript.StdOut.Write " Timed out" & vbCrLf
			'objExec.Terminate()
			Exit Do
		End If
		WScript.Sleep 100
		WScript.StdOut.Write "."
		
		nRuns = nRuns + 100
	Loop
	
	If (objExec.Status = 1) Then
		WScript.StdOut.Write " DONE" & vbCrLf
	End If

	WriteComputerLog strComputer, objExec.StdOut.Readall
	
End Sub ' RunSYDI()

Sub ScanTargets()
	Dim nComputersDone, nErrorNumber
	nComputersDone = 1
	ReportProgress "* Hosts Selected: " & objDbrComputers.Recordcount
	WriteStars
	ReportProgress vbCrLf
	WriteStars
	ReportProgress "* Testing connection and passing to SYDI"
	ReportProgress "*"
	If Not (objDbrComputers.Bof) Then
		objDbrComputers.MoveFirst
	End If
	Do Until objDbrComputers.Eof
		
		ReportProgress "* Testing connection to " & objDbrComputers.Fields.Item("Name") & " (" & nComputersDone & "/" & objDbrComputers.Recordcount & ")"
		objDbrComputers("PassedTest") = TestConnection(objDbrComputers.Fields.Item("Name"))
		If (objDbrComputers.Fields.Item("PassedTest")) Then
			ReportProgress "* Passing " & objDbrComputers.Fields.Item("Name") & " to SYDI-Server"
			RunSYDI(objDbrComputers.Fields.Item("WMIName"))
		End If
		
		nComputersDone = nComputersDone + 1
		objDbrComputers.MoveNext
	Loop
	WriteStars
	ReportProgress vbCrLf
	WriteStars
	ReportProgress "* Status"
	
	If Not (objDbrComputers.Bof) Then
		objDbrComputers.MoveFirst
		objDbrComputers.Sort = "ErrorNumber"
	End If
	
	Do Until objDbrComputers.Eof
		If (nErrorNumber <> objDbrComputers.Fields.Item("ErrorNumber")) Then
			ReportProgress "*"
			ReportProgress "* ErrorNumber: " & objDbrComputers.Fields.Item("ErrorNumber") 
			ReportProgress "* Description: " & objDbrComputers.Fields.Item("ErrorDescription")
		End If
		ReportProgress "*    " & objDbrComputers.Fields.Item("Name")
		nErrorNumber = objDbrComputers.Fields.Item("ErrorNumber")
		objDbrComputers.MoveNext
	Loop
	ReportProgress "*"
	WriteStars
	
End Sub ' ScanTargets

Sub SetOptions(strOption)
	Dim strFlag, strParameter
	Dim nArguments
	nArguments = Len(strOption)
	If (nArguments < 2) Then
		bInvalidArgument = True
	Else
		strFlag = Left(strOption,2)
		Select Case strFlag
			Case "-a"
				strSourceAD = ""
				strSourceType = "ad"
				If (nArguments > 2) Then
					strSourceAD = Right(strOption,(nArguments - 2))
				End If
			Case "-d"
				strSourceDomain = ""
				strSourceType = "domain"
				If (nArguments > 2) Then
					strSourceDomain = Right(strOption,(nArguments - 2))
				End If
				If (strSourceDomain = "") Then
					bInvalidArgument = True
				End If
			Case "-p"
				strPassword = ""
				If (nArguments > 2) Then
					strPassword = Right(strOption,(nArguments - 2))
				End If
			Case "-u"
				strUserName = ""
				bAlternateCredentials = True
				If (nArguments > 2) Then
					strUserName = Right(strOption,(nArguments - 2))
				End If
			Case "-t"
				strSourceTextfile = ""
				strSourceType = "text"
				If (nArguments > 2) Then
					strSourceTextfile = Right(strOption,(nArguments - 2))
				End If
				If (strSourceTextfile = "") Then
					bInvalidArgument = True
				End If
			Case "-h"
				bDisplayHelp = True
			Case Else
				bInvalidArgument = True
		End Select
	
	End If

End Sub ' SetOptions

Function TestConnection(strComputer)
	On Error Resume Next
	Dim objSWbemLocator, objWMIService, colItems, objItem
	If (bAlternateCredentials) Then
		Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
		Set objWMIService = objSWbemLocator.ConnectServer(strComputer,"\root\cimv2",strUserName,strPassword)
	Else
		Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	End If
	If (Err <> 0) Then
		objDbrComputers("ErrorNumber") = Err.Number
		objDbrComputers("ErrorDescription") = Err.Description
	    ReportProgress "* Failed to connect to: " & strComputer
	    Err.Clear
	    TestConnection = False
	    Exit Function
	End If
	Set colItems = objWMIService.ExecQuery("Select DomainRole, Name From Win32_ComputerSystem",,48)
	For Each objItem in colItems
		objDbrComputers("ErrorNumber") = 0
		objDbrComputers("ErrorDescription") = "Passed to SYDI-Server"
		objDbrComputers("DomainRole") = objItem.DomainRole
		objDbrComputers("WMIName") = objItem.Name
	Next
	TestConnection = True
End Function ' TestConnection()

Sub WriteComputerLog(strComputer,strComputerLog)
	Dim objFSO, objFSLogFile
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFSLogFile = objFSO.OpenTextFile(LOGDIRECTORY & strComputer & ".Log", ForWriting, True)
	objFSLogFile.Write strComputerLog
	objFSLogFile.Close
End Sub ' WriteComputerLog

Sub WriteStars
	WScript.Echo "************************************************************"
	objFSWrapperlog.WriteLine "************************************************************"
End Sub ' WriteStars


'==========================================================
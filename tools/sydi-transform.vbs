Option Explicit
'==========================================================
' LANG : VBScript
' NAME : sydi-transform.vbs
' AUTHOR : Patrick Ogenstad (patrick.ogenstad@netsafe.se)
' VERSION : 1.0.1
' DATE : 2004-11-17
' Description : Transforms a SYDI-Server XML file into html output
'
' UPDATES : http://sydiproject.com/ (Part of SYDI-Server)
'
' Running the script:
' cscript.exe sydi-transform.vbs -xServer.XML -sServerHtml.xsl -oServer.html
' For Options: cscript.exe sydi-transform.vbs -h
' Feedback: Please send feedback to patrick.ogenstad@netsafe.se
'
' LICENSE :
' Copyright (c) 2004-2007, Patrick Ogenstad
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
Dim strXMLFile, strXSLFile, strOutputFile
Dim bInvalidArgument, bDisplayHelp

' Script version
Dim strScriptVersion
strScriptVersion = "1.0.1"

'==========================================================
'==========================================================
' Main Body
If LCase (Right (WScript.FullName, 11)) <> "cscript.exe" Then
    MsgBox "This script should be run from a command line (eg ""cscript.exe sydi-transform.vbs"")", vbCritical, "Error"
    WScript.Quit
End If

' Get Options from user
GetOptions

If Not (bDisplayHelp) Then
	If (strXMLFile = "" Or strXSLFile = "" Or strOutputFile = "") Then
		bInvalidArgument = True
	End If 
End If

If (bInvalidArgument) Then
	WScript.Echo "Invalid Arguments" & VbCrLf
	bDisplayHelp = True
End If 

If (bDisplayHelp) Then
	DisplayHelp 
Else
	TransformFiles
End If

'==========================================================
'==========================================================
' Procedures

Sub DisplayHelp
	WScript.Echo "SYDI-Transform v." & strScriptVersion
	WScript.Echo "Usage: cscript.exe sydi-transform.vbs -x[file] -s[file] -o[file]"
	WScript.Echo VbCrLf
	WScript.Echo "Example:"
	WScript.Echo "cscript.exe sydi-transform.vbs -xServer.xml -sServerhtml.xsl -oServer.html"
	WScript.Echo VbCrLf
	WScript.Echo "Options"
	WScript.Echo " -x	- XML Input File"
 	WScript.Echo " -s	- XSL Stylesheet"
 	WScript.Echo " -o	- Output File"
 	WScript.Echo VbCrLf
 	WScript.Echo " -h	- Display help"
 	WScript.Echo VbCrLf
End Sub ' DisplayHelp

Sub GetOptions()
	Dim objArgs, nArgs
	' Default settings
	bInvalidArgument = False

	Set objArgs = WScript.Arguments
	If (objArgs.Count > 0) Then
		For nArgs = 0 To objArgs.Count - 1
			SetOptions objArgs(nArgs)
		Next
	Else
		bDisplayHelp = True
	End If
End Sub ' GetOptions

Sub SetOptions(strOption)
	Dim strFlag, strParameter
	Dim nArguments
	nArguments = Len(strOption)
	If (nArguments < 2) Then
		bInvalidArgument = True
	Else
		strFlag = Left(strOption,2)
		Select Case strFlag
			Case "-x"
				strXMLFile = ""
				If (nArguments > 2) Then
					strXMLFile = Right(strOption,(nArguments - 2))
				End If
			Case "-s"
				strXSLFile = ""
				If (nArguments > 2) Then
					strXSLFile = Right(strOption,(nArguments - 2))
				End If
			Case "-o"
				strOutputFile = ""
				If (nArguments > 2) Then
					strOutputFile = Right(strOption,(nArguments - 2))
				End If
			Case "-h"
				bDisplayHelp = True
			Case Else
				bInvalidArgument = True
		End Select
	End If
End Sub ' SetOptions


Sub TransformFiles()
	Dim objFS, objFSOutput
	Dim objXML, objXSL

	Set objXML = CreateObject("Microsoft.XMLDOM")
	Set objXSL = CreateObject("Microsoft.XMLDOM")
	Set objFS = CreateObject("Scripting.FileSystemObject")
	Set objFSOutput = objFS.CreateTextFile(strOutputFile, True)

	objXML.ValidateOnParse = True
	objXSL.ValidateOnParse = True	
	objXML.Load(strXMLFile)
	objXSL.Load(strXSLFile )

	objFSOutput.Write objXML.transformNode(objXSL)
	objFSOutput.Close
End Sub ' TransformFiles

'==========================================================
Option Explicit
'==========================================================
' LANG : VBScript
' NAME : sydi-audit-localgroups.vbs
' AUTHOR : Patrick Ogenstad (patrick.ogenstad@netsafe.se)
' VERSION : 0.2
' DATE : 2008-01-29
' Description : Creates an Excel overview of SYDI-Server XML files
'
' UPDATES : http://sydiproject.com/ (Part of SYDI-Server)
'
' Running the script:
' cscript.exe sydi-audit-localgroups.vbs -x[directory to xml files]
' For Options: cscript.exe sydi-audit-localgroups.vbs -h
' Feedback: Please send feedback to patrick.ogenstad@netsafe.se
'
' Notes:
' This has has only been tested with Office 2003 and Office 2007
' LICENSE :
' Copyright (c) 2004-2008, Patrick Ogenstad
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

' Script version
Dim strScriptVersion
strScriptVersion = "0.1"

' Variables
Dim strXMLDirectory
Dim bInvalidArgument, bDisplayHelp
Dim intXMLFileCount, intXMLTotalFiles

' Database Records
Dim objDbrXMLFiles, objDbrComputers, objDbrLocalGroups, objDbrComputerServices, objDbrOSDistribution, objDbrRegistryPrograms, objDbrWMIPrograms
Dim objDbrProcesses, objDbrServiceList

' Settings
Dim bServiceComparison


' Constants
Const adVarChar = 200
Const MaxCharacters = 255

'==========================================================
'==========================================================
' Main Body
If LCase (Right (WScript.FullName, 11)) <> "cscript.exe" Then
    MsgBox "This script should be run from a command line (eg ""cscript.exe sydi-overview.vbs"")", vbCritical, "Error"
    WScript.Quit
End If

' Get Options from user
GetOptions


If (bInvalidArgument) Then
	WScript.Echo "Invalid Arguments" & VbCrLf
	bDisplayHelp = True
End If 

If (bDisplayHelp) Then
	DisplayHelp 
Else
	Init
	If (GetFileList(strXMLDirectory)) Then
		If Not (objDbrXMLFiles.Bof) Then
			objDbrXMLFiles.MoveFirst
			intXMLTotalFiles = objDbrXMLFiles.RecordCount
			intXMLFileCount = 1
		End If
		Do Until objDbrXMLFiles.EOF
			ParseXMLFile(objDbrXMLFiles.Fields.Item("XMLFile"))
			objDbrXMLFiles.MoveNext
		Loop
		PopulateExcelfile
	Else
		Wscript.Echo "Error"
	End If
End If

'==========================================================
'==========================================================
' Procedures

Sub DisplayHelp
	WScript.Echo "SYDI-Overview v." & strScriptVersion
	WScript.Echo "Usage: cscript.exe sydi-overview.vbs -x[xml directory]"
	WScript.Echo VbCrLf
	WScript.Echo "Example:"
	WScript.Echo "cscript.exe sydi-overview.vbs -x""D:\sydi output"""
	WScript.Echo VbCrLf
	WScript.Echo "Options"
	WScript.Echo " -x	- XML Input Directory"

 	WScript.Echo VbCrLf
 	WScript.Echo " -h	- Display help"
 	WScript.Echo VbCrLf
End Sub ' DisplayHelp

Function GetFileList(strFolder)
	Dim objFSO, objFolder, colFolder, objItem, strExtention
	Const xmlExtention = ".xml"
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFolder = objFSO.GetFolder(strFolder)
	Set colFolder = objFolder.Files
	
	For Each objItem in colFolder
		strExtention = Right(objItem.Name,4)
		If (strExtention = xmlExtention) Then
			objDbrXMLFiles.AddNew
			objDbrXMLFiles("XMLFile") = objItem.Name
			objDbrXMLFiles.Update
		End If
	Next
	GetFileList = True
End Function 'GetFileList

Sub GetOptions()
	Dim objArgs, nArgs
	' Default settings
	bInvalidArgument = False
	bServiceComparison = True
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
	Set objDbrXMLFiles = CreateObject("ADOR.Recordset")
	objDbrXMLFiles.Fields.Append "XMLFile", adVarChar, MaxCharacters
	objDbrXMLFiles.Open
	
	Set objDbrComputers = CreateObject("ADOR.Recordset")
	objDbrComputers.Fields.Append "Name", adVarChar, MaxCharacters
	objDbrComputers.Fields.Append "LocalGroup", adVarChar, MaxCharacters
	objDbrComputers.Fields.Append "Member", adVarChar, MaxCharacters
	objDbrComputers.Open

	Set objDbrLocalGroups = CreateObject("ADOR.Recordset")
	objDbrLocalGroups.Fields.Append "GroupName", adVarChar, MaxCharacters
	objDbrLocalGroups.Fields.Append "Count", adVarChar, MaxCharacters
	objDbrLocalGroups.Open
	

End Sub 'Init

Sub LocalGroups(strGroupName)	
	objDbrLocalGroups.Filter = " GroupName='" & strGroupName & "'"
	If (objDbrLocalGroups.Bof) Then
		objDbrLocalGroups.AddNew
		objDbrLocalGroups("GroupName") = strGroupName
		objDbrLocalGroups("Count") = 1
		objDbrLocalGroups.Update
	Else
		objDbrLocalGroups("Count") = objDbrLocalGroups("Count") + 1
		objDbrLocalGroups.Update
	End If
End Sub ' LocalGroups




Sub ParseXMLfile(strXMLFile)
	Dim objXMLFile, colNodes
	Dim objNode, objChild
	Dim objGroup, objMember
	Dim strCurrentSystem

	Set objXMLFile = CreateObject("Microsoft.XMLDOM")
	objXMLFile.async = False
	objXMLFile.load(strXMLDirectory & "\" & strXMLFile)
	Wscript.Echo "Parsing: " & strXMLDirectory & "\" & strXMLFile & " (" & intXMLFileCount & "/" & intXMLTotalFiles & ")"
	intXMLFileCount = intXMLFileCount + 1
	Set colNodes = objXMLFile.selectNodes("//computer")

	For Each objNode in colNodes
		For Each objChild in objNode.childNodes
			If (objChild.nodeName = "system") Then
				strCurrentSystem = Scrub(objChild.getAttribute("name"))
			End If

			If (objChild.nodeName = "localgroups") Then
				For Each objGroup in objChild.childNodes
					If (objGroup.nodeName = "group") Then
						For Each objMember in objGroup.childNodes
							objDbrComputers.AddNew
							objDbrComputers("Name") = strCurrentSystem
							objDbrComputers("Localgroup") = objGroup.getAttribute("name")
							LocalGroups(objGroup.getAttribute("name"))
							objDbrComputers("Member") = objMember.getAttribute("name")
							objDbrComputers.Update
						Next
					End If
				Next
			End If
		Next
	Next 

End Sub ' ParseXMLfile


Sub PopulateExcelfile()
	Dim objExcel, intLine, intXLine, intYLine
	Dim objWorksheet, objRange, objRange2, objRange3, objRange4
	Dim strCurrentComputer
	Dim intSheet, intPreviousSheet, intRow
	Dim intGroup, xOffSet, yOffSet
	Const xlPie = 5
	Const xlColumns = 2
	Const xlLocationAsNewSheet = 1
	Const xlNone = -4142
	Const xlThin = 2
	Const xlAscending = 1
	Const xlDescending = 2
	Const xlYes = 1
	intSheet = 2
	
	
	Wscript.Echo VbCRLF & "Opening Excel File"
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = True
	objExcel.Workbooks.Add
	
	objExcel.Sheets(1).Name = "Overview"



	objDbrLocalGroups.Filter = " Count > 0 "
	objDbrLocalGroups.Sort = "GroupName"
	If Not (objDbrLocalGroups.Bof) Then
		objDbrLocalGroups.MoveFirst
	End If
	Do Until objDbrLocalGroups.EOF
		intGroup = intSheet - 1
		If (intSheet < 4) Then
			
			objExcel.Sheets(intSheet).Name = "Group" & intGroup
			objExcel.Sheets(objExcel.Sheets(intSheet).Name).Select
		Else
			intPreviousSheet = intSheet - 1
			objExcel.Sheets.Add ,objExcel.Sheets(intPreviousSheet) ' Add a new sheet after the last one
			objExcel.Sheets(intSheet).Name = "Group" & intGroup
	
	
		End If
		objExcel.Cells(1, 1).Value = "Overview"
		Set objRange = objExcel.Range("A1")
		objExcel.ActiveSheet.Hyperlinks.Add objRange, "", "OVERVIEW!A1", "OVERVIEW"
		objExcel.Cells(2, 2).Value = objDbrLocalGroups("GroupName")
		objExcel.Range("B2").Select
		objExcel.Selection.Font.Bold = True
		objExcel.Selection.Font.Size = 24

		objExcel.Cells(4, 1).Value = "Computer"
		objExcel.Range("A4").Select
		objExcel.Selection.Font.Bold = True

		objExcel.Sheets("Overview").Select
		intRow = intSheet + 2
		objExcel.Cells(intRow, 1).Value = objDbrLocalGroups("GroupName")
		objExcel.Cells(intRow, 2).Value = "Group" & intGroup

		
		objExcel.Sheets("Overview").Select
		
		Set objRange = objExcel.Range("B" & (intSheet + 2))
		objExcel.ActiveSheet.Hyperlinks.Add objRange, "", "Group" & intGroup & "!A1", "Group" & intGroup
		
		
		objDbrLocalGroups.MoveNext
		intSheet = intSheet + 1
	Loop	
	
	objExcel.Sheets("Overview").Select
	objExcel.Cells(1, 1).Value = "Local Group Audit Overview"
	objExcel.Cells(3, 1).Value = "Group Name"
	objExcel.Cells(3, 2).Value = "Sheet"
	objExcel.Range("A3:B3").Select
    objExcel.Selection.Font.Bold = True
	objExcel.ActiveWindow.SplitRow = 3
	objExcel.ActiveWindow.FreezePanes = True
		
	For intGroup = 1 To objExcel.Sheets.Count - 1
		xOffSet = 2
		yOffSet = 5
		objExcel.Sheets("Group" & intGroup).Select
		wscript.echo " " & objExcel.Cells(2, 2).Value
		strCurrentComputer = ""
		objDbrComputers.Filter = " LocalGroup='" & objExcel.Cells(2, 2).Value & "'"
		objDbrComputers.Sort = "Name, Member"
		If Not (objDbrComputers.Bof) Then
			objDbrComputers.MoveFirst
		End If
		Do Until objDbrComputers.EOF
			If (strCurrentComputer = objDbrComputers("Name")) Then
				xOffSet = xOffSet + 1
				objExcel.Cells(yOffSet, xOffSet).Value = objDbrComputers("Member")
			Else
				strCurrentComputer = objDbrComputers("Name")
				xOffSet = 2
				yOffSet = yOffSet + 1
				objExcel.Cells(yOffSet, 1).Value = objDbrComputers("Name")
				objExcel.Cells(yOffSet, xOffSet).Value = objDbrComputers("Member")
			End If

			objDbrComputers.MoveNext
			intLine = intLine + 1
		Loop
		objExcel.Cells.EntireColumn.AutoFit	
		objExcel.ActiveWindow.SplitRow = 5
		objExcel.ActiveWindow.SplitColumn = 1
		
		objExcel.ActiveWindow.FreezePanes = True
	Next

	
	'objExcel.Range("A1:M1").Select
    'objExcel.Selection.Font.Bold = True
	'objExcel.Cells.Select
	
	objExcel.Sheets("Overview").Select
	objExcel.Cells.EntireColumn.AutoFit	
	objExcel.Range("A1").Select
End Sub 'PopulateExcelfile




Function Scrub(strInput)
	If (IsNull(strInput)) Then
		strInput = ""
	End If
	Scrub = strInput
End Function ' Scrub



Sub SetOptions(strOption)
	Dim strFlag, strParameter
	Dim nArguments, i
	nArguments = Len(strOption)
	If (nArguments < 2) Then
		bInvalidArgument = True
	Else
		strFlag = Left(strOption,2)
		Select Case strFlag
			Case "-x"
				strXMLDirectory = ""
				If (nArguments > 2) Then
					strXMLDirectory = Right(strOption,(nArguments - 2))
				End If
			Case "-h"
				bDisplayHelp = True
			Case Else
				bInvalidArgument = True
		End Select
	End If
End Sub ' SetOptions


'==========================================================
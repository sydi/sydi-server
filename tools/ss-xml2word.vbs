Option Explicit
'==========================================================
' LANG : VBScript
' NAME : ss-xml2word.vbs (SYDI-Server XML to Word)
' AUTHOR : Patrick Ogenstad (patrick.ogenstad@netsafe.se)
' VERSION : 0.2
' DATE : 2008-01-29
' Description : Converts a SYDI-Server XML to a Microsoft Word Document
'
' UPDATES : http://sydiproject.com/ (Part of SYDI-Server)
'
' For Options: cscript.exe ss-xml2word.vbs -h
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
strScriptVersion = "0.2"

' Variables for Script Arguments
Dim strSYDIServerSource, strWrittenServerSource
Dim strLanguageFile, bSaveFile, strSaveFile

' Variables for Script functions
Dim bInvalidArgument, bDisplayHelp
Dim bUseDOTFile

' Variables for SYDI-Server Source object

Dim strSystem_Name
Dim objDbrVariables, objDbrIISVariables
Dim objDbrRecordsets, objDbrIISRecordsets
Dim strTempConcatString

' Variables for Word
Dim oWord, oListTemplate, bShowWord
Dim strDOTFile, bUseSpecificTable, strWordTable
Dim strFontBodyText, strFontHeading1, strFontHeading2, strFontHeading3, strFontHeading4, strFontTitle, strFontTOC1, strFontTOC2, strFontTOC3
Dim strFontHeader, strFontFooter
strFontBodyText = "Arial"
strFontHeading1 = "Trebuchet MS"
strFontHeading2 = "Trebuchet MS"
strFontHeading3 = "Trebuchet MS"
strFontHeading4 = "Trebuchet MS"
strFontTitle = "Trebuchet MS"
strFontTOC1 = "Trebuchet MS"
strFontTOC2 = "Trebuchet MS"
strFontTOC3 = "Trebuchet MS"
strFontHeader = "Arial"
strFontFooter = "Arial"
Dim nBaseFontSize

' A few counters are always handy
Dim i, j


' Variables for Written Documentation File
Dim objDbrSections '  (SectionName, Subsection=True/False, Parentsection
Dim objDbrWrittenText ' (TextField, SectionName, Parentsection, Position= Pre/Post

' Constants for Database records
Const adVarChar = 200
Const adLongVarChar = 201
Const MaxCharacters = 255
Const LargeMaxCharacters = 4000

' VBA Constants for Word
' WdListNumberStyle
Const wdListNumberStyleArabic = 0
Const wdListNumberStyleUppercaseRoman = 1
Const wdListNumberStyleLowercaseRoman = 2
Const wdListNumberStyleUppercaseLetter = 3
Const wdListNumberStyleLowercaseLetter = 4
Const wdListNumberStyleOrdinal = 5
Const wdListNumberStyleCardinalText = 6
Const wdListNumberStyleOrdinalText = 7
Const wdListNumberStyleArabicLZ = 22
Const wdListNumberStyleBullet = 23
Const wdListNumberStyleLegal = 253
Const wdListNumberStyleLegalLZ = 254
Const wdListNumberStyleNone = 255
' WdListGalleryType
Const wdBulletGallery = 1
Const wdNumberGallery = 2
Const wdOutlineNumberGallery = 3
' WdBreakType
Const wdPageBreak = 7
' WdBuiltInProperty
Const wdPropertyAuthor = 3
Const wdPropertyComments = 5
' WdBuiltInStyle
Const wdStyleBodyText = -67
Const wdStyleFooter = -33
Const wdStyleHeader = -32
Const wdStyleHeading1 = -2
Const wdStyleHeading2 = -3
Const wdStyleHeading3 = -4
Const wdStyleHeading4 = -5
Const wdStyleTitle = -63
Const wdStyleTOC1 = -20
Const wdStyleTOC2 = -21
Const wdStyleTOC3 = -22
' WdFieldType
'Const wdFieldEmpty = -1
Const wdFieldNumPages = 26
Const wdFieldPage = 33
' WdParagraphAlignment
Const wdAlignParagraphRight = 2
' WdSeekView
Const wdSeekMainDocument = 0
Const wdSeekCurrentPageHeader = 9
Const wdSeekCurrentPageFooter = 10
' Page Viewing
Const wdPaneNone = 0
Const wdPrintView = 3


'==========================================================
'==========================================================
' Main Body
If LCase (Right (WScript.FullName, 11)) <> "cscript.exe" Then
    MsgBox "This script should be run from a command line (eg ""cscript.exe ss-xml2word.vbs"")", vbCritical, "Error"
    WScript.Quit
End If

' Get Options from user
GetOptions
If (strLanguageFile = "") Then
	WScript.Echo "You have to specify a language file"
	bInvalidArgument = True
End If
If (strSYDIServerSource = "") Then
	WScript.Echo "You have to specify a SYDI-Server source file"
	bInvalidArgument = True
End If


If (bInvalidArgument) Then
	WScript.Echo "Invalid Arguments" & VbCrLf
	bDisplayHelp = True
End If 

If (bDisplayHelp) Then
	DisplayHelp 
Else
	Init
	GetSYDIServerSourceData(strSYDIServerSource)
	GetWrittenDocumentationData(strWrittenServerSource)
	If (CreateDocument) Then ' Verify that Word works
		ParseLanguageFile(strLanguageFile)
		UpdateDocumentInformation
	End If
End If

'==========================================================
'==========================================================
' Procedures


Function CheckCriteria(strType,strValue)
	Dim bReturnValue
	bReturnValue = True ' Default to True
	Select Case strType
		Case "record"
			objDbrRecordsets.Filter = " name='" & strValue & "'"
			If (objDbrRecordsets.Recordcount > 0) Then
				bReturnValue = True
			Else
				bReturnValue = False
			End If
		Case "string"
			objDbrVariables.Filter = " name='" & strValue & "'"
			If (objDbrVariables.Recordcount > 0) Then
				bReturnValue = True
			Else
				bReturnValue = False
			End If
	End Select
	CheckCriteria = bReturnValue	
End Function 'CheckCriteria

Sub CheckFiles
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	If Not objFSO.FileExists(strLanguageFile) Then
		Wscript.Echo "Unable to find the file: " & strLanguageFile
		Wscript.Echo "Did you specify the correct path?"
		bInvalidArgument = True
	End If
	If Not objFSO.FileExists(strSYDIServerSource) Then
		Wscript.Echo "Unable to find the file: " & strSYDIServerSource
		Wscript.Echo "Did you specify the correct path?"
		bInvalidArgument = True
	End If
	If (Len(strWrittenServerSource) > 0 AND Not objFSO.FileExists(strSYDIServerSource)) Then
		Wscript.Echo "Unable to find the file: " & strWrittenServerSource
		Wscript.Echo "Did you specify the correct path?"
		bInvalidArgument = True
	End If
End Sub ' Check Files


Function CreateDocument
	'On Error Resume Next
	ReportProgress VbCrLf & "Creating new document"
	Set oWord = CreateObject("Word.Application")
	If (Err <> 0) Then
	    ReportProgress Err.Number & " -- " &  Err.Description 
	    ReportProgress " Could not open Microsoft Word, verify that it is correctly installed on your computer"
	    Err.Clear
		CreateDocument = False
	    Exit Function
	End If

	'oWord.Activate
	
	If (bUseDOTFile) Then
		oWord.Documents.Add strDOTFile
		If (Err <> 0) Then
		    ReportProgress Err.Number & " -- " &  Err.Description
		    ReportProgress " Unable to open the template file " & strDOTFile
		    ReportProgress " Did you use the correct path?"
		    Err.Clear
			CreateDocument = False
		    Exit Function
		End If
	Else
		oWord.Documents.Add
	End If
	oWord.Application.Visible = bShowWord
	ReportProgress " Opening Empty document"
	Set oListTemplate = oWord.ListGalleries(wdOutlineNumberGallery).ListTemplates(1)
	oWord.ListGalleries(wdOutlineNumberGallery).ListTemplates(1).listlevels(1).Numberformat = "%1."
	oWord.ListGalleries(wdOutlineNumberGallery).ListTemplates(1).listlevels(1).NumberStyle = wdListNumberStyleArabic
	oWord.ListGalleries(wdOutlineNumberGallery).ListTemplates(1).listlevels(2).Numberformat = "%1.%2."
	oWord.ListGalleries(wdOutlineNumberGallery).ListTemplates(1).listlevels(2).NumberStyle = wdListNumberStyleArabic
	oWord.ListGalleries(wdOutlineNumberGallery).ListTemplates(1).listlevels(3).Numberformat = "%1.%2.%3."
	oWord.ListGalleries(wdOutlineNumberGallery).ListTemplates(1).listlevels(3).NumberStyle = wdListNumberStyleArabic
	oWord.ListGalleries(wdOutlineNumberGallery).ListTemplates(1).listlevels(4).Numberformat = "%1.%2.%3.%4."
	oWord.ListGalleries(wdOutlineNumberGallery).ListTemplates(1).listlevels(4).NumberStyle = wdListNumberStyleArabic

	If Not (bUseDOTFile) Then
		oWord.ActiveDocument.Styles(wdStyleTOC1).Font.Bold = True
		oWord.ActiveDocument.Styles(wdStyleBodyText).Font.Name = strFontBodyText
		oWord.ActiveDocument.Styles(wdStyleBodyText).Font.Size = nBaseFontSize
		oWord.ActiveDocument.Styles(wdStyleHeading1).Font.Name = strFontHeading1
		oWord.ActiveDocument.Styles(wdStyleHeading1).Font.Size = (nBaseFontSize + 4)
		oWord.ActiveDocument.Styles(wdStyleHeading2).Font.Name = strFontHeading2
		oWord.ActiveDocument.Styles(wdStyleHeading2).Font.Size = (nBaseFontSize + 2)
		oWord.ActiveDocument.Styles(wdStyleHeading3).Font.Name = strFontHeading3
		oWord.ActiveDocument.Styles(wdStyleHeading3).Font.Size = (nBaseFontSize + 1)
		oWord.ActiveDocument.Styles(wdStyleHeading4).Font.Name = strFontHeading4
		oWord.ActiveDocument.Styles(wdStyleHeading4).Font.Size = nBaseFontSize
		oWord.ActiveDocument.Styles(wdStyleTitle).Font.Name = strFontTitle
		oWord.ActiveDocument.Styles(wdStyleTitle).Font.Size = (nBaseFontSize + 4)
		oWord.ActiveDocument.Styles(wdStyleTOC1).Font.Name = strFontTOC1
		oWord.ActiveDocument.Styles(wdStyleTOC1).Font.Size = nBaseFontSize
		oWord.ActiveDocument.Styles(wdStyleTOC2).Font.Name = strFontTOC2
		oWord.ActiveDocument.Styles(wdStyleTOC2).Font.Size = nBaseFontSize
		oWord.ActiveDocument.Styles(wdStyleTOC3).Font.Name = strFontTOC3
		oWord.ActiveDocument.Styles(wdStyleTOC3).Font.Size = nBaseFontSize
		oWord.ActiveDocument.Styles(wdStyleHeader).Font.Name = strFontHeader
		oWord.ActiveDocument.Styles(wdStyleHeader).Font.Size = (nBaseFontSize - 1)
		oWord.ActiveDocument.Styles(wdStyleFooter).Font.Name = strFontFooter
		oWord.ActiveDocument.Styles(wdStyleFooter).Font.Size = (nBaseFontSize - 1)
		ReportProgress " Setting styles"	
	End If
	oWord.Selection.Style = wdStyleBodyText
	CreateDocument = True
End Function ' CreateDocument

Sub DisplayHelp
	WScript.Echo "Ss-XML2Word v." & strScriptVersion
	WScript.Echo "Usage: ss-xml2word.vbs -x[server1.xml] -l[langfile.xml] -s[server1_data.xml]"
	WScript.Echo VbCrLf
	WScript.Echo "Example:"
	WScript.Echo "cscript.exe ss-xml2word.vbs -xD:\sydi\wwwsrv1.xml output -lD:\sydi\lang_english.xml"
	WScript.Echo VbCrLf
	WScript.Echo "Input Options"
	WScript.Echo " -x	- SYDI-Server XML source file (Required)"
	WScript.Echo " -l	- Language file (Required)"
	WScript.Echo " -s	- Written Documentation file"
	WScript.Echo "Output Options"
 	WScript.Echo " -o	- Save to file (-oc:\corpfiles\server1.doc, use in combination with -d"
 	WScript.Echo "   	  if you don't want to display word at all, use a Path or the file will"
 	WScript.Echo "  	  be placed in your default location usually 'My documents')"
 	WScript.Echo "  	  -oC:\corpfiles\server1.xml"
 	WScript.Echo "  	  WARNING USING -o WILL OVERWRITE TARGET FILE WITHOUT ASKING"
 	WScript.Echo "Word Options"
 	WScript.Echo " -b	- Use specific Word Table (-b""Table Contemporary"""
 	WScript.Echo "   	  or -b""Table List 4"")"
 	WScript.Echo " -f	- Base font size (Default: -f12)"
 	WScript.Echo " -d	- Don't display Word while writing (runs faster)"
 	WScript.Echo " -T	- Use .dot file as template (-Tc:\corptemplates\server.dot, ignores -f)"
 	WScript.Echo VbCrLf
 	WScript.Echo " -h	- Display help"
 	WScript.Echo VbCrLf
End Sub ' DisplayHelp

Sub GetOptions()
	Dim objArgs, nArgs
	' Default settings
	bInvalidArgument = False
	bUseDOTFile = False
	nBaseFontSize = 12
	bShowWord = True
	bUseSpecificTable = False
	bSaveFile = False
	Set objArgs = WScript.Arguments
	If (objArgs.Count > 0) Then
		For nArgs = 0 To objArgs.Count - 1
			SetOptions objArgs(nArgs)
		Next
	Else
		bDisplayHelp = True
	End If
	
	CheckFiles
	
End Sub ' GetOptions

Sub GetSYDIServerSourceData(strXMLSource)
	Dim objXMLFile, colNodes, objNode
	Dim objChild, objChild2, objChild3
	Set objXMLFile = CreateObject("Microsoft.XMLDOM")
	objXMLFile.async = False
	objXMLFile.load(strXMLSource)
	ReportProgress "Reading SYDI-Server Source File: " & strXMLSource
	Set colNodes = objXMLFile.selectNodes("//computer")
	For Each objNode in colNodes
		For Each objChild in objNode.childNodes
			If (objChild.nodeName = "generated") Then 
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strScanTime" 
				objDbrVariables("Value") = Scrub(objChild.getAttribute("scantime"))
				objDbrVariables.Update
			End If
			If (objChild.nodeName = "system") Then
				strSystem_Name = Scrub(objChild.getAttribute("name"))
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strNetBiosName"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("name"))
				objDbrVariables.Update				
			End If
			If (objChild.nodeName = "operatingsystem") Then
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strOperatingSystem"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("name"))
				objDbrVariables.Update				
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strServicePack"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("servicepack"))
				objDbrVariables.Update				
			End If
			If (objChild.nodeName = "fqdn") Then
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strFQDN"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("name"))
				objDbrVariables.Update				
			End If
			If (objChild.nodeName = "roles") Then
				strTempConcatString = ""
				For Each objChild2 in objChild.childNodes
					If (objChild2.nodeName = "role") Then
						If (strTempConcatString = "") Then
							strTempConcatString = Scrub(objChild2.getAttribute("name"))
						Else
							strTempConcatString = strTempConcatString & ", " & Scrub(objChild2.getAttribute("name"))
						End If
					End If
				Next
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strRoles"
				objDbrVariables("Value") = strTempConcatString
				objDbrVariables.Update
			End If
			If (objChild.nodeName = "machineinfo") Then
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strManufacturer"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("manufacturer"))
				objDbrVariables.Update				
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strComputerProductname"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("productname"))
				objDbrVariables.Update				
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strIdentifyingNumber"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("identifyingnumber"))
				objDbrVariables.Update				
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strComputerChassis"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("chassis"))
				objDbrVariables.Update				
			End If
			If (objChild.nodeName = "processor") Then
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strProcessorName"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("name"))
				objDbrVariables.Update				
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strProcessorDescription"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("description"))
				objDbrVariables.Update				
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strProcessorSpeed"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("speed"))
				objDbrVariables.Update				
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strProcessorL2Cache"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("l2cachesize"))
				objDbrVariables.Update				
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strProcessorExternalClock"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("externalclock"))
				objDbrVariables.Update				
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strProcessorCount"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("count"))
				objDbrVariables.Update			
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strProcessorHT"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("htsystem"))
				objDbrVariables.Update			
			End If
			If (objChild.nodeName = "memory") Then
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strTotalMemory"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("totalsize"))
				objDbrVariables.Update
				For Each objChild2 in objChild.childNodes
					If (objChild2.nodeName = "memorybank") Then
						objDbrRecordsets.AddNew
						objDbrRecordsets("Name") = "dbrMemory"
						objDbrRecordsets("Column1") = Scrub(objChild2.getAttribute("bank"))
						objDbrRecordsets("Column2") = Scrub(objChild2.getAttribute("capacity"))
						objDbrRecordsets("Column3") = Scrub(objChild2.getAttribute("formfactor"))
						objDbrRecordsets("Column4") = Scrub(objChild2.getAttribute("memorytype"))
						objDbrRecordsets.Update	
					End If
				Next
			End If
			If (objChild.nodeName = "win32_cdromdrive") Then
				For Each objChild2 in objChild.childNodes
					If (objChild2.nodeName = "cdrom") Then
						objDbrRecordsets.AddNew
						objDbrRecordsets("Name") = "dbrCDROM"
						objDbrRecordsets("Column1") = Scrub(objChild2.getAttribute("name"))
						objDbrRecordsets("Column2") = Scrub(objChild2.getAttribute("drive"))
						objDbrRecordsets("Column3") = Scrub(objChild2.getAttribute("manufacturer"))
						objDbrRecordsets.Update	
					End If
				Next
			End If
			If (objChild.nodeName = "win32_tapedrive") Then
				For Each objChild2 in objChild.childNodes
					If (objChild2.nodeName = "tapedrive") Then
						objDbrRecordsets.AddNew
						objDbrRecordsets("Name") = "dbrTapeDrive"
						objDbrRecordsets("Column1") = Scrub(objChild2.getAttribute("name"))
						objDbrRecordsets("Column2") = Scrub(objChild2.getAttribute("description"))
						objDbrRecordsets("Column3") = Scrub(objChild2.getAttribute("manufacturer"))
						objDbrRecordsets.Update	
					End If
				Next
			End If
			If (objChild.nodeName = "win32_sounddevice") Then
				For Each objChild2 in objChild.childNodes
					If (objChild2.nodeName = "card") Then
						objDbrRecordsets.AddNew
						objDbrRecordsets("Name") = "dbrSoundCard"
						objDbrRecordsets("Column1") = Scrub(objChild2.getAttribute("name"))
						objDbrRecordsets("Column2") = Scrub(objChild2.getAttribute("manufacturer"))
						objDbrRecordsets.Update	
					End If
				Next
			End If
			If (objChild.nodeName = "videocontroller") Then
				For Each objChild2 in objChild.childNodes
					If (objChild2.nodeName = "adapter") Then
						objDbrRecordsets.AddNew
						objDbrRecordsets("Name") = "dbrVideoController"
						objDbrRecordsets("Column1") = Scrub(objChild2.getAttribute("name"))
						objDbrRecordsets("Column2") = Scrub(objChild2.getAttribute("adapterram"))
						objDbrRecordsets("Column3") = Scrub(objChild2.getAttribute("compatibility"))
						objDbrRecordsets.Update	
					End If
				Next
			End If
			If (objChild.nodeName = "bios") Then
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strBiosVersion"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("version"))
				objDbrVariables.Update
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strBiosSMVersion"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("smbiosversion")) & " (" & _
					Scrub(objChild.getAttribute("smbiosmajorversion")) & "," & Scrub(objChild.getAttribute("smbbiosminorversion")) & ")"
				objDbrVariables.Update
				strTempConcatString = ""
				For Each objChild2 in objChild.childNodes
					If (objChild2.nodeName = "bioscharacteristics") Then
						If (strTempConcatString = "") Then
							strTempConcatString = Scrub(objChild2.getAttribute("name"))
						Else
							strTempConcatString = strTempConcatString & ", " & Scrub(objChild2.getAttribute("name"))
						End If
					End If
				Next
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strBiosCharacteristics"
				objDbrVariables("Value") = strTempConcatString
				objDbrVariables.Update
			End If		
			If (objChild.nodeName = "osconfiguration") Then
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strOSName"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("osname"))
				objDbrVariables.Update
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strComputerRole"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("computerrole"))
				objDbrVariables.Update
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strComputerDomainType"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("domaintype"))
				objDbrVariables.Update
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strDomainName"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("domainname"))
				objDbrVariables.Update
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strInstallLocation"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("windowslocation"))
				objDbrVariables.Update
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strOSLanguage"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("oslanguage"))
				objDbrVariables.Update
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strInstallDate"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("installdate"))
				objDbrVariables.Update
			End If
			If (objChild.nodeName = "lastuser") Then
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strLastUser"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("name"))
				objDbrVariables.Update
			End If
			If (objChild.nodeName = "windowscomponents") Then
				For Each objChild2 in objChild.childNodes
					If (objChild2.nodeName = "component") Then
						objDbrRecordsets.AddNew
						objDbrRecordsets("Name") = "dbrWindowsComponents"
						objDbrRecordsets("Column1") = Scrub(objChild2.getAttribute("name"))
						objDbrRecordsets("Column2") = Scrub(objChild2.getAttribute("classname"))
						objDbrRecordsets.Update	
					End If
				Next
			End If
			If (objChild.nodeName = "patches") Then
				For Each objChild2 in objChild.childNodes
					If (objChild2.nodeName = "patch") Then
						objDbrRecordsets.AddNew
						objDbrRecordsets("Name") = "dbrPatches"
						objDbrRecordsets("Column1") = Scrub(objChild2.getAttribute("hotfixid"))
						objDbrRecordsets("Column2") = Scrub(objChild2.getAttribute("description"))
						objDbrRecordsets("Column3") = Scrub(objChild2.getAttribute("installdate"))
						objDbrRecordsets.Update	
					End If
				Next
			End If
			If (objChild.nodeName = "installedapplications") Then
				For Each objChild2 in objChild.childNodes
					If (objChild2.nodeName = "msiapplication") Then
						objDbrRecordsets.AddNew
						objDbrRecordsets("Name") = "dbrMSIApplications"
						objDbrRecordsets("Column1") = Scrub(objChild2.getAttribute("productname"))
						objDbrRecordsets("Column2") = Scrub(objChild2.getAttribute("vendor"))
						objDbrRecordsets("Column3") = Scrub(objChild2.getAttribute("version"))
						objDbrRecordsets("Column4") = Scrub(objChild2.getAttribute("installdate"))
						objDbrRecordsets.Update	
					End If
					If (objChild2.nodeName = "regapplication") Then
						objDbrRecordsets.AddNew
						objDbrRecordsets("Name") = "dbrRegApplications"
						objDbrRecordsets("Column1") = Scrub(objChild2.getAttribute("productname"))
						objDbrRecordsets("Column2") = Scrub(objChild2.getAttribute("version"))
						objDbrRecordsets.Update	
					End If
					If (objChild2.nodeName = "productkey") Then
						objDbrRecordsets.AddNew
						objDbrRecordsets("Name") = "dbrProductKey"
						objDbrRecordsets("Column1") = Scrub(objChild2.getAttribute("productname"))
						objDbrRecordsets("Column2") = Scrub(objChild2.getAttribute("productkey"))
						objDbrRecordsets.Update	
					End If
				Next
			End If
			If (objChild.nodeName = "storage") Then
				For Each objChild2 in objChild.childNodes
					If (objChild2.nodeName = "drives") Then
						i = 0
						For Each objChild3 in objChild2.childNodes
							If (objChild3.nodeName = "partition") Then
								objDbrRecordsets.AddNew
								objDbrRecordsets("Name") = "dbrStorage"
								objDbrRecordsets("Column1") = Scrub(objChild2.getAttribute("name"))
								objDbrRecordsets("Column2") = Scrub(objChild2.getAttribute("deviceid"))
								objDbrRecordsets("Column3") = Scrub(objChild2.getAttribute("interface"))
								objDbrRecordsets("Column4") = Scrub(objChild2.getAttribute("totaldisksize"))
								objDbrRecordsets("Column5") = Scrub(objChild3.getAttribute("name"))
								objDbrRecordsets("Column6") = Scrub(objChild3.getAttribute("size"))
								objDbrRecordsets("Column7") = Scrub(objChild3.getAttribute("freespace"))
								objDbrRecordsets("Column8") = Scrub(objChild3.getAttribute("filesystem"))
								objDbrRecordsets.Update
								i = i + 1
							End If
						Next
						If (i = 0) Then
							objDbrRecordsets.AddNew
							objDbrRecordsets("Column1") = Scrub(objChild2.getAttribute("name"))
							objDbrRecordsets("Column2") = Scrub(objChild2.getAttribute("deviceid"))
							objDbrRecordsets("Column3") = Scrub(objChild2.getAttribute("interface"))
							objDbrRecordsets("Column4") = Scrub(objChild2.getAttribute("totaldisksize"))
							objDbrRecordsets.Update
						End If
					End If
				Next
			End If
			If (objChild.nodeName = "network") Then
				For Each objChild2 In objChild.childNodes
					If (objChild2.nodeName = "ip4routes") Then
						For Each objChild3 In objChild2.childNodes
							If (objChild3.nodeName = "route") Then
								objDbrRecordsets.AddNew
								objDbrRecordsets("Name") = "dbrIPRoutes"
								objDbrRecordsets("Column1") = Scrub(objChild3.getAttribute("destination"))
								objDbrRecordsets("Column2") = Scrub(objChild3.getAttribute("mask"))
								objDbrRecordsets("Column3") = Scrub(objChild3.getAttribute("nexthop"))
								objDbrRecordsets.Update
							End If
						Next
					Elseif (objChild2.nodeName = "adapter") Then
						objDbrRecordsets.AddNew
						objDbrRecordsets("Name") = "dbrIPConfiguration"
						objDbrRecordsets("Column1") = Scrub(objChild2.getAttribute("description"))
						objDbrRecordsets("Column2") = Scrub(objChild2.getAttribute("macaddress"))
						For Each objChild3 In objChild2.childNodes
							If (objChild3.nodeName = "ip") Then
								If (objDbrRecordsets("Column3") = "") Then
									objDbrRecordsets("Column3") = Scrub(objChild3.getAttribute("address")) & "/" & Scrub(objChild3.getAttribute("subnetmask"))
								Else
									objDbrRecordsets("Column3") = objDbrRecordsets("Column3") & ", " & Scrub(objChild3.getAttribute("address"))& "/" & Scrub(objChild3.getAttribute("subnetmask"))
								End If
							Elseif (objChild3.nodeName = "gateway") Then
								If (objDbrRecordsets("Column4") = "") Then
									objDbrRecordsets("Column4") = Scrub(objChild3.getAttribute("address"))
								Else
									objDbrRecordsets("Column4") = objDbrRecordsets("Column4") & ", " & Scrub(objChild3.getAttribute("address"))
								End If
							Elseif (objChild3.nodeName = "dnsserver") Then
								If (objDbrRecordsets("Column5") = "") Then
									objDbrRecordsets("Column5") = Scrub(objChild3.getAttribute("address"))
								Else
									objDbrRecordsets("Column5") = objDbrRecordsets("Column5") & ", " & Scrub(objChild3.getAttribute("address"))
								End If
							Elseif (objChild3.nodeName = "dnsdomain") Then
								If (objDbrRecordsets("Column6") = "") Then
									objDbrRecordsets("Column6") = Scrub(objChild3.getAttribute("name"))
								Else
									objDbrRecordsets("Column6") = objDbrRecordsets("Column6") & ", " & Scrub(objChild3.getAttribute("name"))
								End If
							Elseif (objChild3.nodeName = "primarywins") Then
								If (objDbrRecordsets("Column7") = "") Then
									objDbrRecordsets("Column7") = Scrub(objChild3.getAttribute("address"))
								Else
									objDbrRecordsets("Column7") = objDbrRecordsets("Column7") & ", " & Scrub(objChild3.getAttribute("address"))
								End If
							Elseif (objChild3.nodeName = "secondarywins") Then
								If (objDbrRecordsets("Column8") = "") Then
									objDbrRecordsets("Column8") = Scrub(objChild3.getAttribute("address"))
								Else
									objDbrRecordsets("Column8") = objDbrRecordsets("Column8") & ", " & Scrub(objChild3.getAttribute("address"))
								End If
							Elseif (objChild3.nodeName = "dhcpserver") Then
								If (objDbrRecordsets("Column9") = "") Then
									objDbrRecordsets("Column9") = Scrub(objChild3.getAttribute("address"))
								Else
									objDbrRecordsets("Column9") = objDbrRecordsets("Column9") &  ", " & Scrub(objChild3.getAttribute("address"))
								End If
							End If
						Next
						objDbrRecordsets.Update
					End If
				Next
			End If
						
			If (objChild.nodeName = "microsoftiisv2") Then
				For Each objChild2 in objChild.childNodes
					If (objChild2.nodeName = "iiswebserversetting") Then
						objDbrRecordsets.AddNew
						objDbrRecordsets("Name") = "dbrIISWebServer"
						objDbrRecordsets("Column1") = Scrub(objChild2.getAttribute("name"))
						objDbrRecordsets("Column2") = Scrub(objChild2.getAttribute("servercomment"))
						objDbrRecordsets.Update	

						For Each objChild3 in objChild2.childNodes
							If (objChild3.nodeName = "homedirectory") Then
								objDbrIISVariables.AddNew
								objDbrIISVariables("Server") =  Scrub(objChild2.getAttribute("name"))
								objDbrIISVariables("Name") =  "strHomeDirectory"
								objDbrIISVariables("Value") =  Scrub(objChild3.getAttribute("path"))
								objDbrIISVariables.Update
							End If
							If (objChild3.nodeName = "serverbindings") Then
								objDbrIISRecordsets.AddNew
								objDbrIISRecordsets("Server") =  Scrub(objChild2.getAttribute("name"))
								objDbrIISRecordsets("Name") =  "dbrServerBindings"
								objDbrIISRecordsets("Column1") =  Scrub(objChild3.getAttribute("hostname"))
								objDbrIISRecordsets("Column2") =  Scrub(objChild3.getAttribute("ip"))
								objDbrIISRecordsets("Column3") =  Scrub(objChild3.getAttribute("port"))
								objDbrIISRecordsets.Update
							End If
						Next
					End If
				Next
			End If
			If (objChild.nodeName = "eventlogfiles") Then
				For Each objChild2 in objChild.childNodes
					If (objChild2.nodeName = "eventlogfile") Then
						objDbrRecordsets.AddNew
						objDbrRecordsets("Name") = "dbrEventLogFiles"
						objDbrRecordsets("Column1") = Scrub(objChild2.getAttribute("name"))
						objDbrRecordsets("Column2") = Scrub(objChild2.getAttribute("file"))
						objDbrRecordsets("Column3") = Scrub(objChild2.getAttribute("maximumsize"))
						objDbrRecordsets("Column4") = Scrub(objChild2.getAttribute("overwritepolicy"))
						objDbrRecordsets.Update	
					End If
				Next
			End If			
			If (objChild.nodeName = "localgroups") Then
				For Each objChild2 in objChild.childNodes
					If (objChild2.nodeName = "group") Then
						i = 0
						For Each objChild3 in objChild2.childNodes
							If (objChild3.nodeName = "member") Then
								objDbrRecordsets.AddNew
								objDbrRecordsets("Name") = "dbrLocalGroups"
								objDbrRecordsets("Column1") = Scrub(objChild3.getAttribute("name"))
								objDbrRecordsets("Column2") = Scrub(objChild2.getAttribute("name"))
								objDbrRecordsets.Update
								i = i + 1
							End If
						Next
						If (i = 0) Then
							objDbrRecordsets.AddNew
							objDbrRecordsets("Name") = "dbrLocalGroups"
							objDbrRecordsets("Column1") = Scrub(objChild2.getAttribute("name"))
							objDbrRecordsets("Column2") = ""
							objDbrRecordsets.Update
						End If
					End If
				Next
			End If
			If (objChild.nodeName = "localusers") Then
				For Each objChild2 in objChild.childNodes
					If (objChild2.nodeName = "user") Then
						objDbrRecordsets.AddNew
						objDbrRecordsets("Name") = "dbrLocalUsers"
						objDbrRecordsets("Column1") = Scrub(objChild2.getAttribute("name"))
						objDbrRecordsets.Update	
					End If
				Next
			End If
			If (objChild.nodeName = "printspooler") Then
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strPrintSpooler"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("location"))
				objDbrVariables.Update
			End If
			If (objChild.nodeName = "printers") Then
				For Each objChild2 in objChild.childNodes
					If (objChild2.nodeName = "printer") Then
						objDbrRecordsets.AddNew
						objDbrRecordsets("Name") = "dbrPrinters"
						objDbrRecordsets("Column1") = Scrub(objChild2.getAttribute("name"))
						objDbrRecordsets("Column2") = Scrub(objChild2.getAttribute("drivername"))
						objDbrRecordsets("Column3") = Scrub(objChild2.getAttribute("portname"))
						objDbrRecordsets.Update	
					End If
				Next
			End If
			If (objChild.nodeName = "regional") Then
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strTimeZone"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("timezone"))
				objDbrVariables.Update
			End If			
			If (objChild.nodeName = "processes") Then
				For Each objChild2 in objChild.childNodes
					If (objChild2.nodeName = "process") Then
						objDbrRecordsets.AddNew
						objDbrRecordsets("Name") = "dbrProcesses"
						objDbrRecordsets("Column1") = Scrub(objChild2.getAttribute("caption"))
						objDbrRecordsets("Column2") = Scrub(objChild2.getAttribute("executablepath"))
						objDbrRecordsets.Update	
					End If
				Next
			End If

			If (objChild.nodeName = "services") Then
				For Each objChild2 in objChild.childNodes
					If (objChild2.nodeName = "service") Then
						objDbrRecordsets.AddNew
						objDbrRecordsets("Name") = "dbrServices"
						objDbrRecordsets("Column1") = Scrub(objChild2.getAttribute("name"))
						objDbrRecordsets("Column2") = Scrub(objChild2.getAttribute("startmode"))
						objDbrRecordsets("Column3") = Scrub(objChild2.getAttribute("started"))
						objDbrRecordsets("Column4") = Scrub(objChild2.getAttribute("startname"))
						objDbrRecordsets.Update	
					End If
				Next
			End If			
			If (objChild.nodeName = "shares") Then
				For Each objChild2 in objChild.childNodes
					If (objChild2.nodeName = "share") Then
						objDbrRecordsets.AddNew
						objDbrRecordsets("Name") = "dbrShares"
						objDbrRecordsets("Column1") = Scrub(objChild2.getAttribute("name"))
						objDbrRecordsets("Column2") = Scrub(objChild2.getAttribute("path"))
						objDbrRecordsets("Column3") = Scrub(objChild2.getAttribute("description"))
						objDbrRecordsets.Update	
					End If
				Next
			End If			

			If (objChild.nodeName = "win32_startupcommand") Then
				For Each objChild2 in objChild.childNodes
					If (objChild2.nodeName = "command") Then
						objDbrRecordsets.AddNew
						objDbrRecordsets("Name") = "dbrStartupCommands"
						objDbrRecordsets("Column1") = Scrub(objChild2.getAttribute("user"))
						objDbrRecordsets("Column2") = Scrub(objChild2.getAttribute("name"))
						objDbrRecordsets("Column3") = Scrub(objChild2.getAttribute("command"))
						objDbrRecordsets.Update	
					End If
				Next
			End If
			If (objChild.nodeName = "pagefiles") Then
				For Each objChild2 in objChild.childNodes
					If (objChild2.nodeName = "pagefile") Then
						objDbrRecordsets.AddNew
						objDbrRecordsets("Name") = "dbrPagefiles"
						objDbrRecordsets("Column1") = Scrub(objChild2.getAttribute("drive"))
						objDbrRecordsets("Column2") = Scrub(objChild2.getAttribute("initialsize"))
						objDbrRecordsets("Column3") = Scrub(objChild2.getAttribute("maximumsize"))
						objDbrRecordsets.Update	
					End If
				Next
			End If
			
			If (objChild.nodeName = "registry") Then
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strRegistrySize"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("currentsize"))
				objDbrVariables.Update
				objDbrVariables.AddNew
				objDbrVariables("Name") = "strMaxRegistrySize"
				objDbrVariables("Value") = Scrub(objChild.getAttribute("maximumsize"))
				objDbrVariables.Update
			End If			
			
		Next
	Next
End Sub ' GetSYDIServerSourceData

Sub GetWrittenDocumentationData(strXMLSource)
	Dim objXMLFile, colNodes, objNode
	Dim objChild, objChild2, objSectionElements
	Dim objNote
	Set objXMLFile = CreateObject("Microsoft.XMLDOM")
	objXMLFile.async = False
	objXMLFile.load(strXMLSource)
	ReportProgress "Reading Written Documentation: " & strXMLSource
	Set colNodes = objXMLFile.selectNodes("//document")
	For Each objNode in colNodes
			For Each objChild in objNode.childNodes
				
				If (objChild.nodeName = "section") Then
					objDbrSections.AddNew
					objDbrSections("SectionName") = objChild.getAttribute("name")
					objDbrSections("SubSection") = False
					objDbrSections("ParentSection") = "None"
					objDbrSections.Update
					For Each objSectionElements in objChild.childNodes
						If (objSectionElements.nodeName = "prenotes" OR objSectionElements.nodeName = "postnotes") Then
							For Each objNote in objSectionElements.childNodes
								If (objNote.nodeName = "text") Then
									objDbrWrittenText.AddNew
									objDbrWrittenText("TextField") = objNote.text
									objDbrWrittenText("SectionName") = objChild.getAttribute("name")
									objDbrWrittenText("ParentSection") = False
									objDbrWrittenText("Position") = objSectionElements.nodeName
									objDbrWrittenText.Update
								End If
							Next
						End If
						If (objSectionElements.nodeName = "subsection") Then
							objDbrSections.AddNew
							objDbrSections("SectionName") = objSectionElements.getAttribute("name")
							objDbrSections("SubSection") = True
							objDbrSections("ParentSection") = objChild.getAttribute("name")
							objDbrSections.Update
							
							For Each objChild2 in objSectionElements.childNodes
								If (objChild2.nodeName = "prenotes" OR objChild2.nodeName = "postnotes") Then
									For Each objNote in objChild2.childNodes
										If (objNote.nodeName = "text") Then
											objDbrWrittenText.AddNew
											objDbrWrittenText("TextField") = objNote.text
											objDbrWrittenText("SectionName") = objSectionElements.getAttribute("name")
											objDbrWrittenText("ParentSection") = objChild.getAttribute("name")
											objDbrWrittenText("Position") = objChild2.nodeName
											objDbrWrittenText.Update
										End If
									Next
								End If
							Next
						End If
					Next
				End If
				
			Next
	Next
End Sub '  GetWrittenDocumentationData

Sub Init()
	Set objDbrSections = CreateObject("ADOR.Recordset")
	objDbrSections.Fields.Append "SectionName", adVarChar, MaxCharacters
	objDbrSections.Fields.Append "SubSection", adVarChar, MaxCharacters
	objDbrSections.Fields.Append "ParentSection", adVarChar, MaxCharacters
	objDbrSections.Open
	Set objDbrWrittenText = CreateObject("ADOR.Recordset")
	objDbrWrittenText.Fields.Append "TextField", adLongVarChar, 60000
	objDbrWrittenText.Fields.Append "SectionName", adVarChar, MaxCharacters
	objDbrWrittenText.Fields.Append "ParentSection", adVarChar, MaxCharacters
	objDbrWrittenText.Fields.Append "Position", adVarChar, MaxCharacters
	objDbrWrittenText.Open
	Set objDbrVariables = CreateObject("ADOR.Recordset")
	objDbrVariables.Fields.Append "Name", adVarChar, MaxCharacters
	objDbrVariables.Fields.Append "Value", adVarChar, LargeMaxCharacters
	objDbrVariables.Open
	Set objDbrRecordsets = CreateObject("ADOR.Recordset")
	objDbrRecordsets.Fields.Append "Name", adVarChar, MaxCharacters
	objDbrRecordsets.Fields.Append "Column1", adVarChar, MaxCharacters
	objDbrRecordsets.Fields.Append "Column2", adVarChar, MaxCharacters
	objDbrRecordsets.Fields.Append "Column3", adVarChar, MaxCharacters
	objDbrRecordsets.Fields.Append "Column4", adVarChar, MaxCharacters
	objDbrRecordsets.Fields.Append "Column5", adVarChar, MaxCharacters
	objDbrRecordsets.Fields.Append "Column6", adVarChar, MaxCharacters
	objDbrRecordsets.Fields.Append "Column7", adVarChar, MaxCharacters
	objDbrRecordsets.Fields.Append "Column8", adVarChar, MaxCharacters
	objDbrRecordsets.Fields.Append "Column9", adVarChar, MaxCharacters
	objDbrRecordsets.Open
	Set objDbrIISVariables = CreateObject("ADOR.Recordset")
	objDbrIISVariables.Fields.Append "Server", adVarChar, MaxCharacters
	objDbrIISVariables.Fields.Append "Name", adVarChar, MaxCharacters
	objDbrIISVariables.Fields.Append "Value", adVarChar, LargeMaxCharacters
	objDbrIISVariables.Open
	Set objDbrIISRecordsets = CreateObject("ADOR.Recordset")
	objDbrIISRecordsets.Fields.Append "Name", adVarChar, MaxCharacters
	objDbrIISRecordsets.Fields.Append "Server", adVarChar, MaxCharacters
	objDbrIISRecordsets.Fields.Append "Column1", adVarChar, MaxCharacters
	objDbrIISRecordsets.Fields.Append "Column2", adVarChar, MaxCharacters
	objDbrIISRecordsets.Fields.Append "Column3", adVarChar, MaxCharacters
	objDbrIISRecordsets.Fields.Append "Column4", adVarChar, MaxCharacters
	objDbrIISRecordsets.Open
	
	
	
End Sub 'Init

Sub ParseLanguageFile(strXMLFile)
	Dim objXMLFile, colNodes, objNode
	Dim objSection, objChild, objChild2, colChild
	Set objXMLFile = CreateObject("Microsoft.XMLDOM")
	objXMLFile.async = False
	objXMLFile.load(strXMLFile)
	
	ReportProgress "Reading Language File: " & strXMLFile
	Set colNodes = objXMLFile.selectNodes("//language")
	For Each objNode in colNodes
		For Each objSection in objNode.childNodes
			If (objSection.nodeName = "section") Then
				Select Case objSection.getAttribute("type")
					Case "title"
						objDbrWrittenText.Filter = " SectionName='title' and position='prenotes' "
						Do Until objDbrWrittenText.Eof
							WriteText objDbrWrittenText("TextField")
							objDbrWrittenText.MoveNext
						Loop
						oWord.Selection.Style = wdStyleTitle
						oWord.Selection.TypeText Cstr(Replace(objSection.getAttribute("translation"),"%computer%",strSystem_Name)) & VbCrLf
						oWord.Selection.Style = wdStyleBodyText
						objDbrWrittenText.Filter = " SectionName='title' and position='postnotes' "
						Do Until objDbrWrittenText.Eof
							WriteText objDbrWrittenText("TextField")
							objDbrWrittenText.MoveNext
						Loop
					Case "note"
						objDbrWrittenText.Filter = " SectionName='" & objSection.getAttribute("name") & "' and position='prenotes' and parentsection=false "
						Do Until objDbrWrittenText.Eof
							WriteText objDbrWrittenText("TextField")
							objDbrWrittenText.MoveNext
						Loop

						If (objSection.getAttribute("translation") <> "") Then
							oWord.Selection.Font.Bold = True
							WriteText objSection.getAttribute("translation")
							oWord.Selection.Font.Bold = False
						End If
						For Each objChild in objSection.childNodes
							If (objChild.nodeName = "property") Then
								Select Case objChild.getAttribute("type")
									Case "2levellist"
										If (CheckCriteria(objChild.getAttribute("ctype"),objChild.getAttribute("criteria"))) Then Write2LevelList objChild.getAttribute("name")
									Case "boldtext"
										If (CheckCriteria(objChild.getAttribute("ctype"),objChild.getAttribute("criteria"))) Then WriteBold objChild.getAttribute("translation")
									Case "list"
										If (CheckCriteria(objChild.getAttribute("ctype"),objChild.getAttribute("criteria"))) Then WriteList objChild.getAttribute("name")
									Case "loop"
										If (CheckCriteria(objChild.getAttribute("ctype"),objChild.getAttribute("criteria"))) Then WriteLoop objChild.getAttribute("name"), objChild.getAttribute("columns"), objChild.getAttribute("translation")
									Case "space"
										If (CheckCriteria(objChild.getAttribute("ctype"),objChild.getAttribute("criteria"))) Then WriteSpace
									Case "storage"
										WriteStorage objChild.getAttribute("translation")
									Case "string"
										WriteString objChild.getAttribute("name"), objChild.getAttribute("translation")
									Case "table"
										If (CheckCriteria(objChild.getAttribute("ctype"),objChild.getAttribute("criteria"))) Then WriteTable objChild.getAttribute("name"), objChild.getAttribute("columns"), objChild.getAttribute("translation")
										
								End Select
							End If
						Next
						
						objDbrWrittenText.Filter = " SectionName='" & objSection.getAttribute("name") & "' and position='postnotes' and parentsection=false "
						Do Until objDbrWrittenText.Eof
							WriteText objDbrWrittenText("TextField")
							objDbrWrittenText.MoveNext
						Loop
					Case "toc"
						objDbrWrittenText.Filter = " SectionName='toc' and position='prenotes' and parentsection=false "
						Do Until objDbrWrittenText.Eof
							WriteText objDbrWrittenText("TextField")
							objDbrWrittenText.MoveNext
						Loop
						oWord.Selection.Font.Bold = True
						oWord.Selection.TypeText vbCrLf & objSection.getAttribute("translation") & vbCrLf
						oWord.Selection.Font.Bold = False
						oWord.ActiveDocument.TablesOfContents.Add oWord.Selection.Range, False, 2, 3, , , , ,oWord.ActiveDocument.Styles(wdStyleHeading1)& ";1", True
						ReportProgress " Inserting Table Of Contents"
						oWord.Selection.TypeText vbCrLf
						objDbrWrittenText.Filter = " SectionName='toc' and position='postnotes' and parentsection=false "
						Do Until objDbrWrittenText.Eof
							WriteText objDbrWrittenText("TextField")
							objDbrWrittenText.MoveNext
						Loop
						oWord.Selection.InsertBreak wdPageBreak

					Case "heading1"
						If (CheckCriteria(objSection.getAttribute("ctype"),objSection.getAttribute("criteria"))) Then WriteHeader 1, objSection.getAttribute("translation")
						objDbrWrittenText.Filter = " SectionName='" & objSection.getAttribute("name") & "' and position='prenotes' and parentsection=false "
						Do Until objDbrWrittenText.Eof
							WriteText objDbrWrittenText("TextField")
							objDbrWrittenText.MoveNext
						Loop
						objDbrWrittenText.Filter = " SectionName='" & objSection.getAttribute("name") & "' and position='postnotes' and parentsection=false "
						Do Until objDbrWrittenText.Eof
							WriteText objDbrWrittenText("TextField")
							objDbrWrittenText.MoveNext
						Loop
						
						For Each objChild in objSection.childNodes
							If (objChild.nodeName = "subsection") Then
								If (CheckCriteria(objChild.getAttribute("ctype"),objChild.getAttribute("criteria"))) Then WriteHeader 2, objChild.getAttribute("translation")
								objDbrWrittenText.Filter = " SectionName='" & objChild.getAttribute("name") & "' and position='prenotes' and parentsection='" & objSection.getAttribute("name") & "' "
								Do Until objDbrWrittenText.Eof
									WriteText objDbrWrittenText("TextField")
									objDbrWrittenText.MoveNext
								Loop

								For Each objChild2 in objChild.childNodes
									If (objChild2.nodeName = "property") Then
										Select Case objChild2.getAttribute("type")
											Case "2levellist"
												If (CheckCriteria(objChild2.getAttribute("ctype"),objChild2.getAttribute("criteria"))) Then Write2LevelList objChild2.getAttribute("name")
											Case "boldtext"
												If (CheckCriteria(objChild2.getAttribute("ctype"),objChild2.getAttribute("criteria"))) Then WriteBold objChild2.getAttribute("translation")
											Case "ipconfiguration"
												WriteIPConfiguration objChild2.getAttribute("translation")
											Case "list"
												If (CheckCriteria(objChild2.getAttribute("ctype"),objChild2.getAttribute("criteria"))) Then WriteList objChild2.getAttribute("name")
											Case "loop"
												If (CheckCriteria(objChild2.getAttribute("ctype"),objChild2.getAttribute("criteria"))) Then WriteLoop objChild2.getAttribute("name"), objChild2.getAttribute("columns"), objChild2.getAttribute("translation")
											Case "space"
												If (CheckCriteria(objChild2.getAttribute("ctype"),objChild2.getAttribute("criteria"))) Then WriteSpace
											Case "storage"
												If (CheckCriteria(objChild2.getAttribute("ctype"),objChild2.getAttribute("criteria"))) Then WriteStorage objChild2.getAttribute("translation")
											Case "string"
												WriteString objChild2.getAttribute("name"), objChild2.getAttribute("translation")
											Case "table"
												If (CheckCriteria(objChild2.getAttribute("ctype"),objChild2.getAttribute("criteria"))) Then WriteTable objChild2.getAttribute("name"), objChild2.getAttribute("columns"), objChild2.getAttribute("translation")
											Case "webserver"
												If (CheckCriteria(objChild2.getAttribute("ctype"),objChild2.getAttribute("criteria"))) Then WriteWebServer objChild2.getAttribute("translation")
										End Select
									End If
								Next
								objDbrWrittenText.Filter = " SectionName='" & objChild.getAttribute("name") & "' and position='postnotes' and parentsection='" & objSection.getAttribute("name") & "' "
								Do Until objDbrWrittenText.Eof
									WriteText objDbrWrittenText("TextField")
									objDbrWrittenText.MoveNext
								Loop

							End If
							If (objChild.nodeName = "property") Then
								Select Case objChild.getAttribute("type")
									Case "2levellist"
										If (CheckCriteria(objChild2.getAttribute("ctype"),objChild2.getAttribute("criteria"))) Then Write2LevelList objChild.getAttribute("name")
									Case "boldtext"
										If (CheckCriteria(objChild2.getAttribute("ctype"),objChild2.getAttribute("criteria"))) Then WriteBold objChild.getAttribute("translation")
									Case "list"
										If (CheckCriteria(objChild2.getAttribute("ctype"),objChild2.getAttribute("criteria"))) Then WriteList objChild.getAttribute("name")
									Case "loop"
										If (CheckCriteria(objChild2.getAttribute("ctype"),objChild2.getAttribute("criteria"))) Then WriteLoop objChild.getAttribute("name"), objChild.getAttribute("columns"), objChild.getAttribute("translation")
									Case "space"
										If (CheckCriteria(objChild2.getAttribute("ctype"),objChild2.getAttribute("criteria"))) Then WriteSpace
									Case "storage"
										WriteStorage objChild.getAttribute("translation")
									Case "string"
										WriteString objChild.getAttribute("name"), objChild.getAttribute("translation")
									Case "table"
										If (CheckCriteria(objChild2.getAttribute("ctype"),objChild2.getAttribute("criteria"))) Then WriteTable objChild.getAttribute("name"), objChild.getAttribute("columns"), objChild.getAttribute("translation")
										
								End Select
							End If

						Next
						
				End Select
			End If
		Next
	Next
	
End Sub ' ParseLanguageFile

Sub ReportProgress(strMessage)
	WScript.Echo strMessage
End Sub ' ReportProgress


Function Scrub(strInput)
	If (IsNull(strInput)) Then
		strInput = ""
	End If
	Scrub = strInput
End Function ' Scrub

Sub SetOptions(strOption)
	Dim strFlag, strParameter
	Dim nArguments
	nArguments = Len(strOption)
	If (nArguments < 2) Then
		bInvalidArgument = True
	Else
		strFlag = Left(strOption,2)
		Select Case strFlag
			Case "-b"
				strWordTable = ""
				bUseSpecificTable = True
				If (nArguments > 2) Then
					strWordTable = Right(strOption,(nArguments - 2))
				End If
				If (strWordTable = "") Then
					bInvalidArgument = True
				End If
			Case "-d"
					bShowWord = False
			Case "-f"
				If (nArguments > 2) Then
					nBaseFontSize = Right(strOption,(nArguments - 2))
					If Not (IsNumeric(nBaseFontSize)) Then
						bInvalidArgument = True
					End If
				End If
			Case "-l"
				strLanguageFile = ""
				If (nArguments > 2) Then
					strLanguageFile = Right(strOption,(nArguments - 2))
				End If
			Case "-o"
					bSaveFile  = True
				If (nArguments > 2) Then
					strSaveFile = Right(strOption,(nArguments - 2))
				Else
					bInvalidArgument = True
				End If
			Case "-s"
				strWrittenServerSource = ""
				If (nArguments > 2) Then
					strWrittenServerSource = Right(strOption,(nArguments - 2))
				End If
			Case "-T"
					bUseDOTFile  = True
				If (nArguments > 2) Then
					strDOTFile = Right(strOption,(nArguments - 2))
				Else
					bInvalidArgument = True
				End If
			Case "-x"
				strSYDIServerSource = ""
				If (nArguments > 2) Then
					strSYDIServerSource = Right(strOption,(nArguments - 2))
				End If
			Case "-h"
				bDisplayHelp = True
			Case Else
				bInvalidArgument = True
		End Select
	End If
End Sub ' SetOptions

Sub UpdateDocumentInformation
	If (bUseSpecificTable) Then
		For i = 1 To CInt(oWord.ActiveDocument.Tables.Count)
			oWord.ActiveDocument.Tables(i).Style = strWordTable
		Next
	
	End If
	
	If Not (bUseDOTFile) Then
		' Adding header and footer
		If oWord.ActiveWindow.View.SplitSpecial = wdPaneNone Then
			oWord.ActiveWindow.ActivePane.View.Type = wdPrintView
		Else
			oWord.ActiveWindow.View.Type = wdPrintView
		End If
		oWord.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
		oWord.Selection.TypeText strSystem_Name
		oWord.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
		oWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
		oWord.Selection.TypeText "("
		oWord.Selection.Fields.Add oWord.Selection.Range, wdFieldPage
		oWord.Selection.TypeText "/"
		oWord.Selection.Fields.Add oWord.Selection.Range, wdFieldNumPages
		oWord.Selection.TypeText ")"
		oWord.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
	End If
	
	' Update table of contents
	ReportProgress " Updating Tables Of Contents"	
	oWord.ActiveDocument.TablesOfContents.Item(1).Update
	oWord.ActiveDocument.BuiltInDocumentProperties(wdPropertyComments).Value = "Generated by SYDI-Server " & strScriptVersion & " (http://sydiproject.com)"

	If (bSaveFile) Then 
		ReportProgress " Saving Document"	
		oWord.ActiveDocument.SaveAs strSaveFile
		If (Err <> 0) Then
		    ReportProgress Err.Number & " -- " &  Err.Description & " (" & strSaveFile & ")"
		    ReportProgress " Would not save to " & strSaveFile
		    ReportProgress " Did you specify a path?"
		    Err.Clear
		    Exit Sub
		End If
	End If
	
	If (bShowWord = False And bSaveFile = True) Then
		ReportProgress " Document Saved"
		oWord.Application.Quit
		Set oWord = Nothing
	Else
		oWord.Application.Visible = True
	End If
	

End Sub ' UpdateDocumentInformation

Sub WriteBold(strText)
	oWord.Selection.Font.Bold = True
	oWord.Selection.TypeText Cstr(strText) & vbCrLf
	oWord.Selection.Font.Bold = False
End Sub ' WriteBold

Sub WriteHeader(nHeaderLevel,strHeaderText)
	Select Case nHeaderLevel
		Case 1
			oWord.Selection.Style = wdStyleHeading1
			ReportProgress " Writing Section: " & strHeaderText
		Case 2
			oWord.Selection.Style = wdStyleHeading2
			ReportProgress "  Writing Sub-Section: " & strHeaderText
		Case 3
			oWord.Selection.Style = wdStyleHeading3
		Case 4
			oWord.Selection.Style = wdStyleHeading4
	End Select
	oWord.Selection.Range.ListFormat.ApplyListTemplate oListTemplate, True
	oWord.Selection.TypeText strHeaderText & vbCrLf
	oWord.Selection.Style = wdStyleBodyText
End Sub ' WriteHeader

Sub Write2LevelList (strRecord)
	Dim strTemp
	strTemp = ""
	objDbrRecordsets.Filter = " name='" & strRecord & "'"
	Do Until objDbrRecordsets.EOF
		If (objDbrRecordsets.Fields.Item("Column2") = "") Then
				If (strTemp <> "") Then
					oWord.Selection.TypeText vbCrLf
				End If
				oWord.Selection.Font.Bold = True
				oWord.Selection.TypeText objDbrRecordsets.Fields.Item("Column1") & VbCrLf
				oWord.Selection.Font.Bold = False
				strTemp = ""
		ElseIf (strTemp = "") Then
			If (objDbrRecordsets.Fields.Item("Column2") = "") Then
				oWord.Selection.Font.Bold = True
				oWord.Selection.TypeText objDbrRecordsets.Fields.Item("Column1") & VbCrLf
				oWord.Selection.Font.Bold = False
			Else
				strTemp = objDbrRecordsets.Fields.Item("Column2")
				oWord.Selection.Font.Bold = True
				oWord.Selection.TypeText objDbrRecordsets.Fields.Item("Column2") & ": "
				oWord.Selection.Font.Bold = False
				oWord.Selection.TypeText Cstr(objDbrRecordsets.Fields.Item("Column1"))
			End If
		ElseIf (strTemp = objDbrRecordsets.Fields.Item("Column2")) Then
			oWord.Selection.TypeText ", " & objDbrRecordsets.Fields.Item("Column1")
		ElseIf (strTemp <> objDbrRecordsets.Fields.Item("Column2")) Then
			If (objDbrRecordsets.Fields.Item("Column2") = "") Then
				oWord.Selection.Font.Bold = True
				oWord.Selection.TypeText objDbrRecordsets.Fields.Item("Column1") & VbCrLf
				oWord.Selection.Font.Bold = False
			Else
				oWord.Selection.Font.Bold = True
				oWord.Selection.TypeText VbCrLf & objDbrRecordsets.Fields.Item("Column2") & ": "
				oWord.Selection.Font.Bold = False
				oWord.Selection.TypeText Cstr(objDbrRecordsets.Fields.Item("Column1"))
			End If
			strTemp = objDbrRecordsets.Fields.Item("Column2")
		End If
		objDbrRecordsets.MoveNext
	Loop
	oWord.Selection.TypeText VbCrLf
End Sub 'Write2LevelList

Sub WriteIPConfiguration (strTranslation)
	Dim arrTranslation
	i = 1
	arrTranslation = Split(strTranslation,",")
	objDbrRecordsets.Filter = " name='dbrIPConfiguration'"
	Do Until objDbrRecordsets.EOF
		oWord.Selection.Font.Bold = True
		oWord.Selection.TypeText arrTranslation(0) & " " & i & vbCrLf
		oWord.Selection.Font.Bold = False
		oWord.Selection.TypeText arrTranslation(1) & ": " & CStr(objDbrRecordsets.Fields.Item("Column1")) & VbCrLf
		oWord.Selection.TypeText arrTranslation(2) & ": " & CStr(objDbrRecordsets.Fields.Item("Column2")) & VbCrLf
		oWord.Selection.TypeText arrTranslation(3) & ": " & CStr(objDbrRecordsets.Fields.Item("Column3")) & VbCrLf
		If (objDbrRecordsets.Fields.Item("Column4") <> "") Then
			oWord.Selection.TypeText arrTranslation(4) & ": " & CStr(objDbrRecordsets.Fields.Item("Column4")) & VbCrLf
		End If
		If (objDbrRecordsets.Fields.Item("Column5") <> "") Then
			oWord.Selection.TypeText arrTranslation(5) & ": " & CStr(objDbrRecordsets.Fields.Item("Column5")) & VbCrLf
		End If
		If (objDbrRecordsets.Fields.Item("Column6") <> "") Then
			oWord.Selection.TypeText arrTranslation(6) & ": " & CStr(objDbrRecordsets.Fields.Item("Column6")) & VbCrLf
		End If
		If (objDbrRecordsets.Fields.Item("Column7") <> "") Then
			oWord.Selection.TypeText arrTranslation(7) & ": " & CStr(objDbrRecordsets.Fields.Item("Column7")) & VbCrLf
		End If
		If (objDbrRecordsets.Fields.Item("Column8") <> "") Then
			oWord.Selection.TypeText arrTranslation(8) & ": " & CStr(objDbrRecordsets.Fields.Item("Column8")) & VbCrLf
		End If
		If (objDbrRecordsets.Fields.Item("Column9") <> "") Then
			oWord.Selection.TypeText arrTranslation(9) & ": " & CStr(objDbrRecordsets.Fields.Item("Column9")) & VbCrLf
		End If
		i = i + 1
		objDbrRecordsets.MoveNext
	Loop
End Sub 'WriteIPConfiguration

Sub WriteList (strRecord)
	objDbrRecordsets.Filter = " name='" & strRecord & "'"
	Do Until objDbrRecordsets.EOF
		oWord.Selection.TypeText Cstr(objDbrRecordsets.Fields.Item("Column1")) & vbCrLf
		objDbrRecordsets.MoveNext
	Loop
	oWord.Selection.TypeText VbCrLf	
End Sub 'WriteList

Sub WriteLoop(strRecord,nColumns,strTranslation)
	Dim arrTranslation
	arrTranslation = Split(strTranslation,",")
	objDbrRecordsets.Filter = " name='" & strRecord & "'"
	
	Do Until objDbrRecordsets.EOF
		oWord.Selection.Font.Bold = True
		oWord.Selection.TypeText CStr(objDbrRecordsets.Fields.Item("Column1")) & vbCrLf
		oWord.Selection.Font.Bold = False
		oWord.Selection.TypeText arrTranslation(1) & ": " & CStr(objDbrRecordsets.Fields.Item("Column2")) & vbCrLf
		oWord.Selection.TypeText arrTranslation(2) & ": " & Cstr(objDbrRecordsets.Fields.Item("Column3")) & vbCrLf
		If (nColumns > 3) Then
			oWord.Selection.TypeText arrTranslation(3) & ": " & Cstr(objDbrRecordsets.Fields.Item("Column4")) & vbCrLf
		End If
		objDbrRecordsets.MoveNext
	Loop
	oWord.Selection.TypeText VbCrLf
End Sub ' WriteLoop


Sub WriteSpace 
	oWord.Selection.TypeText VbCrLf
End Sub ' WriteSpace

Sub WriteStorage (strTranslation)
	Dim strTemp, arrTranslation
	arrTranslation = Split(strTranslation,",")
	objDbrRecordsets.Filter = " name='dbrStorage'"
	Do Until objDbrRecordsets.EOF
		If (strTemp <> objDbrRecordsets.Fields.Item("Column1")) Then
			oWord.Selection.Font.Bold = True
			oWord.Selection.TypeText CStr(objDbrRecordsets.Fields.Item("Column1")) & " - " &  CStr(objDbrRecordsets.Fields.Item("Column2")) & VbCrLf
			oWord.Selection.Font.Bold = False
			oWord.Selection.TypeText arrTranslation(0) & ": " & CStr(objDbrRecordsets.Fields.Item("Column3")) & VbCrLf
			oWord.Selection.TypeText arrTranslation(1) & ": " & CStr(objDbrRecordsets.Fields.Item("Column4")) & VbCrLf
			oWord.Selection.TypeText CStr(objDbrRecordsets.Fields.Item("Column5")) & " " & CStr(objDbrRecordsets.Fields.Item("Column6")) & _
				" Gb (" & CStr(objDbrRecordsets.Fields.Item("Column7")) & " Gb " & arrTranslation(2) & ") " &  CStr(objDbrRecordsets.Fields.Item("Column8")) & vbCrLf
		Else
			oWord.Selection.TypeText CStr(objDbrRecordsets.Fields.Item("Column5")) & " " & CStr(objDbrRecordsets.Fields.Item("Column6")) & _
				" Gb (" & CStr(objDbrRecordsets.Fields.Item("Column7")) & " Gb " & arrTranslation(2) & ") " &  CStr(objDbrRecordsets.Fields.Item("Column8")) & vbCrLf
		End If
		strTemp = objDbrRecordsets.Fields.Item("Column1")
		objDbrRecordsets.MoveNext
	Loop
	oWord.Selection.TypeText VbCrLf
End Sub 'WriteStorage

Sub WriteString(strVariable,strTranslation)
	objDbrVariables.Filter = " name='" & strVariable & "'"
	Do Until objDbrVariables.Eof
		'oWord.Selection.Font.Bold = True
		oWord.Selection.TypeText strTranslation & ": " 
		'oWord.Selection.Font.Bold = False
		oWord.Selection.TypeText Cstr(objDbrVariables("Value")) & vbCrLf						
		objDbrVariables.MoveNext
	Loop
End Sub ' WriteString

Sub WriteStringIIS(strServer,strVariable,strTranslation)
	objDbrIISVariables.Filter = " server='" & strServer & "' And name='" & strVariable & "' "
	Do Until objDbrIISVariables.Eof
		'oWord.Selection.Font.Bold = True
		oWord.Selection.TypeText strTranslation & ": " 
		'oWord.Selection.Font.Bold = False
		oWord.Selection.TypeText Cstr(objDbrIISVariables("Value")) & vbCrLf						
		objDbrIISVariables.MoveNext
	Loop
End Sub ' WriteStringIIS

Sub WriteTable(strRecord,nColumns,strTranslation)
	Dim arrTranslation
	arrTranslation = Split(strTranslation,",")
	objDbrRecordsets.Filter = " name='" & strRecord & "'"
	oWord.ActiveDocument.Tables.Add oWord.Selection.Range, objDbrRecordsets.Recordcount + 1, nColumns
	If Not (bUseSpecificTable) Then
		oWord.Selection.Font.Bold = True
	End If
	oWord.Selection.TypeText arrTranslation(0) : oWord.Selection.MoveRight
	If Not (bUseSpecificTable) Then
		oWord.Selection.Font.Bold = True
	End If
	oWord.Selection.TypeText arrTranslation(1) : oWord.Selection.MoveRight
	If (nColumns > 2) Then
		If Not (bUseSpecificTable) Then
			oWord.Selection.Font.Bold = True
		End If
		oWord.Selection.TypeText arrTranslation(2) : oWord.Selection.MoveRight
	End If
	If (nColumns > 3) Then
		If Not (bUseSpecificTable) Then
			oWord.Selection.Font.Bold = True
		End If
		oWord.Selection.TypeText arrTranslation(3) : oWord.Selection.MoveRight
	End If
	Do Until objDbrRecordsets.EOF
		oWord.Selection.TypeText CStr(objDbrRecordsets.Fields.Item("Column1")) : oWord.Selection.MoveRight
		oWord.Selection.TypeText Cstr(objDbrRecordsets.Fields.Item("Column2")) : oWord.Selection.MoveRight
		If (nColumns > 2) Then
			oWord.Selection.TypeText Cstr(objDbrRecordsets.Fields.Item("Column3")) : oWord.Selection.MoveRight
		End If
		If (nColumns > 3) Then
			oWord.Selection.TypeText Cstr(objDbrRecordsets.Fields.Item("Column4")) : oWord.Selection.MoveRight
		End If
		objDbrRecordsets.MoveNext
	Loop
	oWord.Selection.TypeText VbCrLf
End Sub ' WriteTable

Sub WriteTableIIS(strServer,strRecord,nColumns,strTranslation)
	Dim arrTranslation
	arrTranslation = Split(strTranslation,",")

	objDbrIISRecordsets.Filter = " server='" & strServer & "' And Name='" & strRecord & "' "
	oWord.ActiveDocument.Tables.Add oWord.Selection.Range, objDbrIISRecordsets.Recordcount + 1, nColumns
	If Not (bUseSpecificTable) Then
		oWord.Selection.Font.Bold = True
	End If
	oWord.Selection.TypeText arrTranslation(0) : oWord.Selection.MoveRight
	If Not (bUseSpecificTable) Then
		oWord.Selection.Font.Bold = True
	End If
	oWord.Selection.TypeText arrTranslation(1) : oWord.Selection.MoveRight
	If (nColumns > 2) Then
		If Not (bUseSpecificTable) Then
			oWord.Selection.Font.Bold = True
		End If
		oWord.Selection.TypeText arrTranslation(2) : oWord.Selection.MoveRight
	End If
	If (nColumns > 3) Then
		If Not (bUseSpecificTable) Then
			oWord.Selection.Font.Bold = True
		End If
		oWord.Selection.TypeText arrTranslation(3) : oWord.Selection.MoveRight
	End If
	Do Until objDbrIISRecordsets.EOF
		oWord.Selection.TypeText CStr(objDbrIISRecordsets.Fields.Item("Column1")) : oWord.Selection.MoveRight
		oWord.Selection.TypeText Cstr(objDbrIISRecordsets.Fields.Item("Column2")) : oWord.Selection.MoveRight
		If (nColumns > 2) Then
			oWord.Selection.TypeText Cstr(objDbrIISRecordsets.Fields.Item("Column3")) : oWord.Selection.MoveRight
		End If
		If (nColumns > 3) Then
			oWord.Selection.TypeText Cstr(objDbrIISRecordsets.Fields.Item("Column4")) : oWord.Selection.MoveRight
		End If
		objDbrIISRecordsets.MoveNext
	Loop
	oWord.Selection.TypeText VbCrLf
End Sub ' WriteTableIIS

Sub WriteText(strText)
	oWord.Selection.TypeText Cstr(strText) & VbCrLf
End Sub ' WriteText

Sub WriteWebServer (strTranslation)
	Dim arrTranslation
	Dim strTemp
	arrTranslation = Split(strTranslation,",")
	objDbrRecordsets.Filter = " name='dbrIISWebServer'"
	Do Until objDbrRecordsets.EOF
		WriteHeader 3, CStr(objDbrRecordsets.Fields.Item("Column2"))
		strTemp = arrTranslation(1) & "," & arrTranslation(2) & "," & arrTranslation(3) 
		WriteStringIIS CStr(objDbrRecordsets.Fields.Item("Column1")), "strHomeDirectory", arrTranslation(0)
		WriteTableIIS CStr(objDbrRecordsets.Fields.Item("Column1")), "dbrServerBindings", 3, strTemp
		objDbrRecordsets.MoveNext
	Loop	
End Sub ' WriteWebServer

'==========================================================
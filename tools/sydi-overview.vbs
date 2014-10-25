Option Explicit
'==========================================================
' LANG : VBScript
' NAME : sydi-overview.vbs
' AUTHOR : Patrick Ogenstad (patrick.ogenstad@netsafe.se)
' VERSION : 0.3
' DATE : 2008-12-9
' Description : Creates an Excel overview of SYDI-Server XML files
'
' UPDATES : http://sydiproject.com/ (Part of SYDI-Server)
'
' Running the script:
' cscript.exe sydi-overview.vbs -x[directory to xml files]
' For Options: cscript.exe sydi-overview.vbs -h
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
strScriptVersion = "0.3"

' Variables
Dim strXMLDirectory
Dim bInvalidArgument, bDisplayHelp
Dim intXMLFileCount, intXMLTotalFiles

' Database Records
Dim objDbrXMLFiles, objDbrComputers, objDbrComputerServices, objDbrOSDistribution, objDbrRegistryPrograms, objDbrWMIPrograms
Dim objDbrProcesses, objDbrServiceList, objDbrSubnet, objDbrSubnetMember, objDbrProductKeys

' Settings
Dim bServiceComparison, bIPSubnets


' Constants
Const adInteger = 3
Const adNumeric = 131
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
	WScript.Echo "Tab Options"
	WScript.Echo " -t	- Tabs to include (Default: -tis)"
 	WScript.Echo "   i	- IP Subnets"
 	WScript.Echo "   s	- Service Comparison"
 	WScript.Echo VbCrLf
 	WScript.Echo " -h	- Display help"
 	WScript.Echo VbCrLf
End Sub ' DisplayHelp

Function GetBinary(iDecimal) 
	Dim strBinary
	strBinary = ""

	Do Until iDecimal = 0
		If iDecimal Mod 2 = 0 Then
			strBinary = "0" & strBinary
		Else
			strBinary = "1" & strBinary
		End If
		iDecimal = iDecimal \ 2
	Loop
	GetBinary = strBinary
End Function ' GetBinary


Function GetBinaryAddress(strAddress) 
	Dim strBinaryValue, strBinaryAddress
	Dim arrAddress, iOctet
	arrAddress = Split(strAddress,".")
	
	For iOctet = 0 To 3
		strBinaryValue = ""
		Do Until arrAddress(iOctet) = 0
			If arrAddress(iOctet) Mod 2 = 0 Then
				strBinaryValue = "0" & strBinaryValue
			Else
				strBinaryValue = "1" & strBinaryValue
			End If
			arrAddress(iOctet) = arrAddress(iOctet) \ 2
		Loop
		Select Case Len(strBinaryValue)
			Case 0
				strBinaryValue = "00000000"
			Case 1
				strBinaryValue = "0000000" & strBinaryValue
			Case 2
				strBinaryValue = "000000" & strBinaryValue
			Case 3
				strBinaryValue = "00000" & strBinaryValue
			Case 4
				strBinaryValue = "0000" & strBinaryValue
			Case 5
				strBinaryValue = "000" & strBinaryValue
			Case 6
				strBinaryValue = "00" & strBinaryValue
			Case 7
				strBinaryValue = "0" & strBinaryValue
		End Select
		If (Len(strBinaryAddress) > 0) Then
			strBinaryAddress = strBinaryAddress & "." & strBinaryValue
		Else
			strBinaryAddress = strBinaryValue
		End If
		
	Next
	GetBinaryAddress = strBinaryAddress
End Function ' GetBinaryAddress


Function GetDecimal(strBinary)
  Dim iDecimal, xPos
  iDecimal = 0
  For xPos = 8 to 1 step -1
	If (Mid(strBinary, xPos, 1) = "1") Then
		iDecimal = iDecimal + (2 ^ (Len(strBinary)-xPos))
	End If
  Next
  GetDecimal = iDecimal
End Function ' GetDecimal


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
	bIPSubnets = True
	Set objArgs = WScript.Arguments
	If (objArgs.Count > 0) Then
		For nArgs = 0 To objArgs.Count - 1
			SetOptions objArgs(nArgs)
		Next
	Else
		bDisplayHelp = True
	End If
End Sub ' GetOptions

Function GetServiceStatus (strComputerName, strServiceName)
	Dim strReturn
	strComputerName = Replace(strComputerName,"'","")
	strServiceName = Replace(strServiceName,"'","")
	objDbrComputerServices.Filter = " Computer='" & strComputerName & "' and Service='" & strServiceName & "' "
	If (objDbrComputerServices.Bof) Then
		strReturn = "N/A"
	Else
		strReturn = objDbrComputerServices("Status")
	End If
	GetServiceStatus = strReturn
End Function ' GetServiceStatus

Function GetSubnetMaskLength(strBinaryMask)
	Dim strSubnetMask, xPos, nMaskLength
	strSubnetMask = Replace(strBinaryMask,".","")
	For xPos = 1 To 32
		If (Mid(strSubnetMask,xPos,1) = 0) Then
			Exit For
		End If
	Next
	nMaskLength = xPos - 1
	GetSubnetMaskLength = nMaskLength
End Function ' GetSubnetMaskLength


Sub Init()
	Set objDbrXMLFiles = CreateObject("ADOR.Recordset")
	objDbrXMLFiles.Fields.Append "XMLFile", adVarChar, MaxCharacters
	objDbrXMLFiles.Open
	
	Set objDbrComputers = CreateObject("ADOR.Recordset")
	objDbrComputers.Fields.Append "Name", adVarChar, MaxCharacters
	objDbrComputers.Fields.Append "OS", adVarChar, MaxCharacters
	objDbrComputers.Fields.Append "ServicePack", adVarChar, MaxCharacters
	objDbrComputers.Fields.Append "CPU", adVarChar, MaxCharacters
	objDbrComputers.Fields.Append "Memory", adVarChar, MaxCharacters
	objDbrComputers.Fields.Append "IPAddress", adVarChar, MaxCharacters
	objDbrComputers.Fields.Append "LastUser", adVarChar, MaxCharacters
	objDbrComputers.Fields.Append "ServiceTag", adVarChar, MaxCharacters
	objDbrComputers.Fields.Append "ScanTime", adVarChar, MaxCharacters
	objDbrComputers.Fields.Append "Chassis", adVarChar, MaxCharacters
	objDbrComputers.Fields.Append "SMBIOS", adVarChar, MaxCharacters
	objDbrComputers.Fields.Append "Language", adVarChar, MaxCharacters
	objDbrComputers.Fields.Append "Manufacturer", adVarChar, MaxCharacters
	objDbrComputers.Fields.Append "ProductName", adVarChar, MaxCharacters
	objDbrComputers.Open

	Set objDbrComputerServices = CreateObject("ADOR.Recordset")
	objDbrComputerServices.Fields.Append "Computer", adVarChar, MaxCharacters
	objDbrComputerServices.Fields.Append "Service", adVarChar, MaxCharacters
	objDbrComputerServices.Fields.Append "Status", adVarChar, MaxCharacters
	objDbrComputerServices.Open
	
	Set objDbrOSDistribution = CreateObject("ADOR.Recordset")
	objDbrOSDistribution.Fields.Append "Name", adVarChar, MaxCharacters
	objDbrOSDistribution.Fields.Append "Count", adVarChar, MaxCharacters
	objDbrOSDistribution.Open

	Set objDbrRegistryPrograms = CreateObject("ADOR.Recordset")
	objDbrRegistryPrograms.Fields.Append "Name", adVarChar, MaxCharacters
	objDbrRegistryPrograms.Fields.Append "Count", adVarChar, MaxCharacters
	objDbrRegistryPrograms.Open

	Set objDbrWMIPrograms = CreateObject("ADOR.Recordset")
	objDbrWMIPrograms.Fields.Append "Name", adVarChar, MaxCharacters
	objDbrWMIPrograms.Fields.Append "Version", adVarChar, MaxCharacters
	objDbrWMIPrograms.Fields.Append "Count", adVarChar, MaxCharacters
	objDbrWMIPrograms.Open

	Set objDbrProcesses = CreateObject("ADOR.Recordset")
	objDbrProcesses.Fields.Append "Name", adVarChar, MaxCharacters
	objDbrProcesses.Fields.Append "Executable", adVarChar, MaxCharacters
	objDbrProcesses.Fields.Append "Count", adVarChar, MaxCharacters
	objDbrProcesses.Open

	Set objDbrProductKeys = CreateObject("ADOR.Recordset")
	objDbrProductKeys.Fields.Append "Name", adVarChar, MaxCharacters
	objDbrProductKeys.Fields.Append "Key", adVarChar, MaxCharacters
	objDbrProductKeys.Fields.Append "Count", adVarChar, MaxCharacters
	objDbrProductKeys.Open

	
	Set objDbrServiceList = CreateObject("ADOR.Recordset")
	objDbrServiceList.Fields.Append "Name", adVarChar, MaxCharacters
	objDbrServiceList.Fields.Append "Count", adVarChar, MaxCharacters
	objDbrServiceList.Open

	Set objDbrSubnet = CreateObject("ADOR.Recordset")
	objDbrSubnet.Fields.Append "Subnet", adVarChar, MaxCharacters
	objDbrSubnet.Fields.Append "Mask", adVarChar, MaxCharacters
	objDbrSubnet.Fields.Append "Length", adVarChar, MaxCharacters
	objDbrSubnet.Fields.Append "Members", adVarChar, MaxCharacters
	objDbrSubnet.Fields.Append "IPOctet1", adInteger, MaxCharacters
	objDbrSubnet.Fields.Append "IPOctet2", adInteger, MaxCharacters
	objDbrSubnet.Fields.Append "IPOctet3", adInteger, MaxCharacters
	objDbrSubnet.Fields.Append "IPOctet4", adInteger, MaxCharacters	
	objDbrSubnet.Fields.Append "Priority", adInteger, MaxCharacters
	objDbrSubnet.Open

	
	Set objDbrSubnetMember = CreateObject("ADOR.Recordset")
	objDbrSubnetMember.Fields.Append "Computer", adVarChar, MaxCharacters
	objDbrSubnetMember.Fields.Append "IP", adVarChar, MaxCharacters
	objDbrSubnetMember.Fields.Append "Subnet", adVarChar, MaxCharacters
	objDbrSubnetMember.Fields.Append "Mask", adVarChar, MaxCharacters
	objDbrSubnetMember.Fields.Append "Length", adVarChar, MaxCharacters
	objDbrSubnetMember.Fields.Append "IPOctet1", adInteger, MaxCharacters
	objDbrSubnetMember.Fields.Append "IPOctet2", adInteger, MaxCharacters
	objDbrSubnetMember.Fields.Append "IPOctet3", adInteger, MaxCharacters
	objDbrSubnetMember.Fields.Append "IPOctet4", adInteger, MaxCharacters	
	objDbrSubnetMember.Fields.Append "Priority", adInteger, MaxCharacters
	objDbrSubnetMember.Open	

End Sub 'Init






Sub ParseXMLfile(strXMLFile)
	Dim objXMLFile, colNodes
	Dim objNode, objChild
	Dim objProgram, objProcess, objService, objAdapter, objIP
	Set objXMLFile = CreateObject("Microsoft.XMLDOM")
	objXMLFile.async = False
	objXMLFile.load(strXMLDirectory & "\" & strXMLFile)
	Wscript.Echo "Parsing: " & strXMLDirectory & "\" & strXMLFile & " (" & intXMLFileCount & "/" & intXMLTotalFiles & ")"
	intXMLFileCount = intXMLFileCount + 1
	Set colNodes = objXMLFile.selectNodes("//computer")
	objDbrComputers.AddNew
	For Each objNode in colNodes
		For Each objChild in objNode.childNodes
			If (objChild.nodeName = "generated") Then
				objDbrComputers("scantime") = Scrub(objChild.getAttribute("scantime"))
			End If
			If (objChild.nodeName = "system") Then
				objDbrComputers("name") = Scrub(objChild.getAttribute("name"))
			End If
			If (objChild.nodeName = "operatingsystem") Then
				objDbrComputers("os") = Scrub(objChild.getAttribute("name"))
				OSDistribution(Scrub(objChild.getAttribute("name")))
				objDbrComputers("servicepack") = Scrub(objChild.getAttribute("servicepack"))
			End If
			If (objChild.nodeName = "machineinfo") Then
				objDbrComputers("manufacturer") = Scrub(objChild.getAttribute("manufacturer"))
				objDbrComputers("productname") = Scrub(objChild.getAttribute("productname"))
				objDbrComputers("servicetag") = Scrub(objChild.getAttribute("identifyingnumber"))
				objDbrComputers("chassis") = Scrub(objChild.getAttribute("chassis"))
			End If
			If (objChild.nodeName = "processor") Then
				objDbrComputers("cpu") = Scrub(objChild.getAttribute("speed"))
			End If
			If (objChild.nodeName = "memory") Then
				objDbrComputers("memory") = Scrub(objChild.getAttribute("totalsize"))
			End If
			If (objChild.nodeName = "bios") Then
				objDbrComputers("smbios") = Scrub(objChild.getAttribute("smbiosversion"))
			End If
			If (objChild.nodeName = "lastuser") Then
				objDbrComputers("lastuser") = Scrub(objChild.getAttribute("name"))
			End If
			If (objChild.nodeName = "osconfiguration") Then
				objDbrComputers("language") = Scrub(objChild.getAttribute("oslanguage"))
			End If
			If (objChild.nodeName = "installedapplications") Then
				For Each objProgram in objChild.childNodes
					If (objProgram.nodeName = "regapplication") Then
						RegistryPrograms Scrub(objProgram.getAttribute("productname"))
					End If
					If (objProgram.nodeName = "msiapplication") Then
						WMIPrograms Scrub(objProgram.getAttribute("productname")),Scrub(objProgram.getAttribute("version"))
					End If
					If (objProgram.nodeName = "productkey") Then
						ProductKeys Scrub(objProgram.getAttribute("productname")),Scrub(objProgram.getAttribute("productkey"))
					End If
				Next
			End If


			
			'objDbrSubnet, objDbrSubnetMember
			If (objChild.nodeName = "network") Then
				For Each objAdapter in objChild.childNodes
					If (objAdapter.nodeName = "adapter") Then
						For Each objIP in objAdapter.childNodes
							If (objIP.nodeName = "ip") Then
								If (Scrub(objIP.getAttribute("address")) <> "0.0.0.0" AND (Len(objIP.getAttribute("address")) < 17)) Then
									objDbrComputers("ipaddress") = Scrub(objIP.getAttribute("address"))
									PopulateSubnets objIP.getAttribute("address"),objIP.getAttribute("subnetmask"), objDbrComputers("name") 
								End If
							End If
						Next
					End If
				Next
			End If

			If (objChild.nodeName = "processes") Then
				For Each objProcess in objChild.childNodes
					Processes Scrub(objProcess.getAttribute("caption")),Scrub(objProcess.getAttribute("executablepath"))
				Next
			End If			
			' Services
			If (objChild.nodeName = "services") Then
				For Each objService in objChild.childNodes
					Services objDbrComputers("name"),Scrub(objService.getAttribute("name")),Scrub(objService.getAttribute("startmode")),Scrub(objService.getAttribute("started"))
				Next
			End If

			
			
		Next
	Next 
	objDbrComputers.Update
	SetSubnetPriority
	SetHostPriority
End Sub ' ParseXMLfile


Sub PopulateSubnets(strIPAddress,strNetMask,strComputer)
	Dim strBinaryIP, strBinaryMask, i
	Dim nMaskLength, strIPSubnet, arrIPSubnet, arrIPAddress
	
	strBinaryIP = GetBinaryAddress(strIPAddress)
	strBinaryMask = GetBinaryAddress(strNetMask)
	nMaskLength =  GetSubnetMaskLength(strBinaryMask)
	
	For i = 1 To 35
		If (Mid(strBinaryIP,i,1) = ".") Then
			strIPSubnet = strIPSubnet & "."
		Else
			If (Cbool(Mid(strBinaryIP,i,1)) And Cbool(Mid(strBinaryMask,i,1))) Then
				strIPSubnet = strIPSubnet & "1"
			Else
				strIPSubnet = strIPSubnet & "0"
			End If	
		End If
		
	Next
	
	arrIPSubnet = Split(strIPSubnet,".")
	strIPSubnet =  GetDecimal(arrIPSubnet(0)) & "." & GetDecimal(arrIPSubnet(1)) & "." & GetDecimal(arrIPSubnet(2)) & "." & GetDecimal(arrIPSubnet(3))


	
	objDbrSubnet.Filter = " Subnet='" & strIPSubnet & "' And Length='" & nMaskLength & "'"
	If (objDbrSubnet.Bof) Then
		objDbrSubnet.AddNew
		objDbrSubnet("Subnet") = strIPSubnet
		objDbrSubnet("Mask") = strNetMask
		objDbrSubnet("Length") = nMaskLength
		objDbrSubnet("Members") = 1
		objDbrSubnet("IPOctet1") = arrIPSubnet(0)
		objDbrSubnet("IPOctet2") = arrIPSubnet(1)
		objDbrSubnet("IPOctet3") = arrIPSubnet(2)
		objDbrSubnet("IPOctet4") = arrIPSubnet(3)
		objDbrSubnet.Update
	Else
		objDbrSubnet("Members") = objDbrSubnet("Members") + 1
		objDbrSubnet.Update
	End If

	arrIPAddress = Split(strIPAddress,".")
	objDbrSubnetMember.AddNew
	objDbrSubnetMember("Computer") = strComputer
	objDbrSubnetMember("IP") = strIPAddress
	objDbrSubnetMember("Subnet") = strIPSubnet
	objDbrSubnetMember("Mask") = strNetMask
	objDbrSubnetMember("Length") = nMaskLength
	objDbrSubnetMember("IPOctet1") = arrIPAddress(0)
	objDbrSubnetMember("IPOctet2") = arrIPAddress(1)
	objDbrSubnetMember("IPOctet3") = arrIPAddress(2)
	objDbrSubnetMember("IPOctet4") = arrIPAddress(3)
	
	objDbrSubnetMember.Update
	

End Sub ' PopulateSubnets

Sub OSDistribution(strOperatingSystem)

	objDbrOSDistribution.Filter = " Name='" & strOperatingSystem & "'"
	If (objDbrOSDistribution.Bof) Then
		objDbrOSDistribution.AddNew
		objDbrOSDistribution("Name") = strOperatingSystem
		objDbrOSDistribution("Count") = 1
		objDbrOSDistribution.Update
	Else
		objDbrOSDistribution("Count") = objDbrOSDistribution("Count") + 1
		objDbrOSDistribution.Update
	End If
End Sub ' OSDistribution

Sub PopulateExcelfile()
	Dim objExcel, intLine, intXLine, intYLine, intIPHostLine
	Dim objWorksheet, objRange, objRange2, objRange3, objRange4
	Dim strServiceName
	
	Const xlPie = 5
	Const xlColumns = 2
	Const xlLocationAsNewSheet = 1
	Const xlNone = -4142
	Const xlThin = 2
	Const xlAscending = 1
	Const xlDescending = 2
	Const xlYes = 1
	
	Wscript.Echo "Opening Excel File"
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = True
	objExcel.Workbooks.Add
	objExcel.Cells(1, 1).Value = "Name"
	objExcel.Cells(1, 2).Value = "Os"
	objExcel.Cells(1, 3).Value = "Service Pack"
	objExcel.Cells(1, 4).Value = "CPU"
	objExcel.Cells(1, 5).Value = "Memory"
	objExcel.Cells(1, 6).Value = "IP Address"
	objExcel.Cells(1, 7).Value = "Last User"
	objExcel.Cells(1, 8).Value = "Manufacturer"
	objExcel.Cells(1, 9).Value = "Product Name"
	objExcel.Cells(1, 10).Value = "Chassis"
	objExcel.Cells(1, 11).Value = "Service Tag"
	objExcel.Cells(1, 12).Value = "SMBIOS"
	objExcel.Cells(1, 13).Value = "Language"
	objExcel.Cells(1, 14).Value = "Scan Time"

	objExcel.Range("A1:M1").Select
    objExcel.Selection.Font.Bold = True
	
	Wscript.Echo "Writing Basic Computer Information"
	intLine = 2
	If Not (objDbrComputers.Bof) Then
		objDbrComputers.MoveFirst
	End If
	Do Until objDbrComputers.EOF
		objExcel.Cells(intLine, 1).Value = objDbrComputers("Name")
		objExcel.Cells(intLine, 2).Value = objDbrComputers("OS")
		objExcel.Cells(intLine, 3).Value = objDbrComputers("ServicePack")
		objExcel.Cells(intLine, 4).Value = Clng(objDbrComputers("CPU"))
		objExcel.Cells(intLine, 5).Value = Clng(objDbrComputers("Memory"))
		objExcel.Cells(intLine, 6).Value = objDbrComputers("IPAddress")
		objExcel.Cells(intLine, 7).Value = objDbrComputers("LastUser")
		objExcel.Cells(intLine, 8).Value = objDbrComputers("Manufacturer")
		objExcel.Cells(intLine, 9).Value = objDbrComputers("ProductName")
		objExcel.Cells(intLine, 10).Value = objDbrComputers("Chassis")
		objExcel.Cells(intLine, 11).Value = objDbrComputers("ServiceTag")
		objExcel.Cells(intLine, 12).Value = objDbrComputers("SMBIOS")
		objExcel.Cells(intLine, 13).Value = objDbrComputers("Language")
		objExcel.Cells(intLine, 14).Value = objDbrComputers("ScanTime")
		objDbrComputers.MoveNext
		intLine = intLine + 1
	Loop
    objExcel.Cells.Select
    objExcel.Cells.EntireColumn.AutoFit	
    objExcel.ActiveWindow.SplitRow = 0.8
    objExcel.ActiveWindow.FreezePanes = True
	objExcel.Range("A1").Select
	objExcel.Sheets(1).Name = "Computers"
	objExcel.Sheets(2).Name = "WMI Programs"
	objExcel.Sheets(3).Name = "Registry Programs"
	objExcel.Sheets.Add ,objExcel.Sheets(3) ' Add a new sheet after the last one
	objExcel.Sheets(4).Name = "Processes"
	objExcel.Sheets.Add ,objExcel.Sheets(4) ' Add a new sheet after the last one
	objExcel.Sheets(5).Name = "OS Distribution Data"
	
	Wscript.Echo "Writing OS Distribution Data"
	' OS Distribution
	objExcel.Sheets("OS Distribution Data").Select
	objExcel.Cells(1, 1).Value = "Name"
	objExcel.Cells(1, 2).Value = "Count"
	objExcel.Range("A1:B1").Select
    objExcel.Selection.Font.Bold = True
	intLine = 2
	objDbrOSDistribution.Filter = " Count > 0 "
	If Not (objDbrOSDistribution.Bof) Then
		objDbrOSDistribution.MoveFirst
	End If
	Do Until objDbrOSDistribution.EOF
		objExcel.Cells(intLine, 1).Value = objDbrOSDistribution("Name")
		objExcel.Cells(intLine, 2).Value = Clng(objDbrOSDistribution("Count"))
		intLine = intLine + 1
		objDbrOSDistribution.MoveNext
	Loop
	intLine = intLine - 1
    objExcel.Cells.Select
    objExcel.Cells.EntireColumn.AutoFit	
	objExcel.Range("A1").Select
	Wscript.Echo "Creating OS Distribution Pie Chart"
    objExcel.Charts.Add
    objExcel.ActiveChart.ChartType = xlPie
    objExcel.ActiveChart.SetSourceData objExcel.Sheets("OS Distribution Data").Range( _
        "A2:B" & intLine), xlColumns
    objExcel.ActiveChart.Location xlLocationAsNewSheet, "Chart OS Distribution"
    With objExcel.ActiveChart
        .HasTitle = True
        .ChartTitle.Characters.Text = "OS Distribution"
	End With
	objExcel.ActiveChart.ApplyDataLabels ,,,,,,,True ' Add percentage
	objExcel.ActiveChart.PlotArea.Select
	With objExcel.Selection.Border
		.Weight = xlThin
		.LineStyle = xlNone
	End With
	objExcel.Selection.Interior.ColorIndex = xlNone
	objExcel.ActiveChart.ChartArea.Select

	Wscript.Echo "Writing Registry Programs Data"
	' Registry Programs
	objExcel.Sheets("Registry Programs").Select
	objExcel.Cells(1, 1).Value = "Name"
	objExcel.Cells(1, 2).Value = "Count"
	objExcel.Range("A1:B1").Select
    objExcel.Selection.Font.Bold = True
	intLine = 2
	objDbrRegistryPrograms.Filter = " Count > 0 "
	If Not (objDbrRegistryPrograms.Bof) Then
		objDbrRegistryPrograms.MoveFirst
	End If
	Do Until objDbrRegistryPrograms.EOF
		objExcel.Cells(intLine, 1).Value = objDbrRegistryPrograms("Name")
		objExcel.Cells(intLine, 2).Value = Clng(objDbrRegistryPrograms("Count"))
		intLine = intLine + 1
		objDbrRegistryPrograms.MoveNext
	Loop
    objExcel.Cells.Select
    objExcel.Cells.EntireColumn.AutoFit	
	objExcel.Range("A1").Select
	Set objWorksheet = objExcel.Worksheets("Registry Programs")
	Set objRange = objWorksheet.UsedRange
	Set objRange2 = objExcel.Range("A1")
	Set objRange3 = objExcel.Range("B1")
	objRange.Sort objRange2,xlAscending,objRange3,,xlDescending, ,,xlYes

	' WMI Programs
	Wscript.Echo "Writing Windows Installer Applications Data"
	objExcel.Sheets("WMI Programs").Select
	objExcel.Cells(1, 1).Value = "Name"
	objExcel.Cells(1, 2).Value = "Version"
	objExcel.Cells(1, 3).Value = "Count"
	objExcel.Range("A1:C1").Select
    objExcel.Selection.Font.Bold = True
	intLine = 2
	objDbrWMIPrograms.Filter = " Count > 0 "
	If Not (objDbrWMIPrograms.Bof) Then
		objDbrWMIPrograms.MoveFirst
	End If
	Do Until objDbrWMIPrograms.EOF
		objExcel.Cells(intLine, 1).Value = objDbrWMIPrograms("Name")
		objExcel.Cells(intLine, 2).Value = objDbrWMIPrograms("Version")
		objExcel.Cells(intLine, 3).Value = Clng(objDbrWMIPrograms("Count"))
		intLine = intLine + 1
		objDbrWMIPrograms.MoveNext
	Loop
    objExcel.Cells.Select
    objExcel.Cells.EntireColumn.AutoFit	
	objExcel.Range("A1").Select
	Set objWorksheet = objExcel.Worksheets("WMI Programs")
	Set objRange = objWorksheet.UsedRange
	Set objRange2 = objExcel.Range("A1")
	Set objRange3 = objExcel.Range("B1")
	Set objRange4 = objExcel.Range("C1")
	objRange.Sort objRange2,xlAscending,objRange3,,xlDescending, objRange4,xlDescending,xlYes

	' Processes
	Wscript.Echo "Writing Processes"
	objExcel.Sheets("Processes").Select
	objExcel.Cells(1, 1).Value = "Name"
	objExcel.Cells(1, 2).Value = "Executable"
	objExcel.Cells(1, 3).Value = "Count"
	objExcel.Range("A1:C1").Select
    objExcel.Selection.Font.Bold = True
	intLine = 2
	objDbrProcesses.Filter = " Count > 0 "
	If Not (objDbrProcesses.Bof) Then
		objDbrProcesses.MoveFirst
	End If
	Do Until objDbrProcesses.EOF
		objExcel.Cells(intLine, 1).Value = objDbrProcesses("Name")
		objExcel.Cells(intLine, 2).Value = objDbrProcesses("Executable")
		objExcel.Cells(intLine, 3).Value = Clng(objDbrProcesses("Count"))
		intLine = intLine + 1
		objDbrProcesses.MoveNext
	Loop
    objExcel.Cells.Select
    objExcel.Cells.EntireColumn.AutoFit	
	objExcel.Range("A1").Select
	Set objWorksheet = objExcel.Worksheets("Processes")
	Set objRange = objWorksheet.UsedRange
	Set objRange2 = objExcel.Range("A1")
	Set objRange3 = objExcel.Range("B1")
	Set objRange4 = objExcel.Range("C1")
	objRange.Sort objRange2,xlAscending,objRange3,,xlDescending, objRange4,xlDescending,xlYes
	
	' Product Keys
	objExcel.Sheets.Add ,objExcel.Sheets(objExcel.Sheets.Count) ' Add a new sheet after the last one
	objExcel.Sheets(objExcel.Sheets.Count).Name = "Product Keys"
	Wscript.Echo "Writing Product Keys"
	'objExcel.Sheets("Processes").Select
	objExcel.Cells(1, 1).Value = "Product"
	objExcel.Cells(1, 2).Value = "Key"
	objExcel.Cells(1, 3).Value = "Count"
	objExcel.Range("A1:C1").Select
    objExcel.Selection.Font.Bold = True
	intLine = 2
	objDbrProductKeys.Filter = " Count > 0 "
	If Not (objDbrProductKeys.Bof) Then
		objDbrProductKeys.Sort = "Name"
		objDbrProductKeys.MoveFirst
	End If
	Do Until objDbrProductKeys.EOF
		objExcel.Cells(intLine, 1).Value = objDbrProductKeys("Name")
		objExcel.Cells(intLine, 2).Value = objDbrProductKeys("Key")
		objExcel.Cells(intLine, 3).Value = Clng(objDbrProductKeys("Count"))
		intLine = intLine + 1
		objDbrProductKeys.MoveNext
	Loop
    objExcel.Cells.Select
    objExcel.Cells.EntireColumn.AutoFit	
	objExcel.Range("A1").Select
	'Set objWorksheet = objExcel.Worksheets("Product Keys")
	'Set objRange = objWorksheet.UsedRange
	'Set objRange2 = objExcel.Range("A1")
	'Set objRange3 = objExcel.Range("B1")
	'Set objRange4 = objExcel.Range("C1")
	'objRange.Sort objRange2,xlAscending,objRange3,,xlDescending, objRange4,xlDescending,xlYes
	
	If (bServiceComparison) Then
	'If (bServiceComparison) Then
		objExcel.Sheets.Add ,objExcel.Sheets(objExcel.Sheets.Count) ' Add a new sheet after the last one
		objExcel.Sheets(objExcel.Sheets.Count).Name = "Service Comparison"
	'End If


	' Services
		Wscript.Echo "Writing Services Comparison Data"
		objExcel.Sheets("Service Comparison").Select	
		objExcel.Cells(4, 1).Value = "Name"
		' Print Service Names
		intLine = 2
		objDbrServiceList.Filter = " Count > 0 "
		objDbrServiceList.Sort = "Name"
		If Not (objDbrServiceList.Bof) Then
			objDbrServiceList.MoveFirst
		End If
		Do Until objDbrServiceList.EOF
			objExcel.Cells(4, intLine).Value = objDbrServiceList("Name")
			objExcel.Selection.Font.Bold = True
			objDbrServiceList.MoveNext
			intLine = intLine + 1
		Loop

		' Print Computer Names
		intLine = 5
		If Not (objDbrComputers.Bof) Then
			objDbrComputers.MoveFirst
		End If
		Do Until objDbrComputers.EOF
			objExcel.Cells(intLine, 1).Value = objDbrComputers("Name")
			objDbrComputers.MoveNext
			intLine = intLine + 1
		Loop

		intXLine = 5
		intYLine = 2
		Do Until objExcel.Cells(4, intYLine).Value = ""
			strServiceName = Cstr(objExcel.Cells(4, intYLine).Value)
			Wscript.Echo " " & strServiceName & ": Comparing Computers"
			Do Until objExcel.Cells(intXLine, 1).Value = ""
				'objExcel.Cells(intXLine, intYLine).Value = "N/A " & intXLine & ":" & intYLine
				objExcel.Cells(intXLine, intYLine).Value = GetServiceStatus (objExcel.Cells(intXLine, 1).Value, strServiceName)
				'GetServiceStatus (strComputerName, strServiceName)
				intXLine = intXLine + 1
			Loop
			intXLine = 5
			intYLine = intYLine + 1
		Loop
'		objDbrComputerServices("Computer") = strServiceName
'		objDbrComputerServices("Service") = strServiceName
'		objDbrComputerServices("Status") = strStartMode & "-" & strStarted
'	
		
	    objExcel.Cells.Select
	    objExcel.Cells.EntireColumn.AutoFit	
		objExcel.ActiveWindow.SplitColumn = 1
	    objExcel.ActiveWindow.SplitRow = 4
	    objExcel.ActiveWindow.FreezePanes = True
	End If

	objExcel.Range("A1").Select


	If (bIPSubnets) Then
		objExcel.Sheets.Add ,objExcel.Sheets(objExcel.Sheets.Count) ' Add a new sheet after the last one
		objExcel.Sheets(objExcel.Sheets.Count).Name = "Subnets"
	
		intLine = 2
		objDbrSubnet.Filter = " Members > 0 "
		objDbrSubnet.Sort = "Subnet"
		objExcel.Cells(1, 1).Value = "Subnet Address"
		objExcel.Cells(1, 2).Value = "Subnet Mask"
		objExcel.Cells(1, 3).Value = "Subnet Length"
		objExcel.Cells(1, 4).Value = "Computers"

		If Not (objDbrSubnet.Bof) Then
			objDbrSubnet.MoveFirst
			objDbrSubnet.Sort = "Priority"
		End If
		Do Until objDbrSubnet.EOF
			objExcel.Cells(intLine, 1).Value = objDbrSubnet("Subnet")
			objExcel.Cells(intLine, 2).Value = objDbrSubnet("Mask")
			objExcel.Cells(intLine, 3).Value = objDbrSubnet("Length")
			objExcel.Cells(intLine, 4).Value = objDbrSubnet("Members")
			objExcel.Range("A1:D1").Select
			objExcel.Selection.Font.Bold = True
			objExcel.Sheets.Add ,objExcel.Sheets(objExcel.Sheets.Count) ' Add a new sheet after the last one
			objExcel.Sheets(objExcel.Sheets.Count).Name = objDbrSubnet("Subnet") & "(" & objDbrSubnet("Length") & ")"
			objDbrSubnetMember.Filter = " Subnet='" & objDbrSubnet("Subnet") & "' And Length='" & objDbrSubnet("Length") & "'"

			objDbrSubnetMember.Sort = "Priority"
			If Not (objDbrSubnetMember.Bof) Then
				objDbrSubnetMember.MoveFirst
				objExcel.Cells(1, 1).Value = "IP Address"
				objExcel.Cells(1, 2).Value = "Node"
				objExcel.Cells(1, 4).Value = "Subnets"
				objExcel.Cells(2, 1).Value = objDbrSubnet("Subnet")
				objExcel.Cells(2, 2).Value = "Network Address"
				objExcel.Range("A1:B1").Select
				objExcel.Selection.Font.Bold = True
				intIPHostLine = 3
			End If
			Do Until objDbrSubnetMember.EOF
				objExcel.Cells(intIPHostLine, 1).Value = objDbrSubnetMember("IP")
				objExcel.Cells(intIPHostLine, 2).Value = objDbrSubnetMember("Computer")
				intIPHostLine = intIPHostLine + 1
				objDbrSubnetMember.MoveNext
			Loop
			objExcel.ActiveWindow.SplitRow = 1
			objExcel.ActiveWindow.FreezePanes = True
			objExcel.Cells.Select
			objExcel.Cells.EntireColumn.AutoFit	
			objExcel.Sheets("Subnets").Select
			objDbrSubnet.MoveNext
			intLine = intLine + 1
		Loop
		objExcel.Cells.Select
	    objExcel.Cells.EntireColumn.AutoFit	
	End If
	
	'objDbrSubnetMember.AddNew
	'objDbrSubnetMember("Computer") = strComputer
	'objDbrSubnetMember("IP") = strIPAddress
	'objDbrSubnetMember("Subnet") = strIPSubnet
	'objDbrSubnetMember("Mask") = strNetMask
	'objDbrSubnetMember("Length") = nMaskLength
	'objDbrSubnetMember.Update

	
	
	objExcel.Sheets("Computers").Select
	
End Sub 'PopulateExcelfile

Sub Processes(strProcessName,strExecutable)
	strProcessName = Replace(strProcessName,"'","")
	strExecutable = Replace(strExecutable,"'","")
	objDbrProcesses.Filter = " Name='" & strProcessName & "' And Executable='" & strExecutable & "'"
	If (objDbrProcesses.Bof) Then
		objDbrProcesses.AddNew
		objDbrProcesses("Name") = strProcessName
		objDbrProcesses("Executable") = strExecutable
		objDbrProcesses("Count") = 1
		objDbrProcesses.Update
	Else
		objDbrProcesses("Count") = objDbrProcesses("Count") + 1
		objDbrProcesses.Update
	End If
End Sub ' Processes



Sub ProductKeys(strProductName,strProductKey)
	strProductName = Replace(strProductName,"'","")
	strProductKey = Replace(strProductKey,"'","")
	objDbrProductKeys.Filter = " Name='" & strProductName & "' And Key='" & strProductKey & "'"
	If (objDbrProductKeys.Bof) Then
		objDbrProductKeys.AddNew
		objDbrProductKeys("Name") = strProductName
		objDbrProductKeys("Key") = strProductKey
		objDbrProductKeys("Count") = 1
		objDbrProductKeys.Update
	Else
		objDbrProductKeys("Count") = objDbrProductKeys("Count") + 1
		objDbrProductKeys.Update
	End If
End Sub ' ProductKeys



Sub RegistryPrograms(strProgramName)
	strProgramName = Replace(strProgramName,"'","")
	objDbrRegistryPrograms.Filter = " Name='" & strProgramName & "'"
	If (objDbrRegistryPrograms.Bof) Then
		objDbrRegistryPrograms.AddNew
		objDbrRegistryPrograms("Name") = strProgramName
		objDbrRegistryPrograms("Count") = 1
		objDbrRegistryPrograms.Update
	Else
		objDbrRegistryPrograms("Count") = objDbrRegistryPrograms("Count") + 1
		objDbrRegistryPrograms.Update
	End If
End Sub ' RegistryPrograms


Function Scrub(strInput)
	If (IsNull(strInput)) Then
		strInput = ""
	End If
	Scrub = strInput
End Function ' Scrub



Sub Services (strComputerName, strServiceName, strStartMode, strStarted)
	strComputerName = Replace(strComputerName,"'","")
	strServiceName = Replace(strServiceName,"'","")
	strStartMode = Replace(strStartMode,"'","")
	strStarted = Replace(strStarted,"'","")
	objDbrServiceList.Filter = " Name='" & strServiceName & "'"
	If (objDbrServiceList.Bof) Then
		objDbrServiceList.AddNew
		objDbrServiceList("Name") = strServiceName
		objDbrServiceList("Count") = 1
		objDbrServiceList.Update
	Else
		objDbrServiceList("Count") = objDbrServiceList("Count") + 1
		objDbrServiceList.Update
	End If

	objDbrComputerServices.AddNew
	objDbrComputerServices("Computer") = strComputerName
	objDbrComputerServices("Service") = strServiceName
	objDbrComputerServices("Status") = strStartMode & "-" & strStarted
	objDbrComputerServices.Update
End Sub ' Services

Sub SetHostPriority()
	Dim nPriority
	nPriority = 1
	If Not (objDbrSubnetMember.Bof) Then
		objDbrSubnetMember.Sort = "IPOctet1,IPOctet2,IPOctet3,IPOctet4"
	End If
	
	Do Until objDbrSubnetMember.EOF
	
		objDbrSubnetMember("Priority") = nPriority
		nPriority = nPriority + 1
		objDbrSubnetMember.MoveNext
	Loop
End Sub 'SetHostPriority()



Sub SetOptions(strOption)
	Dim strFlag, strParameter
	Dim nArguments, i
	nArguments = Len(strOption)
	If (nArguments < 2) Then
		bInvalidArgument = True
	Else
		strFlag = Left(strOption,2)
		Select Case strFlag
			Case "-t"
				bServiceComparison = False
				bIPSubnets = False
				If (nArguments > 2) Then
					For i = 3 To nArguments
						strParameter = Mid(strOption,i,1)
						Select Case strParameter
							Case "i"
								bIPSubnets = True
							Case "s"
								bServiceComparison = True
							Case Else
								bInvalidArgument = True
						End Select
					Next
				End If
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

Sub SetSubnetPriority()
	Dim nPriority
	nPriority = 1
	If Not (objDbrSubnet.Bof) Then
		objDbrSubnet.Sort = "IPOctet1,IPOctet2,IPOctet3,IPOctet4"
	End If
	
	Do Until objDbrSubnet.EOF
		objDbrSubnet("Priority") = nPriority
		nPriority = nPriority + 1
		objDbrSubnet.MoveNext
	Loop
End Sub 'SetSubnetPriority()


Sub WMIPrograms(strProgramName,strVersion)
	strProgramName = Replace(strProgramName,"'","")
	strVersion = Replace(strVersion,"'","")
	objDbrWMIPrograms.Filter = " Name='" & strProgramName & "' And Version='" & strVersion & "'"
	If (objDbrWMIPrograms.Bof) Then
		objDbrWMIPrograms.AddNew
		objDbrWMIPrograms("Name") = strProgramName
		objDbrWMIPrograms("Version") = strVersion
		objDbrWMIPrograms("Count") = 1
		objDbrWMIPrograms.Update
	Else
		objDbrWMIPrograms("Count") = objDbrWMIPrograms("Count") + 1
		objDbrWMIPrograms.Update
	End If
End Sub ' RegistryPrograms


'==========================================================
<?xml version='1.0' encoding="UTF-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
         
                version="1.0"><xsl:output method="html"/>

<xsl:include href="sydi-html-styles.xsl"/>
<xsl:template match="/">


<xsl:call-template name="general-server" />
<html>
<head>
	<title>Documentation for <xsl:value-of  select = "computer/system/@name" /></title>
</head>
<body>

<h1><xsl:value-of  select = "computer/system/@name" /></h1>

<b>NetBIOS: </b><xsl:value-of select = "computer/system/@name" /><br />
<xsl:if test="computer/fqdn">
	<b>FQDN: </b><xsl:value-of  select = "computer/fqdn/@name" /><br />
</xsl:if>
<b>OS: </b><xsl:value-of select = "computer/operatingsystem/@name" /><xsl:text> </xsl:text><xsl:value-of select = "computer/operatingsystem/@servicepack" /><br />
<b>Roles: </b>
<xsl:choose>
	<xsl:when test='count(/computer/roles/role) > 0'>
		<xsl:for-each select ="/computer/roles/role" >	
			<xsl:if test =  "not(position()=1)">, </xsl:if>
			<xsl:value-of select = "@name"/>
		</xsl:for-each> 
		<br />
	</xsl:when>
	<xsl:otherwise>
	None Found<br />
	</xsl:otherwise>
</xsl:choose>

<b>Identifying Number: </b><xsl:value-of select = "computer/machineinfo/@identifyingnumber" /><br />
<b>Scan Time: </b><xsl:value-of select = "computer/generated/@scantime" /><br />

<br />

<h1><a >Table Of Contents</a></h1>
<ol>
	<li><a href="#hardwareplatform">Hardware Platform</a></li>
	<ol>
		<li><a href="#hardwareplatform_general">General Information</a></li>
		<xsl:if test="computer/bios">
			<li><a href="#hardwareplatform_bios">BIOS Information</a></li>
		</xsl:if>
	</ol>
	
	<li><a href="#softwareplatform">Software Platform</a></li>
	<ol>
		<li><a href="#softwareplatform_general">General Information</a></li>
		<xsl:if test="computer/windowscomponents/component">
			<li><a href="#softwareplatform_windowscomponents">Windows Components</a></li>
		</xsl:if>
		<xsl:if test="computer/patches/patch">
			<li><a href="#softwareplatform_patches">Installed Patches</a></li>
		</xsl:if>
		<xsl:if test="computer/installedapplications/msiapplication">
			<li><a href="#softwareplatform_msiapplications">Currently Installed Programs (Windows Installer)</a></li>
		</xsl:if>
		<xsl:if test="computer/installedapplications/regapplication">
			<li><a href="#softwareplatform_regapplications">Currently Installed Programs (Registry)</a></li>
		</xsl:if>
		<xsl:if test="computer/installedapplications/productkey">
			<li><a href="#softwareplatform_productkeys">Product Keys</a></li>
		</xsl:if>
	</ol>

	<li><a href="#storage">Storage</a></li>
	<ol>
		<li><a href="#storage_general">General Information</a></li>
	</ol>
	<li><a href="#network">Network Configuration</a></li>

	<xsl:if test="computer/microsoftiisv2">
		<li><a href="#iis">Internet Information Services</a></li>
		<ol>
		<xsl:if test="computer/microsoftiisv2/iiswebserversetting">
			<li><a href="#iis_www">Web Server</a></li>
		</xsl:if>
		</ol>
	</xsl:if>

	
	<li><a href="#miscellaneous">Miscellaneous Configuration</a></li>
	<ol>
		<xsl:if test="computer/eventlogfiles">
			<li><a href="#miscellaneous_eventlog">Event Log Files</a></li>
		</xsl:if>
		<xsl:if test="computer/localgroups/group">
			<li><a href="#miscellaneous_localgroups">Local Groups</a></li>
		</xsl:if>
		<xsl:if test="computer/localusers/user">
			<li><a href="#miscellaneous_localusers">Local Users</a></li>
		</xsl:if>
		<li><a href="#miscellaneous_printers">Printers</a></li>
		<li><a href="#miscellaneous_regional">Regional Settings</a></li>
		<xsl:if test="computer/processes/process">
			<li><a href="#miscellaneous_processes">Currently running processes</a></li>
		</xsl:if>
		<xsl:if test="computer/services/service">
			<li><a href="#miscellaneous_services">Services</a></li>
		</xsl:if>
		<xsl:if test="computer/shares/share">
			<li><a href="#miscellaneous_shares">Shares</a></li>
		</xsl:if>
		<xsl:if test="computer/win32_startupcommand/command">
			<li><a href="#miscellaneous_startupcommand">Startup Command</a></li>
		</xsl:if>
		<li><a href="#miscellaneous_virtualmemory">Virtual Memory</a></li>
		<xsl:if test="computer/registry">
			<li><a href="#miscellaneous_registry">Windows Registry</a></li>
		</xsl:if>
	</ol>
</ol>


<h1 id="hardwareplatform">Hardware Platform</h1>
<h2 id="hardwareplatform_general">General Information</h2>
<b>Manufacturer: </b><xsl:value-of select = "computer/machineinfo/@manufacturer" /><br />
<b>Product Name: </b><xsl:value-of select = "computer/machineinfo/@productname" /><br />
<b>Identifying Number: </b><xsl:value-of select = "computer/machineinfo/@identifyingnumber" /><br />
<b>Chassis: </b><xsl:value-of select = "computer/machineinfo/@chassis" /><br />


<br /><strong>Processor</strong><br />
<b>Name: </b><xsl:value-of select = "computer/processor/@name" /><br />
<b>Description: </b><xsl:value-of select = "computer/processor/@description" /><br />
<b>Speed: </b><xsl:value-of select = "computer/processor/@speed" /> MHz<br />
<b>L2 Cache Size: </b><xsl:value-of select = "computer/processor/@l2cachesize" /> KB<br />
<b>External Clock: </b><xsl:value-of select = "computer/processor/@externalclock" /> MHz<br />
The system has <xsl:value-of select = "computer/processor/@count" /> processor(s)<br />
<xsl:choose>
	<xsl:when test='computer/processor/@htsystem= "True" '>
		 	<xsl:text>The system has Hyper-Threading enabled.</xsl:text><br />
	</xsl:when>
</xsl:choose>
<br />

<strong>Memory</strong><br />
<b>Total Memory: </b> <xsl:value-of select = "computer/memory/@totalsize" /> MB<br />
<table>
	<tr>
		<th>Bank</th>
		<th>Capacity</th>
		<th>Form</th>
		<th>Type</th>
	</tr>
	<xsl:for-each select ="/computer/memory/memorybank" >
		<tr>
			<td>
				<xsl:choose>
					<xsl:when test='@bank= "" '>
						 	<xsl:text> </xsl:text>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select = "@bank" />
					</xsl:otherwise>
				</xsl:choose>
			</td>
			<td><xsl:value-of select = "@capacity"/> MB</td>
			<td><xsl:value-of select = "@formfactor"/></td>
			<td><xsl:value-of select = "@memorytype"/></td>
		</tr>
	</xsl:for-each> 
</table>

<xsl:if test="computer/win32_cdromdrive/cdrom">
<br />
<strong>CD-ROM</strong><br />
<table>
	<tr>
		<th>Name</th>
		<th>Drive</th>
		<th>Manufacturer</th>
	</tr>
	<xsl:for-each select ="/computer/win32_cdromdrive/cdrom" >
		<tr>
			<td><xsl:value-of select = "@name"/></td>
			<td><xsl:value-of select = "@drive"/></td>
			<td><xsl:value-of select = "@manufacturer"/></td>
		</tr>
	</xsl:for-each> 
</table>
</xsl:if>

<xsl:if test="computer/win32_tapedrive/tapedrive">
<br />
<strong>Tape Drive</strong><br />
<table>
	<tr>
		<th>Name</th>
		<th>Description</th>
		<th>Manufacturer</th>
	</tr>
	<xsl:for-each select ="/computer/win32_tapedrive/tapedrive" >
		<tr>
			<td><xsl:value-of select = "@name"/></td>
			<td><xsl:value-of select = "@description"/></td>
			<td><xsl:value-of select = "@manufacturer"/></td>
		</tr>
	</xsl:for-each> 
</table>
</xsl:if>


<xsl:if test="computer/win32_sounddevice/card">
<br />
<strong>Sound Card</strong><br />
<table>
	<tr>
		<th>Name</th>
		<th>Manufacturer</th>
	</tr>
	<xsl:for-each select ="/computer/win32_sounddevice/card" >
		<tr>
			<td><xsl:value-of select = "@name"/></td>
			<td><xsl:value-of select = "@manufacturer"/></td>
		</tr>
	</xsl:for-each> 
</table>
</xsl:if>

<xsl:if test="computer/videocontroller/adapter">
<br />
<strong>Video Controller</strong><br />
<table>
	<tr>
		<th>Name</th>
		<th>Adapter RAM</th>
		<th>Compatibility</th>
	</tr>
	<xsl:for-each select ="/computer/videocontroller/adapter" >
		<tr>
			<td><xsl:value-of select = "@name"/></td>
			<td><xsl:value-of select = "@adapterram"/> MB</td>
			<td><xsl:value-of select = "@compatibility"/></td>
		</tr>
	</xsl:for-each> 
</table>
</xsl:if>

<xsl:if test="computer/bios">
<h2 id="hardwareplatform_bios">Bios Information</h2>
<b>Bios Version: </b><xsl:value-of select = "computer/bios/@version" /><br />
<b>SMBios Version: </b><xsl:value-of select = "computer/bios/@smbiosversion" /> (Major: <xsl:value-of select = "computer/bios/@smbiosmajorversion" />
	, Minor: <xsl:value-of select = "computer/bios/@smbiosmajorversion" />)<br />
<b>Bios Characteristics: </b>

<xsl:for-each select ="/computer/bios/bioscharacteristics" >	
	<xsl:if test =  "not(position()=1)">, </xsl:if>
	<xsl:value-of select = "@name"/>
</xsl:for-each> 
</xsl:if>

<h1 id="softwareplatform">Software Platform</h1>
<h2 id="softwareplatform_general">General Information</h2>
<b>OS Name: </b><xsl:value-of select = "computer/osconfiguration/@osname" /><xsl:text> </xsl:text><xsl:value-of select = "computer/operatingsystem/@servicepack" /><br />
<b>OS Configuration: </b><xsl:value-of select = "computer/osconfiguration/@computerrole" /> in the <xsl:value-of select = "computer/osconfiguration/@domainname" /><xsl:text> </xsl:text><xsl:value-of select = "computer/osconfiguration/@domaintype" /><br />
<b>Windows Location: </b><xsl:value-of select = "computer/osconfiguration/@windowslocation" /> <br />
<b>Install Date: </b><xsl:value-of select = "computer/osconfiguration/@installdate" /> <br />
<b>Operating System Language: </b><xsl:value-of select = "computer/osconfiguration/@oslanguage" /> <br />
<xsl:if test="computer/lastuser">
	<b>Last Logged on User: </b><xsl:value-of select = "computer/lastuser/@name" /> <br />
</xsl:if>


<xsl:if test="computer/windowscomponents/component">
<h2 id="softwareplatform_windowscomponents">Windows Components</h2>
<xsl:for-each select ="/computer/windowscomponents/component" >
	<xsl:choose>
		<xsl:when test='@level="1" '>
			<xsl:value-of select = "@name" /><br />
		</xsl:when>
		<xsl:otherwise>
			<xsl:value-of select = "@name" /><br />
		</xsl:otherwise>
	</xsl:choose>
</xsl:for-each>
</xsl:if>

<xsl:if test="computer/patches/patch">
<h2 id="softwareplatform_patches">Installed Patches</h2>
<table>
	<tr>
		<th>Patch ID</th>
		<th>Description</th>
		<th>Install Date</th>
	</tr>
	<xsl:for-each select ="/computer/patches/patch" >
		<tr>
			<td><a href="http://support.microsoft.com/kb/{@hotfixid}"><xsl:value-of select = "@hotfixid"/></a></td>
			<td><xsl:value-of select = "@description"/></td>
			<td><xsl:value-of select = "@installdate"/></td>
		</tr>
	</xsl:for-each> 
</table>
</xsl:if>

<xsl:if test="computer/installedapplications/msiapplication">
<h2 id="softwareplatform_msiapplications">Currently Installed Programs (Windows Installer)</h2>
<table>
	<tr>
		<th>Name</th>
		<th>Vendor</th>
		<th>Version</th>
		<th>Install Date</th>
	</tr>
	<xsl:for-each select ="/computer/installedapplications/msiapplication" >
		<tr>
			<td><xsl:value-of select = "@productname"/></td>
			<td><xsl:value-of select = "@vendor"/></td>
			<td><xsl:value-of select = "@version"/></td>
			<td><xsl:value-of select = "@installdate"/></td>
		</tr>
	</xsl:for-each> 
</table>
</xsl:if>

<xsl:if test="computer/installedapplications/regapplication">
<h2 id="softwareplatform_regapplications">Currently Installed Programs (Registry)</h2>
<table>
	<tr>
		<th>Name</th>
		<th>Version</th>
	</tr>
	<xsl:for-each select ="/computer/installedapplications/regapplication" >
		<tr>
			<td><xsl:value-of select = "@productname"/></td>
			<td><xsl:value-of select = "@version"/></td>
		</tr>
	</xsl:for-each> 
</table>
</xsl:if>

<xsl:if test="computer/installedapplications/productkey">
<h2 id="softwareplatform_productkeys">Product Keys</h2>
<table>
	<tr>
		<th>Product</th>
		<th>Key</th>
	</tr>
	<xsl:for-each select ="/computer/installedapplications/productkey" >
		<tr>
			<td><xsl:value-of select = "@productname"/></td>
			<td><xsl:value-of select = "@productkey"/></td>
		</tr>
	</xsl:for-each> 
</table>
</xsl:if>



<h1 id="storage">Storage</h1>
<h2 id="storage_general">General Information</h2>
<xsl:for-each select ="/computer/storage/drives" >
	<strong><xsl:value-of select = "@name"/> - <xsl:value-of select = "@deviceid"/></strong><br />
	<b>Interface: </b><xsl:value-of select = "@interface"/><br />
	<b>Total Disk Size: </b><xsl:value-of select = "@totaldisksize"/> Gb<br />
	<xsl:for-each select ="partition" >
		<xsl:value-of select = "@name"/>\ <xsl:value-of select = "@size"/> Gb (<xsl:value-of select = "@freespace"/> Gb Free) <xsl:value-of select = "@filesystem"/><br />
	</xsl:for-each>
</xsl:for-each> 

<h1 id="network">Network Configuration</h1>
<h2 id="ipconfiguration">IP Configuration</h2>
<xsl:for-each select ="/computer/network/adapter" >
	<strong><xsl:value-of select = "@description"/></strong><br />
	<b>MAC Address: </b><xsl:value-of select = "@macaddress"/><br />
	<xsl:for-each select ="ip" >
		<b>IP Address: </b><xsl:value-of select = "@address"/>/<xsl:value-of select = "@subnetmask"/><br />
	</xsl:for-each>
	<xsl:for-each select ="gateway" >
		<b>Gateway: </b><xsl:value-of select = "@address"/><br />
	</xsl:for-each>
	<b>DNS Domain: </b><xsl:value-of select = "dnsdomain/@name"/><br />
	<b>DNS Server(s): </b>
	<xsl:for-each select ="dnsserver" >	
		<xsl:if test =  "not(position()=1)">,&#160;</xsl:if>
		<xsl:value-of select = "@address"/>
	</xsl:for-each>
	<br /> 	
	<xsl:if test='count(primarywins) > 0'>
		<b>Primary WINS: </b><xsl:value-of select = "primarywins/@address"/><br />
	</xsl:if>
	<xsl:if test='count(secondarywins) > 0'>
		<b>Secondary WINS: </b><xsl:value-of select = "secondarywins/@address"/><br />
	</xsl:if>

	<xsl:if test='count(dhcpserver) > 0'>
		<b>DHCP Server: </b><xsl:value-of select = "dhcpserver/@address"/><br />
	</xsl:if>
</xsl:for-each>
<xsl:if test="computer/network/ip4routes">
	<h2 id="ip4routes">IP Routes</h2>
	<table>
		<tr>
			<th>Destination</th>
			<th>Subnet Mask</th>
			<th>Gateway</th>
		</tr>
		<xsl:for-each select ="/computer/network/ip4routes/route" >
			<tr>
				<td><xsl:value-of select = "@destination"/></td>
				<td><xsl:value-of select = "@mask"/></td>
				<td><xsl:value-of select = "@nexthop"/></td>
			</tr>
		</xsl:for-each> 
	</table>	
</xsl:if>

<xsl:if test="computer/microsoftiisv2">
	<h1 id="iis">Internet Information Services</h1>
	<xsl:if test="computer/microsoftiisv2/iiswebserversetting">
		<h2 id="iis_www">Web Server</h2>
		<xsl:for-each select ="/computer/microsoftiisv2/iiswebserversetting" >
			<h3><xsl:value-of select = "@servercomment"/></h3>
			<b>Home Directory: </b><xsl:value-of select = "homedirectory/@path"/><br />
			<br />
			<table>
				<tr>
					<th>Hostname</th>
					<th>Ip</th>
					<th>Port</th>
				</tr>
				<xsl:for-each select ="serverbindings" >
					<tr>
						<td><xsl:value-of select = "@hostname"/></td>
						<td><xsl:value-of select = "@ip"/></td>
						<td><xsl:value-of select = "@port"/></td>
					</tr>
				</xsl:for-each> 
			</table>	

		</xsl:for-each>
	</xsl:if>
</xsl:if>

<h1 id="miscellaneous">Miscellaneous Configuration</h1>
<xsl:if test="computer/eventlogfiles">
<h2 id="miscellaneous_eventlog">Event Log Files</h2>
<xsl:for-each select ="/computer/eventlogfiles/eventlogfile" >
	<strong><xsl:value-of select = "@name"/></strong><br />
	<b>File: </b><xsl:value-of select = "@file"/><br />
	<b>Maximum Size: </b><xsl:value-of select = "@maximumsize"/> Mb<br />
	<b>Overwritepolicy: </b><xsl:value-of select = "@overwritepolicy"/><br />
	<br />
</xsl:for-each>
</xsl:if>

<xsl:if test="computer/localgroups/group">
<h2 id="miscellaneous_localgroups">Local Groups</h2>
<xsl:for-each select ="/computer/localgroups/group" >
	<b><xsl:value-of select = "@name"/></b><br />
	<xsl:for-each select ="member" >
		- <xsl:value-of select = "@name"/><br />
	</xsl:for-each>
</xsl:for-each>
</xsl:if>

<xsl:if test="computer/localusers/user">
<h2 id="miscellaneous_localusers">Local Users</h2>
<table>
	<tr>
		<th>User</th>
		<th>Description</th>
	</tr>
	<xsl:for-each select ="/computer/localusers/user" >
		<tr>
			<td><xsl:value-of select = "@name"/></td>
			<td><xsl:value-of select = "@description"/></td>
		</tr>
	</xsl:for-each> 
</table>	
</xsl:if>

<h2 id="miscellaneous_printers">Printers</h2>
<b>Print Spooler Location: </b><xsl:value-of select = "computer/printspooler/@location" /><br />

<xsl:if test="computer/printers/printer">
<table>
	<tr>
		<th>Name</th>
		<th>Driver</th>
		<th>Port</th>
	</tr>
	<xsl:for-each select ="/computer/printers/printer" >
		<tr>
			<td><xsl:value-of select = "@name"/></td>
			<td><xsl:value-of select = "@drivername"/></td>
			<td><xsl:value-of select = "@portname"/></td>
		</tr>
	</xsl:for-each> 
</table>
</xsl:if>


<h2 id="miscellaneous_regional">Regional Settings</h2>
<b>Time Zone: </b><xsl:value-of select = "computer/regional/@timezone" /><br />

<xsl:if test="computer/processes/process">
<h2 id="miscellaneous_processes">Currently running processes</h2>
<table>
	<tr>
		<th>Name</th>
		<th>Executable</th>
	</tr>
	<xsl:for-each select ="/computer/processes/process" >
		<tr>
			<td><xsl:value-of select = "@caption"/></td>
			<td><xsl:value-of select = "@executablepath"/></td>
		</tr>
	</xsl:for-each> 
</table>
</xsl:if>


<xsl:if test="computer/services/service">
<h2 id="miscellaneous_services">Services</h2>
<table>
	<tr>
		<th>Name</th>
		<th>Startmode</th>
		<th>Started</th>
		<th>StartName</th>
	</tr>
	<xsl:for-each select ="/computer/services/service" >
	<tr>
		<td><xsl:value-of select = "@name"/></td>
		<td><xsl:value-of select = "@startmode"/></td>
		<td><xsl:value-of select = "@started"/></td>
		<td><xsl:value-of select = "@startname"/></td>
	</tr>
	</xsl:for-each> 
</table>
</xsl:if>

<xsl:if test="computer/shares/share">
<h2 id="miscellaneous_shares">Shares</h2>
<table>
	<tr>
		<th>Name</th>
		<th>Path</th>
		<th>Description</th>
	</tr>
	<xsl:for-each select ="/computer/shares/share" >
		<tr>
			<td><xsl:value-of select = "@name"/></td>
			<td>
			<xsl:choose>
				<xsl:when test='@path= "" '>
					 	<xsl:text> </xsl:text>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select = "@path"/>
				</xsl:otherwise>
			</xsl:choose>
			</td>
			<td>
			<xsl:choose>
				<xsl:when test='@description= "" '>
					 	<xsl:text> </xsl:text>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select = "@description"/>
				</xsl:otherwise>
			</xsl:choose>
			</td>

		</tr>
	</xsl:for-each> 
</table>
</xsl:if>


<xsl:if test="computer/win32_startupcommand/command">
<h2 id="miscellaneous_startupcommand">Startup Command</h2>
<table>
	<tr>
		<th>User</th>
		<th>Name</th>
		<th>Command</th>
	</tr>
	<xsl:for-each select ="/computer/win32_startupcommand/command" >
	<tr>
		<td><xsl:value-of select = "@user"/></td>
		<td><xsl:value-of select = "@name"/></td>
		<td><xsl:value-of select = "@command"/></td>
	</tr>
	</xsl:for-each> 
</table>
</xsl:if>


<h2 id="miscellaneous_virtualmemory">Virtual Memory</h2>
<strong>Pagefile(s)</strong><br />
<xsl:for-each select ="/computer/pagefiles/pagefile" >
	<xsl:value-of select = "@drive"/>\ (<xsl:value-of select = "@initialsize"/> Mb - <xsl:value-of select = "@maximumsize"/> Mb) <br />
</xsl:for-each>

<xsl:if test="computer/registry">
<h2 id="miscellaneous_registry">Windows Registry</h2>
<b>Current Registry Size: </b><xsl:value-of select = "computer/registry/@currentsize" /> Mb<br />
<b>Maximum Registry Size: </b><xsl:value-of select = "computer/registry/@maximumsize" /> Mb<br />
</xsl:if>


</body>
</html>

</xsl:template>
</xsl:stylesheet>
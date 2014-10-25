<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:template match="/computer">
!System - <xsl:value-of select="system/@name" />
:SystemName: <xsl:value-of select="system/@name" />
:OperatingSystem: <xsl:value-of select="operatingsystem/@name" /><xsl:text> </xsl:text><xsl:value-of select = "operatingsystem/@servicepack" />
:Tg:{ data |
   data.Collect{e | ["||{!}*" , e.Item(0), ":*||", e.Item(1),"||",Newline ]} 
}
@@Tg([["FQDN","<xsl:value-of select="fqdn/@name"/>"],["Operating System","<xsl:value-of select="operatingsystem/@name" />"],["Service Pack","<xsl:value-of select="operatingsystem/@servicepack" />"],
["Identifying Number","<xsl:value-of select="machineinfo/@identifyingnumber" />"],["Roles","<xsl:for-each select="roles/role/@name"><xsl:value-of select="."/><xsl:if  test="position() != last()">, </xsl:if></xsl:for-each>"],["Scan Time","<xsl:value-of select="generated/@scantime" />"]])@@


!!Hardware
----

!!!Machine Info
||{!}*Manufacturer:*||<xsl:value-of select="machineinfo/@manufacturer" />||
||{!}*Product Name:*||""<xsl:value-of select="machineinfo/@productname" />""||
||{!}*Identifying Number:*||""<xsl:value-of select="machineinfo/@identifyingnumber" />""||
||{!}*Chassis Type:*||<xsl:value-of select="machineinfo/@chassis" />||

!!!Processor(s)
||{!}*Count:*||<xsl:value-of select="processor/@count"/>||
||{!}*Name:*||<xsl:value-of select="processor/@name"/>||
||{!}*Description:*||<xsl:value-of select="processor/@description"/>||
||{!}*Speed:*||<xsl:value-of select="processor/@speed"/> Mhz||
||{!}*L2 Cache Size:*||<xsl:value-of select="processor/@l2cachesize"/>KB||
||{!}*External Clock:*||<xsl:value-of select="processor/@externalclock"/>Mhz||

!!!Memory Banks- <xsl:value-of select="memory/@totalsize"/>MB
||{!}*Capacity*||{!}*Form Factor*||{!}*Type*||
<xsl:for-each select="memory/memorybank">||<xsl:value-of select="@capacity"/>||<xsl:value-of select="@formfactor"/>||<xsl:value-of select="@memorytype"/>||
</xsl:for-each>

!!!Bios Features - <xsl:value-of select="bios/@version"/> - <xsl:value-of select="bios/@smbiosversion"/>.<xsl:value-of select="bios/@smbiosmajorversion"/>.<xsl:value-of select="bios/@smbbiosminorversion"/>
<xsl:for-each select="bios/bioscharacteristics/@name">
        *<xsl:value-of select="."/>
</xsl:for-each>


!!!Storage Drives
||{!}*Name*||{!C2}*Device ID*||{!}*Interface*||{!}*Size*||
<xsl:for-each select="storage/drives">
<xsl:sort select="@deviceid"/>||{R<xsl:value-of select="count(partition)+2"/>}<xsl:value-of select="@name"/>||{C2}<xsl:value-of select="@deviceid"/>||<xsl:value-of select="@interface"/>||<xsl:value-of select="@totaldisksize"/>GB||
||{!}*Letter*||{!}*Capacity*||{!}*Free Space*||{!}*File System*||
<xsl:for-each select="partition">
<xsl:sort select="@name"/>||<xsl:value-of select="@name"/>||<xsl:value-of select="@size"/>GB||<xsl:value-of select="@freespace"/>GB||<xsl:value-of select="@filesystem"/>||
</xsl:for-each>
</xsl:for-each>

!!!Network Adapters

<xsl:for-each select ="network/adapter" >
||{C2}*<xsl:value-of select = "@description"/>*||
||*MAC Address: *||<xsl:value-of select = "@macaddress"/>||
<xsl:for-each select ="ip" >||*IP Address: *||<xsl:value-of select = "@address"/>/<xsl:value-of select = "@subnetmask"/>||
</xsl:for-each>
<xsl:if test='count(gateway) > 0'><xsl:for-each select ="gateway" >||*Gateway:*||<xsl:value-of select = "@address"/>||
</xsl:for-each></xsl:if>
<xsl:if test='count(dnsdomain) > 0'>||*DNS Domain:*||<xsl:value-of select = "dnsdomain/@name"/>||</xsl:if>
||*DNS Server(s):*||<xsl:for-each select ="dnsserver" ><xsl:if test =  "not(position()=1)">, </xsl:if><xsl:value-of select = "@address"/></xsl:for-each>||
<xsl:if test='count(primarywins) > 0'>||*Primary WINS:*||<xsl:value-of select = "primarywins/@address"/>||
</xsl:if>
<xsl:if test='count(secondarywins) > 0'>||*Secondary WINS:*||<xsl:value-of select = "secondarywins/@address"/>||
</xsl:if>
<xsl:if test='count(dhcpserver) > 0'>||*DHCP Server:*||<xsl:value-of select = "dhcpserver/@address"/>||
</xsl:if>
</xsl:for-each>

<xsl:if test="network/ip4routes">
!!!IP Routes
||{!}*Destination*||{!}*Subnet Mask*||{!}*Gateway*||
<xsl:for-each select ="network/ip4routes/route" >
<xsl:sort select="@destination"/>||<xsl:value-of select = "@destination"/>||<xsl:value-of select = "@mask"/>||<xsl:value-of select = "@nexthop"/>||
</xsl:for-each> 
</xsl:if>



!!Software
----
!!!Operation System
||{!}*Operating System:*||<xsl:value-of select="osconfiguration/@osname" />||
||{!}*System Role:*||<xsl:value-of select="osconfiguration/@computerrole"/>||
||{!}*Domain Name:*||<xsl:value-of select="osconfiguration/@domainname" />||
||{!}*Windows Location:*||<xsl:value-of select="osconfiguration/@windowslocation" />||
||{!}*Install Date:*||<xsl:value-of select="osconfiguration/@installdate" />||
||{!}*Operating System Language:*||<xsl:value-of select="osconfiguration/@oslanguage" />||
<xsl:if test="computer/lastuser">
	||{!}*Operating System Language:*||<xsl:value-of select="osconfiguration/@oslanguage" />||<xsl:value-of select = "computer/lastuser/@name" />||
</xsl:if>

!!!Windows Components
||{!}*Class*||{!}*Component*||
<xsl:for-each select="windowscomponents/component">
<xsl:sort select="@classname"/>||<xsl:value-of select="@classname"/>||<xsl:value-of select="@name"/>||
</xsl:for-each>

!!!Patches
||{!}*Hotfix ID*||{!}*Description*||
<xsl:for-each select="patches/patch">
<xsl:sort select="@hotfixid"/>||<xsl:value-of select="@hotfixid"/>||<xsl:value-of select="@description"/>||
</xsl:for-each>

!!!Installed Applications (Windows Installer)
||{!}*Product*||{!}*Vendor*||{!}*Version*||{!}*Install Date*||
<xsl:choose>
<xsl:when test="count(installedapplications/msiapplication) != 0">
<xsl:for-each select="installedapplications/msiapplication">
<xsl:sort select="@productname"/>||<xsl:value-of select="@productname"/>||<xsl:value-of select="@vendor"/>||<xsl:value-of select="@version"/>||<xsl:value-of select="@installdate"/>||
</xsl:for-each>
</xsl:when>
<xsl:otherwise>||{C4} No software applications were retrieved||
</xsl:otherwise>
</xsl:choose>

!!!Installed Applications (Registry)
||{!}*Product*||{!}*Version*||
<xsl:choose>
<xsl:when test="count(installedapplications/regapplication) != 0">
<xsl:for-each select="installedapplications/regapplication">
<xsl:sort select="@productname"/>||<xsl:value-of select="@productname"/>||<xsl:value-of select="@version"/>||
</xsl:for-each>
</xsl:when>
<xsl:otherwise>||{C4} No software applications were retrieved||
</xsl:otherwise>
</xsl:choose>

!!Miscellaneous Configuration
----

<xsl:if test="eventlogfiles">
!!!Event Log Files
<xsl:for-each select ="eventlogfiles/eventlogfile" >
*<xsl:value-of select = "@name"/>*
||*File:*||<xsl:value-of select="@file"/>||
||*Maximum Size:*||<xsl:value-of select="@maximumsize"/> Mb||
||*Overwritepolicy:*||<xsl:value-of select="@overwritepolicy"/>||
</xsl:for-each>
</xsl:if>

!!!Users and Groups
*Users*
<xsl:for-each select="localusers/user/@name">
        *<xsl:value-of select="."/>
</xsl:for-each>
*Groups*
<xsl:for-each select="localgroups/group/@name">
        *<xsl:value-of select="."/>
</xsl:for-each>

!!!Printing
<xsl:if test="count(printspooler) &gt; 0">||*Print Spooler Location:*||<xsl:value-of select="printspooler/@location"/>||
</xsl:if>

<xsl:if test="printers/printer">
||{!}*Name*||{!}*Driver*||{!}*Port*||
<xsl:for-each select ="printers/printer">
<xsl:sort select="@name"/>||<xsl:value-of select = "@name"/>||<xsl:value-of select = "@drivername"/>||<xsl:value-of select = "@portname"/>||
</xsl:for-each>
</xsl:if>

!!!Regional Settings
||{!}*Time Zone:*||<xsl:value-of select="regional/@timezone" />||

<xsl:if test="processes/process">
!!!Currently running processes
||{!}*Name*||{!}*Executable*||
<xsl:for-each select ="processes/process" >
<xsl:sort select="@caption"/>||<xsl:value-of select = "@caption"/>||<xsl:value-of select = "@executablepath"/>||
</xsl:for-each> 
</xsl:if>

<xsl:if test="services/service">
!!!Services
||{!}*Name*||{!}*Startmode*||{!}*Started*||{!}*StartName*||
<xsl:for-each select ="/computer/services/service" >
<xsl:sort select="@name"/>||<xsl:value-of select = "@name"/>||<xsl:value-of select = "@startmode"/>||<xsl:value-of select = "@started"/>||<xsl:value-of select = "@startname"/>||
</xsl:for-each> 
</xsl:if>

<xsl:if test="shares/share">
!!!Shares
||{!}*Name*||{!}*Path*||{!}*Description*||
<xsl:for-each select ="shares/share" >
<xsl:sort select="@name"/>||<xsl:value-of select = "@name"/>||<xsl:value-of select = "@path"/>||<xsl:value-of select = "@description"/>||
</xsl:for-each> 
</xsl:if>

!!!Virtual Memory
*Pagefile(s)*
<xsl:for-each select ="pagefiles/pagefile" >
||*<xsl:value-of select = "@drive"/>\*|| (<xsl:value-of select = "@initialsize"/> Mb - <xsl:value-of select = "@maximumsize"/> Mb)||
</xsl:for-each>

<xsl:if test="registry">
!!!Windows Registry
||*Current Registry Size:*||<xsl:value-of select = "registry/@currentsize" /> Mb||
||*Maximum Registry Size:*||<xsl:value-of select = "registry/@maximumsize" /> Mb||
</xsl:if>

</xsl:template>
</xsl:stylesheet>
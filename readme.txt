*********************************************************************
                     SYDI-Server 2.3 Readme File                    
                             2009-01-24
*********************************************************************

---------------------------------------------------------------------
                        ~ TABLE OF CONTENTS ~
---------------------------------------------------------------------
1. What is SYDI?

	1.1. Author
	1.2. License

2. Files Included in SYDI-Server

3. System Requirements

4. Running SYDI-Server

5. What Next?
   
   5.1. SYDI Word output
   5.2. SYDI XML output
    5.2.1. SYDI-Server XML 2 Word
    5.2.2. SYDI-Overview
    5.2.3. FlexWiki
   5.3. Running against multiple computers

6. Known issues

7. I want more!

8. Feedback

9. Updates

---------------------------------------------------------------------

1. What is SYDI?
----------------

SYDI is a project which purpose is to help people in writing
documentation for their networks.

SYDI consists of three different packages. SYDI-Server which is
intended to document Windows servers in general and SYDI-SQL
which is intended to document MS-SQL servers and SYDI-Exchange
which is used to document Microsoft Exchange organizations.

SYDI is written in VBScript.

SYDI's home page can be found here:
http://sydiproject.com


My blog where I among other things write about SYDI, I also
write short fiction stories related to IT Security. Please
check it out at:
http://ogenstad.net

1.1. Author
-----------

Patrick Ogenstad (patrick.ogenstad@netsafe.se) has created SYDI. I
works as a Network Consultant at a company in Sweden
called Netsafe (http://www.netsafe.se).

1.2. License
------------

As you might have noticed SYDI doesn't cost much, it is however licensed
under the BSD license.


2. Files Included in SYDI-Server
--------------------------------

examples
 empty.xml		- An empty shell for ss-xml2word.vbs use this as
			  a template when you are documenting your systems
 howto.xml		- Examples of how you can use the written
			  documentation files
 sample.xml		- Example of what a documentation file could look
			  like
language
 lang_danish.xml	- Danish Language File for ss-xml2word.vbs
 lang_dutch.xml		- Dutch Language File for ss-xml2word.vbs
 lang_english.xml	- English Language File for ss-xml2word.vbs
 lang_german.xml	- German Language File for ss-xml2word.vbs
 lang_italian.xml	- Italian Language File for ss-xml2word.vbs
 lang_norwegian.xml	- Norwegian Language File for ss-xml2word.vbs
 lang_portuguese.xml	- Portuguese Language File for ss-xml2word.vbs
 lang_swedish.xml	- Swedish Language File for ss-xml2word.vbs

tools
 ss-xml2word.vbs	- Creates a Word document out of a SYDI-Server XML
			  file
 sydi-audit-localgroups.vbs	- Creates an overviewable Excel document listing
			  members in local groups from member servers and clients
 sydi-overview.vbs	- Creates an overviewable Excel document out of
			  multiple SYDI-Server XML files
 sydi-transform.vbs	- Transforms SYDI-Server XML files based on a given
			  transformation file
 sydi-wrapper.vbs	- A Wrapper script to run SYDI-Server against
			  multiple computers

xml
 flexwikiserver.xsl	- Unmaintained transformation file for FlexWiki
 serverhtml.xsl		- Transformation file XML->Html
 sydi-html-styles.xsl	- Styles for serverhtml.xsl (change it to suit your
			  needs)

changelog.txt		- Changes from previous versions
license.txt		- SYDI-Server license file
readme.txt		- This File
subscribe.html		- Subscribe to the SYDI Newsletter
sydi-server.vbs		- The SYDI-Server Script

3. System Requirements
----------------------

Operating System:	Windows 2000, Windows XP, Windows Server 2003,
                 	Windows Vista, Windows NT 4.0*

Microsoft Word:		Word 2000, Word XP, Word 2003, Word 2007

Services:		SYDI collects information using WMI so the
			Windows Management Instrumentation service must
			be running.

Note that Microsoft Word is only a required on the computer you are
running the script from, and only if you are using Microsoft Word
as the output format.

*) Not supported but some parts work, you have to install WMI for NT4
   to get it working. But seriously, don't run NT4.

4. Running SYDI-Server
----------------------

You don't have to install SYDI. You just have to unzip the files to where
you want it.

The best way to run SYDI-Server is to open a command shell and run the
script with.

To avoid getting a lot of messageboxes I recommend that you run the script
from the command-line with "cscript sydi-server.vbs".

If you don't want to run the script with its default settings type:
"cscript sydi-server.vbs -h"

The only files needed to run SYDI-Server is the program itself,
sydi-server.vbs. If you are planning on using XML you might want
to check the files in the xml directory.

5. What Next?
-------------

The document that SYDI creates is intended to be used as a base for your
server documentation. A more comprehensive tutorial can be found at:
http://ogenstad.net/2006/10/18/how-to-document-servers-with-sydi-part-1-of-3/

5.1. SYDI Word output
---------------------

The idea is that you should run the script against a server and then change
the text appearing inside brackets.

5.2. SYDI XML output
--------------------

To test the XML options run:
cscript.exe sydi-server.vbs -tServer1 -ex -sh -oServer1.xml

Then copy the resulting Server1.xml file into the xml directory and
open the server1.xml in your browser.

You can also use the script sydi-transform.vbs to convert the
xml file into a html file.

5.2.1. SYDI-Server XML 2 Word
-----------------------------

The script SYDI-Server XML 2 Word (ss-xml2word.vbs), lets you create a
word document from a SYDI-Server XML file. A few reasons why you would want
to do this are:

1) You have servers in an environment where Word isn't availible, for
   example in a DMZ
2) You want to write documentation as well as collect data. If you've
   written documentation in word and later want to scan the computer
   again you would have to recreate the written part. With ss-xml2word
   the written part is kept in a separate XML file and is written to
   the final word document.
3) If you want the documentation to be in another language than English.

At the moment the only other language maintained is Swedish. If you want
support your native language you can edit a language.xml file, contact me
for help if you don't know what to do and want to distribute it. It's
very easy, all you have to do is to change the text inside the "translation"
attributes.

To run ss-xml2word.vbs you execute:

cscript.exe ss-xml2word.vbs -xServer1.xml -llang_english.xml -sServer1_docs.xml

Where Server1.xml is a file you have produced from SYDI-Server,
Server1_docs.xml is a file you have written yourself, using a base template
from the Examples directory. Create text tags at the location you want them to
appear.
<text>This is a text tag</text>
   
If you are using a source xml file and you don't see any results in the final
word document, try opening the xml file in a browser and see if you get any errors.

5.2.2. SYDI-Overview
--------------------

You can use sydi-overview.vbs to scan multiple SYDI-Server xml
documents and create an Excel file which gives you a quick overview
of your computers.

cscript.exe sydi-overview.vbs -x[directory to your xml files]

Use quotes if you have spaces in your path.

5.2.3. FlexWiki
---------------

Please note FlexWiki is supported but unmaintained. This is because I
(Patrick) don't use it, if you want support for it let me know.

The XML files can be exported to FlexWiki (http://www.flexwiki.com/)

What you do is first create the XML file

cscript.exe sydi-server.vbs -tServer1 -ex -oServer1.xml

Then you use the sydi-transform script to create a FlexWiki file

cscript.exe sydi-transform.vbs -xServer1.xml -sflexwikiserver.xsl -oServer1.wiki

Then you just drop the Server1.wiki file to your FlexWiki folder.


5.3. Running against multiple computers
---------------------------------------

If you want to target multiple computers you should use the script
sydi-wrapper.vbs. You will have to edit the script to specify the location
of sydi-server.vbs, where you want to place the output files and
which options you want to use from SYDI-Server.

6. Known issues
---------------

1) When targeting Windows Server 2003 servers the Windows Installer applications
   might be missing. This is because the WMI provider for Windows Installer
   isn't installed by default. You can add this with Add/Remove Windows components
   Management and Monitoring Tools / WMI Windows Installer Provider.

7. I want more!
---------------

Ok, this looks good but I want more! If you have any feature requests
which you want implemented in a future version of SYDI-Server (or any
other SYDI product), you should visit:
http://sourceforge.net/tracker/?group_id=116471&atid=674897

You can add feature requests anonymously but if you register an account
at Sourceforge you will receive an email when the feature request is updated. It is also easier for me to have someone to contact if I don't understand the
request. If I don't know what you want and I'm unable to contact you I will just delete the request.


8. Feedback
-----------

In my opinion one of the best things about writing open source stuff
is getting feedback. So please if you have any comments good or bad
please let me know. I am interested in hearing what you think of the
software and how you are using it. So please take your time to send an
email to patrick.ogenstad@netsafe.se. You can also post a comment in the
forum found at http://sourceforge.net/projects/sydi/.

9. Updates
----------

If you want updates you should check out the newsletter at
http://sydiproject.com/email or use the subscribe.html file included
in this release.

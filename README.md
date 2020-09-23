<div align="center">

## Reading a Configuration File


</div>

### Description

This code shows how to

'read a configuration file and use

'the key-value pair. The configuration

'file must be in "key=value" format.

'(Similar to a java property file)
 
### More Info
 
configuration file name with full path

'The code must be modified to

'use the actual key variables

'from your configuration file.

'The code uses a sample of few keys.

'You must customize it to your needs.

'The property file may also have comments.

'The comment lines must start with #.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Subodh Dash](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/subodh-dash.md)
**Level**          |Intermediate
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[System Services/ Functions](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/system-services-functions__4-23.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/subodh-dash-reading-a-configuration-file__4-6825/archive/master.zip)

### API Declarations

Free to use / free to distribute


### Source Code

```
'Let the configuration file be in
'c:\myproject\myconfile.extension
'And the contents of the file is as follows
#Hash lines are comment lines
#You may add as many hash lines as you want
icsDrive=C
icsProjectDir=ICS
icsMinDiskSpace=5000
icsMaxDiskSpace=100000
icsMailApplication=Outlook.Application
icsEMailTo=mymail@yourhost.com
#End of the configuration file
'
'
'
Option Explicit
DIM FSO
SET FSO = CreateObject ("Scripting.FileSystemObject")
'Declare the variables to be used from the property file
 DIM icsDrive
 DIM icsProjectDir
 DIM icsMinDiskSpace
 DIM icsMaxDiskSpace
 DIM icsMailApplication
 DIM icsEMailTo
 Main
Sub Main
	CALL SetConfigFromFile("c:\myproject\myconfile.extension")
	msgbox icsDrive
	msgbox icsProjectDir
	msgbox icsMinDiskSpace
	msgbox icsMaxDiskSpace
	msgbox icsMailApplication
	msgbox icsEMailTo
End sub
'***Read Configuration File***
Sub SetConfigFromFile(fileName)
 DIM strConfigLine
 DIM fConFile
 DIM EqPos, strLen, varName, varVal
 SET fConFile = fso.OpenTextFile(fileName)
 WHILE NOT fConFile.AtEndOfStream
 	strConfigLine = fConFile.ReadLine
	strConfigLine = TRIM(strConfigLine)
	IF (INSTR(1,strConfigLine,"#",1) <> 1 AND LEN(strConfigLine) <> 0) THEN
		EqPos = INSTR(1,strConfigLine,"=",1)
		strLen = LEN(strConfigLine)
		varName = LCASE(TRIM(MID(strConfigLine, 1, EqPos - 1)))
		varVal = TRIM(MID(strConfigLine, EqPos + 1, strLen - EqPos))
		SELECT CASE varName
			'ADD EACH OCCURRENCE OF THE CONFIGURATION FILE VARIABLES(KEYS)
			CASE LCASE("icsDrive")
				IF varVal <> "" THEN icsDrive = varVal
			CASE LCASE("icsProjectDir")
				IF varVal <> "" THEN icsProjectDir = varVal
			CASE LCASE("icsMinDiskSpace")
				IF varVal <> "" THEN icsMinDiskSpace = varVal
			CASE LCASE("icsMaxDiskSpace")
				IF varVal <> "" THEN icsMaxDiskSpace = varVal
			CASE LCASE("icsMailApplication")
				IF varVal <> "" THEN icsMailApplication = varVal
			CASE LCASE("icsEMailTo")
				IF varVal <> "" THEN icsEMailTo = varVal
		END SELECT
	END IF
 WEND
 fConFile.Close
End Sub
```


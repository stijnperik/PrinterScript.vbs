option Explicit
on error resume next

'------------------------------------------------------------------------------------
'	Settings / Config Section
'------------------------------------------------------------------------------------
Dim arrPrintServers
arrPrintServers = array("")
'^ Add the names or ips of your Print Servers.

Dim strAntiPrintGroup
strAntiPrintGroup = ""
'^ If a machine or user is a member of this group, this print script will exit before making any changes.

Dim arrPrintGroupPrefix
arrPrintGroupPrefix = array("")
'^ Set to empty string to disable,
'  The Name Prefix of all the printer groups.
'  To add another prefix, add: ,"Your printer prefix"

Dim bDeleteLocalPrinters
bDeleteLocalPrinters = FALSE
'^ Set to true to delete local printers.

Dim strDeleteLocalPrintersGroup
strDeleteLocalPrintersGroup = ""
'^ Set to empty string to disable
'  Set the name of a group that will delete local printers.

Dim bEnableOUPrinters
bEnableOUPrinters = FALSE
'^ Set to true to enable ou printer groups.

Dim strPrintOUGroupPrefix
strPrintOUGroupPrefix = ""
'^ Set to empty string to disable,
'  The Name Prefix of the printer groups in each ou.

Dim bNoDefaultLocal
bNoDefaultLocal = FALSE
'^ Set to true to prevent local printers from being the default printer.

Dim strNoDefaultLocalGroup
strNoDefaultLocalGroup = ""
'^ Set to empty string to disable
'  Set the name of a group that will prevent local printers from being the default printer.

Dim strLogFilePath
strLogFilePath = "C:\users\%USER%\print-log.txt"
'^ Set to empty string to disable,
'  Set the log directory for the script.
'  Supports: %COMPUTER% and %USER%
'  Recomended to be disabled.

'------------------------------------------------------------------------------------
'	Set up script enviroment
'------------------------------------------------------------------------------------
Dim objShell
Set objShell = CreateObject( "WScript.Shell" )

Dim objNetwork
Set objNetwork = WScript.CreateObject("WScript.Network")

Dim strComputerName
strComputerName = objShell.ExpandEnvironmentStrings("%ComputerName%")

Dim objWMIService
Set objWMIService = GetObject("winmgmts:\\" & strComputerName & "\root\cimv2")

Dim objFileSystem
Set objFileSystem = CreateObject("Scripting.FileSystemObject")

Dim objSysInfo
Set objSysInfo = CreateObject("ADSystemInfo")

Dim strUserName
strUserName = objShell.ExpandEnvironmentStrings("%UserName%")

'------------------------------------------------------------------------------------
' 	Create and open a log file.
'------------------------------------------------------------------------------------
Dim objLogFile

Dim bLoggin
bLoggin = NOT (strLogFilePath = "")


IF bLoggin THEN
	strLogFilePath = Replace(strLogFilePath, "%USER%", strUserName)
	strLogFilePath = Replace(strLogFilePath, "%COMPUTER%", strComputerName)

	Set objLogFile = objFileSystem.OpenTextFile(strLogFilePath, 2, true)
END IF

Function log(str)
	
	IF bLoggin THEN 
		'objShell.run("cmd.exe /C echo " & str & ">>" & strLogFilePath)
		objLogFile.WriteLine(str)
	END IF

END Function

log("Running on " & strComputerName & " for " & strUserName)

'------------------------------------------------------------------------------------
' 	Util function for prefi checking
'------------------------------------------------------------------------------------

Function hasPrefix(str, prefix)
	hasPrefix = FALSE
	Dim x
	x=0
	Dim loopSize
	loopSize = UBound(prefix) 	
	Do While x <= loopSize
		IF Left(str, Len(prefix(x))) = prefix(x) = TRUE THEN
    			hasPrefix = TRUE
			Exit Do	
		END IF
	x=x+1
	Loop
END Function

'------------------------------------------------------------------------------------
' 	Create an array of assigned printers.
'------------------------------------------------------------------------------------
Dim boolNoPrinters
boolNoPrinters = FALSE

Dim arrMemberOf()
Dim objGroup

Dim intSize
intSize = 0

Function assignPrinter(strPrinter)
	IF strPrinter = strAntiPrintGroup THEN 
		boolNoPrinters = TRUE
		log("WARNING: Found membership for " & strAntiPrintGroup & " print script will exit.")
	ELSEIF strPrinter = strDeleteLocalPrintersGroup THEN
		bDeleteLocalPrinters = TRUE
		log("WARNING: Found membership for " & strDeleteLocalPrintersGroup & " setting 'bDeleteLocalPrinters' to true.")
	ELSEIF strPrinter = strNoDefaultLocalGroup THEN
		bNoDefaultLocal = TRUE
		log("WARNING: Found membership for " & strNoDefaultLocalGroup & " setting 'bNoDefaultLocal' to true.")
	ELSEIF hasPrefix(strPrinter, arrPrintGroupPrefix) THEN
		ReDim Preserve arrMemberOf(intSize)
		arrMemberOf(intSize) = strPrinter
	    intSize = intSize + 1
		log("Found printer group membership: " & strPrinter)
	ELSE
		log("Skipped none printer group: " & strPrinter)
	END IF
END Function


'------------------------------------------------------------------------------------
' 	Assign all the Computers assigned printer groups.
'------------------------------------------------------------------------------------
Dim j
j = 0

Dim strTestedGroup
Dim arrTestedGroups()

Function shouldCheckGroup(strGroupName)
	FOR EACH strTestedGroup in arrTestedGroups

		IF strTestedGroup = strGroupName THEN
			shouldCheckGroup = FALSE
			EXIT Function
		END IF
	NEXT

	ReDim Preserve arrTestedGroups(j)
	arrTestedGroups(j) = strGroupName
    j = j + 1

	shouldCheckGroup = TRUE
END Function

Function assignFromMembership(objGroupInfo)
	
	IF shouldCheckGroup(objGroupInfo.cn) = TRUE THEN
		log("Searching group for printers: " & objGroupInfo.cn)

		Dim objGroup
		Dim objPrinterGroup

		FOR EACH objPrinterGroup IN objGroupInfo.getex("memberof")
			On Error Resume Next
			Set objGroup = GetObject("LDAP://" & objPrinterGroup)
			'x=msgbox(objGroup.cn ,0, "Your Title Here")
			assignPrinter(objGroup.cn)
			assignFromMembership(objGroup)
		NEXT
	ELSE
		log("Skipped searching group: " & objGroupInfo.cn)
		log("Group allready applied")
	END IF
END Function

Function assignFromGroup(strGroupName)
	assignFromMembership(GetObject("LDAP://" & strGroupName))
END Function

assignFromGroup(objSysInfo.ComputerName)
assignFromGroup(objSysInfo.UserName)

'------------------------------------------------------------------------------------
' 	If we have enabled ou printers, find the current ou group
'------------------------------------------------------------------------------------

IF bEnableOUPrinters = TRUE THEN
	
	Function getOU(objInfo) 
		getOU = Mid(objInfo.DistinguishedName, Len(objInfo.Name) + 2)
	END Function

	Dim objComputerInfo
	Set objComputerInfo = GetObject("LDAP://" & objSysInfo.ComputerName)

	Dim strComputerOU
	strComputerOU = getOU(objComputerInfo)
	
'------------------------------------------------------------------------------------
' 	Find all Printer groups inside the ou
'------------------------------------------------------------------------------------
	
	Function assignFromOU(strOU)
		log("Searching for OU Printer Groups in: '" & strOU & "' ")

		Dim objOUInfo
		Set objOUInfo = GetObject("LDAP://" & strOU)

		objOUInfo.Filter = Array("group")

		Dim pref
		Dim objOUGroup

		FOR EACH objOUGroup in objOUInfo
			pref = Mid(objOUGroup.cn, 1, Len(strPrintOUGroupPrefix))
			IF pref = strPrintOUGroupPrefix THEN
				log("Found Printer Group: " & objOUGroup.cn)

				assignFromMembership(objOUGroup)
			END IF
		NEXT

	END Function

	assignFromOU(strComputerOU)

END IF
	
'------------------------------------------------------------------------------------
'	If we are a member of the antio print group, then exit
'------------------------------------------------------------------------------------

IF boolNoPrinters = TRUE THEN
	log("Print script disabled via group, exiting.")

	IF bLoggin THEN
		objLogFile.Close
	END IF

	WScript.Quit
END IF

'------------------------------------------------------------------------------------
'	Grab all the existing printers.
'------------------------------------------------------------------------------------
Dim arrInstalledPrinters
Set arrInstalledPrinters = objWMIService.ExecQuery("Select * From Win32_Printer")

'------------------------------------------------------------------------------------
' 	Delete all existing printers and record witch one is default
'------------------------------------------------------------------------------------
Dim strDefaultPrinter
Dim objInstalledPrinter

log("Removing pre existing printers.")

FOR Each objInstalledPrinter IN arrInstalledPrinters
	
	on error resume next

	IF objInstalledPrinter.Network OR bDeleteLocalPrinters THEN
		log("Deleted printer: " & objInstalledPrinter.Name)
		objInstalledPrinter.Delete_
	END IF

NEXT

'------------------------------------------------------------------------------------
'	Add all the printers assigned to this machine.
'------------------------------------------------------------------------------------

Dim strPrintServer
Dim strPrinterGroup

log("Adding network printers from shares.")

FOR EACH strPrinterGroup IN arrMemberOf
	
	FOR EACH strPrintServer IN arrPrintServers
		
		on error resume next
		objNetwork.AddWindowsPrinterConnection "\\" & strPrintServer & "\" & strPrinterGroup
		log("Adding Printer: \\" & strPrintServer & "\" & strPrinterGroup)

		IF Err.Number = 0 THEN

			EXIT FOR

		ELSE
			
			IF Err.Description <> "" THEN
				log("Failed to add printer: \\" & strPrintServer & "\" & strPrinterGroup)
				log("Error: " & Err.Description)
			END IF
		END IF

		
	NEXT

NEXT

'------------------------------------------------------------------------------------
'	Close then log file.
'------------------------------------------------------------------------------------

IF bLoggin THEN
	objLogFile.Close
END IF
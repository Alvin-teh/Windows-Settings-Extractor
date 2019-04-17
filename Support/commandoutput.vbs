'< ============================================================ Declaration of Variables ============================================================ >
Option Explicit

Dim WSHshell, f, strCombine, StdIn, StdOut, strEnumerateRegistry, strPermission, strRetrievePermission, strCurrentDirectory,strSupportFolderEscaped ,  strOutputFolder, strSupportFolder, testfile, strFullConfig, strComputerName, tmpKerbOutput, arrKerberosVal, folder, RegistryFile, strTime, strWindowsUpdate, strMbsacli, strWsusscn, strAccountList, strKERBEROSTICKETVALIDATIONCLIENT, arrFile, intRow, strTemp, TELines, readStart, readSec, readThird, arrValueNames, arrValueTypes, regType, ii, command, objShell, OutFile, readFileLine, fileLine, field, arrFields, fileLines, filePath, AccountPolicy, splitText, textSID, noSID, nameSID, TEReadFile,changeLine, changeLine2, fso, return, subkey, strKERBEROS, colItems, objItem, strService, strAuditPolicy, strAccountPolicy, registry, sComputer, service, cInstances, oInstance, strValue, i, strRegistry, strNewContents, nameOutput, noOutput, CurrentDate, getSID, strUser, strDomain, RegEx

'Application Declaration
Set WSHshell = CreateObject("WScript.Shell")
Set objShell = CreateObject("Shell.Application")

' File System Variables and Object
Const fsoForReading = 1
Const fsoForWriting = 2
Const fsoForAppending = 8
Set fso = CreateObject("Scripting.FileSystemObject")

'Host name of current workstation
strComputerName = WSHshell.ExpandEnvironmentStrings("%COMPUTERNAME%")
'Retrieve the current directory where the scripts reside
strCurrentDirectory = Wscript.Arguments.item(0)

'Directory to save all output files
strOutputFolder = strCurrentDirectory & "\" & strComputerName

'Directory of support files
strSupportFolder = strCurrentDirectory & "\Support"


'Location of output files
strAccountPolicy = strOutputFolder & "\VB_AccountPolicy.txt"
strAuditPolicy = chr(34) & strOutputFolder & "\VB_AuditPolicy.txt" & chr(34)
strTime = strOutputFolder & "\VB_Time.txt"
strKERBEROS = strOutputFolder & "\VB_KERBEROS.txt"
strKERBEROSTICKETVALIDATIONCLIENT = strOutputFolder & "\VB_KERBEROSTICKETVALIDATIONCLIENT.txt"
strRegistry = strOutputFolder & "\VB_Registry.txt"
strService = strOutputFolder & "\VB_Service.txt"
strWindowsUpdate = chr(34) & strOutputFolder & "\VB_WindowsUpdate.txt" & chr(34)
strPermission = chr(34) & strOutputFolder & "\VB_Permission.txt" & chr(34)
strFullConfig = strOutputFolder & "\VB_FullConfig.txt"

wscript.echo strAccountPolicy

'Location of supporting script files
strCombine = chr(34) & strSupportFolder & "\combine.vbs" & chr(34)
strEnumerateRegistry = chr(34) & strSupportFolder & "\EnumerateRegistry.vbs" & chr(34)
strRetrievePermission = chr(34) & strSupportFolder & "\RetrievePermission.vbs"& chr(34)

'Miscellaneous
sComputer = "."
'< ============================================================ Start of Method Call ============================================================ >
call CreateFolderOutputFiles()
call RetrieveAccountPolicy()
call RetrieveServices()
call RetrieveKerberos()
call RetrieveKerberosTicketValidation()
call RetrieveAuditPolicy()
call RetrieveWindowsUpdate()
call RetrievePermission()
call RetrieveRegistry()
wscript.sleep 5000
call ScanInfo()
call ConvertSID()
call CombineOutputFiles()

'< ============================================================ End of Method Call ============================================================ >

'< ============================================================ General Method: Create folder and output files ============================================================ >
Sub CreateFolderOutputFiles()
	Wscript.echo "CreateFolderOutputFiles()"
	'CREATE text File, Clear text in text file
	If Not(fso.FolderExists(strOutputFolder)) Then
		'Create folder if it does not exist
		wscript.echo "folder does not exist"
		wscript.echo ">>>>>" & strOutputFolder
		wscript.echo ">>>>>" & Wscript.ScriptFullName 
		folder = fso.CreateFolder(strOutputFolder)
	End If
	
	arrFile = Array(strAccountPolicy, strAuditPolicy, strKERBEROS, strKERBEROSTICKETVALIDATIONCLIENT, strRegistry, strService, strWindowsUpdate, strTime, strPermission, strFullConfig)
	
	For i = 0 to Ubound(arrFile) 
		If not fso.FileExists(arrFile(i)) Then
			On Error Resume Next
			Set f = fso.CreateTextFile(arrFile(i), True)
			f.MoveFile arrFile(i), strOutputFolder
		Else
			Set f = fso.OpenTextFile(arrFile(i), fsoForWriting)
			f.Write ""
		End If
		f.Close
	Next 
End Sub 
'< ============================================================ Extraction Method: Retrieve Windows Services  ============================================================ >
Sub RetrieveServices()
	Wscript.echo "RetrieveServices()"
	sComputer = "."
	Set cInstances = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & sComputer & "\root\cimv2:Win32_Service").Instances_  
		For Each oInstance In cInstances
			service = service & "NAME:" & oInstance.Properties_("displayname").Value & ":SERVICE:" & oInstance.Properties_("name").Value & ":MODE:" & oInstance.Properties_("StartMode").Value & ":STARTNAME:" & oInstance.Properties_("StartName").Value 
			service = service & vbCrLf
		Next
	ReplaceStringToFile strService, service
End Sub

'< ============================================================ Extraction Method: Retrieve Kerberos Configuration Settings  ============================================================ >
Sub RetrieveKerberos()
	Wscript.echo "RetrieveKerberos()"
	arrKerberosVal = Array("MaxTicketAge", "MaxRenewAge", "MaxServiceAge", "MaxClockSkew") 
	Set cInstances = GetObject("winmgmts:\\" & sComputer & "\root\rsop\computer")  
	Set colItems = cInstances.ExecQuery("Select * from RSOP_SecuritySettingNumeric") 
	For Each objItem in colItems
		i = 0
		For i = LBound(arrKerberosVal) to UBound(arrKerberosVal)
			If arrKerberosVal(i) = objItem.KeyName Then
				tmpKerbOutput = tmpKerbOutput & vbcrlf & "Key Name: " & objItem.KeyName
			End If
		Next
	Next
	If IsNull(tmpKerbOutput) OR IsEmpty (tmpKerbOutput) Then  
		ReplaceStringToFile strKERBEROS, "KERBEROS: NO"
	Else  
		ReplaceStringToFile strKERBEROS, "KERBEROS: YES"  
	End If  
End Sub

'< ============================================================ Extraction Method: Retrieve Kerberos Ticket Validation Client Configuration Settings  ============================================================ >
Sub RetrieveKerberosTicketValidation()
	Wscript.echo "RetrieveKerberosTicketValidation()"
	arrKerberosVal = Array("TicketValidateClient")
	Set cInstances = GetObject("winmgmts:\\" & sComputer & "\root\rsop\computer")
	Set colItems = cInstances.ExecQuery("Select * from RSOP_SecuritySettingBoolean")
	For Each objItem in colItems
		i=0
		For i = LBound(arrKerberosVal) to UBound(arrKerberosVal)
			If arrKerberosVal(i) = objItem.Keyname Then  
				tmpKerbOutput = tmpKerbOutput & vbcrlf & "POLICY:" & objItem.KeyName & ":SETTING:" & objItem.Setting
			End If
		Next
	Next
	If tmpKerbOutput="" Then
		ReplaceStringToFile strKERBEROSTICKETVALIDATIONCLIENT, "KERBEROS: NO" 
	Else
		ReplaceStringToFile strKERBEROSTICKETVALIDATIONCLIENT, "KERBEROS: YES"
	End If
End Sub

'< ============================================================ Extraction Method:Retrieve Audit & Account Policy Settings  ============================================================ >
Sub RetrieveAuditPolicy()
	Wscript.echo "RetrieveAuditPolicy()"
	'Command to retrieve Audit Policy settings
	objShell.ShellExecute "cmd.exe", "/k auditpol.exe /get /category:* > " & strAuditPolicy, "", "runas", 1
End Sub

'< ============================================================ Extraction Method:Retrieve Account Policy Settings  ============================================================ >
Sub RetrieveAccountPolicy()
	Wscript.echo "RetrieveAccountPolicy()"
	'Command to retrieve Account Policy
	objShell.ShellExecute "cmd.exe", "/k secedit.exe /export /cfg " & chr(34) & strAccountPolicy & chr(34), "", "runas", 1
End Sub

'< ============================================================ Extraction Method:Retrieve Windows Update ============================================================ >
Sub RetrieveWindowsUpdate()
	Wscript.echo "RetrieveWindowsUpdate()"
	'Command to retrieve windowsupdate
	'strWindowsUpdate
	objShell.ShellExecute "cmd.exe", "/k " & "mbsacli.exe /xmlout /catalog wsusscn2.cab > " & strWindowsUpdate ,strSupportFolder, "runas", 1
End Sub

'< ============================================================ Extraction Method:Retrieve Registry settings (Entire HKLM directory)  ============================================================ >
Sub RetrieveRegistry()
	Wscript.echo "RetrieveRegistry()"
	wscript.echo strComputerName
	objShell.ShellExecute "cmd.exe", "/k cscript " & strEnumerateRegistry & " " & chr(34) & strCurrentDirectory & chr(34) & " " & strComputerName & " " & chr(34)& strSupportFolder &chr(34), "", "runas", 1
End Sub

'< ============================================================ Extraction Method:Retrieve Permission settings ============================================================ >
Sub RetrievePermission()
	Wscript.echo "RetrievePermission()"
	objShell.ShellExecute "cmd.exe", "/k cscript " & strRetrievePermission & " > " & strPermission, "", "runas", 1
End Sub

'< ============================================================ General Method:Records the scan time and host name ============================================================ >
Sub ScanInfo()
	Wscript.echo "ScanInfo()"
	Set OutFile = fso.OpenTextFile(strTime, fsoForWriting)
	CurrentDate = Now
	
	'Write current time on the system
	OutFile.WriteLine FormatDateTime(CurrentDate, vbGeneralDate)
	
	'Write computer name
	OutFile.WriteLine strComputerName
	OutFile.Close
End Sub

'< ============================================================ General Method:Convert SID in Account Policy to Readable Format ============================================================ >
Sub ConvertSID()
	Wscript.echo "ConvertSID()"
	'Open the text file for reading if the file is not empty
	Set AccountPolicy = fso.OpenTextFile(strAccountPolicy, fsoForReading)
	
	If fso.GETFILE(strAccountPolicy).Size > 0 Then
		Set AccountPolicy = fso.OpenTextFile(strAccountPolicy, fsoForReading)
		readFileLine = AccountPolicy.ReadAll
	End If
	
	AccountPolicy.Close
	
	'Replace all * and white spaces within the file. This step is necessary to ensure proper SID to name conversion
	readFileLine = Replace(readFileLine, "*", "")
	readFileLine = Replace(readFileLine, " ", "")
	
	Set OutFile = fso.OpenTextFile(strAccountPolicy, fsoForWriting)
	OutFile.WriteLine readFileLine
	OutFile.Close
	
	Set RegEx = New RegExp	
	
	'TO ENSURE THE CORRECT LINE IS BEING USED
	Set AccountPolicy = fso.OpenTextFile(strAccountPolicy, fsoForReading)
	Do Until AccountPolicy.AtEndOfStream
		readFileLine = AccountPolicy.ReadLine
		TELines = AccountPolicy.line - 1
		RegEx.Pattern = "\[Privilege\sRights\]"
		Set strTemp = RegEx.Execute(readFileLine)
		If strTemp.Count > 0 Then
			changeLine = TELines
		Else
			RegEx.Pattern = "\[Version\]"
			Set strTemp = RegEx.Execute(readFileLine)
			If strTemp.Count > 0 Then
				changeLine2 = TELines
			End If
		End If
	Loop
	AccountPolicy.Close
	
	'MAKE SID IN ACCOUNTPOLICY.TXT READABLE
	Set AccountPolicy = fso.OpenTextFile(strAccountPolicy, fsoForReading)
	Do Until AccountPolicy.AtEndOfStream
		readFileLine = AccountPolicy.ReadLine
		TELines = AccountPolicy.line - 1
		If TELines >= changeLine AND TELines <= changeLine2 Then
			Dim intSize
			Dim arrTestArray	
			Dim a, temp, j 
			arrTestArray = Array()
			intSize = 0
		
			RegEx.Pattern = "(.*?)\s*=\s*(.*)"
			Set strTemp = RegEx.Execute(readFileLine)
			If strTemp.Count > 0 Then
				'Capturing group 1: attribute name
				readFileLine = strTemp(0).SubMatches(0) & "="
				'Capturing group 2: Multi-valued settings
				splitText = strTemp(0).SubMatches(1)
				splitText = Split(splitText, ",")
				
				'Call the ReplaceSIDToName to replace the SID with the corresponding name
				For Each textSID in splitText
					noSID = textSID
					
					nameSID = ReplaceSIDToName(textSID)
					ReDim Preserve arrTestArray(intSize)
					arrTestArray(intSize) = nameSID
					intSize = intSize + 1
				Next
				
				'Sort the entries in alphabetical order
				For a = ( UBound(arrTestArray) - 1 ) to 0 Step -1
					For j= 0 to a
						If UCase( arrTestArray( j ) ) > UCase( arrTestArray( j + 1 ) ) Then
							temp = arrTestArray( j + 1 )
							arrTestArray( j + 1 ) = arrTestArray( j )
							arrTestArray( j ) = temp
						End If
					Next
				Next

				For i = LBound(arrTestArray) to UBound(arrTestArray)
					If i = UBound(arrTestArray) Then
						readFileLine = readFileLine & arrTestArray(i)
					Else
						readFileLine = readFileLine & arrTestArray(i) & ", "
					End If
				Next
						
			End If
		End If
		TEReadFile = TEReadFile & readFileLine & vbCrLf
	Loop
	AccountPolicy.Close
	Set OutFile =  fso.OpenTextFile(strAccountPolicy, fsoForWriting)
	OutFile.WriteLine TEReadFile
	OutFile.Close
End Sub

Sub CombineOutputFiles()
Dim input
	REM run another script to combine all output files to VB_FullConfig.txt
	objShell.ShellExecute "cmd.exe", "/k cscript " & strCombine & " " & chr(34) & strCurrentDirectory & chr(34) & " " & strComputerName , "", "runas", 1

End Sub
	
REM CHANGE SID TO NAME
Function ReplaceSIDToName(text)
	sComputer = "."
	RegEx.Pattern = "S-"
	Set strTemp = RegEx.Execute(text)
	
	'If the Regular Expression has a match, it is a valid SID. Proceed to convert the SID to name
	'wscript.echo text
	If strTemp.Count > 0 Then
		
		Set cInstances = GetObject("winmgmts:\\" & sComputer & "\root\cimv2")
		Set colItems = cInstances.Get("Win32_SID.SID='" & text & "'")
		On Error Resume Next
		strUser = colItems.AccountName
		strDomain = colItems.ReferencedDomainName
		
		'Replace the SID with the converted value if able to match locally. Otherwise try to query the Active Directory
		If strUser <> "" Then
			If Len(strDomain) Then
				text = UCase(strDomain) & "\" & UCase(strUser)
				ReplaceSIDToName =  text
			Else
				text = UCase(strDomain) & UCase(strUser)
				ReplaceSIDToName = text
			End If
		'Try to query the Active Directory for the SID
		Else 
			'wscript.echo "directory entry" & text
			ReplaceSIDToName = FindUser (text)
			If ReplaceSIDToName = "" Then
				ReplaceSIDToName = text
			End If
		End If
	Else
		'If the Regular Expression does not match, then it is NOT a valid SID
		ReplaceSIDToName = text
	End if
	
End Function

'APPEND TEXT IN THE TEXT FILE
Sub ReplaceStringToFile(filename, text)
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.OpenTextFile(filename, fsoForAppending)
	f.WriteLine text
	f.close
End Sub

'Searches Directory for the User matching the SID passed as parameter
Function FindUser(userSID)
    Dim auserSID
    Dim objUser, domainName, accountName
	Dim regEx,strReplace
	strReplace=""
    
	'Query Account Name
	'wscript.echo userSID
	Set objUser = GetObject("LDAP://<SID="& userSID & ">")
    accountName=objUser.get("sAMAccountName")
	
	'Query Domain Name
	Set regEx = New RegExp
	regEx.Global  = True
	regEx.Pattern = "(-\d+$)"
	regEx.IgnoreCase = True
	
	auserSID=regEx.Replace(userSID, strReplace)
	wscript.echo auserSID
	SET objUser = GetObject("LDAP://<SID="& auserSID & ">")
	domainName=objUser.get("name")
    
	FindUser = domainName & "\" & accountName
	'wscript.echo FindUser
End Function
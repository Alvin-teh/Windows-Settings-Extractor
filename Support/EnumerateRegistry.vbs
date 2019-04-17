'Force explicit variable declaration.
Option Explicit

'Variable for storing arguments
Dim strComputerName, strCurrentDirectory, strSupportFolder
strCurrentDirectory = Wscript.Arguments.item(0)
strComputerName = Wscript.Arguments.item(1)
strSupportFolder = Wscript.Arguments.item(2)

'< ============================================================ Declaration of Variables ============================================================ >
Const fsoForReading = 1
Const fsoForWriting = 2
Const fsoForAppending = 8

Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_CURRENT_CONFIG = &H80000005

Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY = 3
Const REG_DWORD = 4
Const REG_MULTI_SZ = 7

'Separating the Registry Path into multiple sections 
'Example: HKEY_USERS\.DEFAULT\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer|NoDriveTypeAutoRun
'strHive = HKEY_USERS
'strPath = .DEFAULT\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer
'strValueName = NoDriveTypeAutoRun
Dim strHive, _
strPath, _
strValueName

'Miscellaneous Declaration
Dim fso, sComputer, strRegistryList, strRegistryOutput, oReg, filePath, fileWrite, readFilePath, strValue

Dim RegEx : Set RegEx = New RegExp
Set fso = CreateObject("Scripting.FileSystemObject")
sComputer="."

'File Name Declaration for Input and Output File
strRegistryList = strSupportFolder & "\Registry_List.txt"
strRegistryOutput = strCurrentDirectory & "\" & strComputerName & "\VB_Registry.txt"

'Registry Object Declarion required for StdRegProv
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & sComputer & "\root\default:StdRegProv") 

'Open the Registry_List.txt file for reading
Set filePath = fso.OpenTextFile(strRegistryList , fsoForReading)
Set fileWrite = fso.OpenTextFile(strRegistryOutput, fsoForWriting)

'Brief explanation on the the different kinds of output
Dim explanationString
explanationString = _
">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" & vbCrLf & _
"The purpose of this section serves to address some queries that you might have of the outputs found in this file." & vbCRLf & _ 
"You might observe a couple of different types of output:" & vbCRLf & _
"> HKEY exists: In this situation you will find a value beside the HKEY after the ""="" sign."  & vbCrLf & _
"In some circumstances you might observe an empty field, this means that the HKEY exists but no value was defined. "  & vbCrLf & _
"This phenomena might occur when a setting is configured in the group policy as ""Not defined"" resulting in the creation of an empty value in the registry." & vbCRLf & _
"> HKEY does not exist: In this situation you will find a ""<DOES NOT EXIST>"" tag after the ""="" sign. This means that the HKEY entry does not exist in the registry."  & vbCRLf & _
"" & vbCRLf & _
"NOTE: The information used in this section of the registry script relies entirely on the value from the registry."  & vbCrLf & _
"There are multiple ways to apply a particular Windows setting on a machine, through gpedit, services(if available), or directly to the registry. "  & vbCrLf & _ 
"Some of the settings on gpedit might exhibit a default behavior and this value might not be reflected in the registry if the setting was not explicitly applied" & vbCrLf & _
"" & vbCRLf & _
">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" & vbCRLf
fileWrite.Write explanationString

'< ============================================================ Start of Extraction Function ============================================================ >
'Reads the entire file line by line
Do While Not filePath.AtEndOfStream
	Dim strTemp, _
	HIVE_KEY, _
	arrValueNames, _
	arrValueTypes, _
	writeString, _
	i
	
	readFilePath = filePath.ReadLine
	
	RegEx.Pattern = "^(.*?)\\(.*)\|(.*)"
	
	Set strTemp = RegEx.Execute(readFilePath)
	
	If strTemp.Count > 0 Then
		strHive = UCase(strTemp(0).SubMatches(0)) 'Matching Group 1
		strPath = strTemp(0).SubMatches(1) 'Matching Group 2
		strValueName = strTemp(0).SubMatches(2) 'Matching Group 3 
	
		'We will have to convert the strHive string to the corresponding hive value. This conversion is necessary as required by the StdRegProv: Enumvalues method
		HIVE_KEY = HiveKeyValue(strHive)
	End If
	
	'Filter the registry keys by the value types
	oReg.EnumValues HIVE_KEY, strPath, arrValueNames, arrValueTypes
	
	'Check if either of the arrays arrValueNames or arrValueTypes are empty. 
	'If the array is empty, the registry key will be marked as <DOES NOT EXIST>. Otherwise the program will continue the extraction
	
	If IsArrayDimmed(arrValueNames) = 0 Then
		writeString = strHive & "\" & strPath & "|" & strValueName & "=<DOES NOT EXIST>"
		fileWrite.Write writeString & vbCrLf
	Else 
	On Error Resume Next
	For i = 0 To UBound(arrValueNames)
		strValue = ""
		'Case Insensitive Compare
		If StrComp(arrValueNames(i),strValueName, 1) = 0 Then
		Do 
			Select Case arrValueTypes(i)
				'REG_SZ Data Type
				Case REG_SZ
					oReg.GetStringValue HIVE_KEY, strPath, strValueName, strValue
					regType	= "S"
					writeString = strHive & "\" & strPath & "|" & strValueName & "=" & strValue
					Exit For
				
				'REG_EXPAND_SZ Data Type
				Case REG_EXPAND_SZ
					oReg.GetExpandedStringValue HIVE_KEY, strPath, strValueName, strValue

					writeString = strHive & "\" & strPath & "|" & strValueName & "=" & strValue
					Exit For
				
				'REG_BINARY Data Type
				Case REG_BINARY
					Dim arrBytes, uByte
					
					oReg.GetBinaryValue HIVE_KEY, strPath, strValueName, arrBytes

						For Each uByte in arrBytes
							strValue = uByte
						Next
					writeString = strHive & "\" & strPath & "|" & strValueName & "=" & strValue
					Exit For
				
				'REG_DWORD Data Type
				Case REG_DWORD
					oReg.GetDWORDValue HIVE_KEY, strPath, strValueName, strValue

					writeString = strHive & "\" & strPath & "|" & strValueName & "=" & strValue
					Exit For
					
				'REG_MULTI_SZ Data Type
				Case REG_MULTI_SZ
					Dim arrValues, arrValue
					
					oReg.GetMultiStringValue HIVE_KEY, strPath, strValueName, arrValues				  				

					For Each arrValue in arrValues
						If strValue="" Then
							strValue = arrValue
						Else
							strValue = strValue & "," & arrValue
						End If
					Next
					writeString = strHive & "\" & strPath & "|" & strValueName & "=" & strValue
					Exit For
				
				'Catch any error if the key does not exist
				Case Else
					writeString = strHive & "\" & strPath & "|" & strValueName & "=<DOES NOT EXIST>"
					Exit For
			End Select 
		Loop
		Else
			writeString = strHive & "\" & strPath & "|" & strValueName & "=<DOES NOT EXIST>"
		End If
	Next
	fileWrite.Write writeString & vbCrLf
	End If
Loop

Function HiveKeyValue(hiveKey)
	Select case hiveKey
		Case "HKEY_CLASSES_ROOT"
			HiveKeyValue = HKEY_CLASSES_ROOT
		Case "HKEY_CURRENT_CONFIG"
			HiveKeyValue = HKEY_CURRENT_USER
		Case "HKEY_CURRENT_USER"
			HiveKeyValue = HKEY_CURRENT_USER
		Case "HKEY_LOCAL_MACHINE"
			HiveKeyValue = HKEY_LOCAL_MACHINE
		Case "HKEY_USERS"
			HiveKeyValue = HKEY_USERS
		Case "HKEY_CURRENT_CONFIG"
			HiveKeyValue= HKEY_CURRENT_CONFIG
	End Select
End Function

Function IsArrayDimmed(arr)
   IsArrayDimmed = False
   If IsArray(arr) Then
     On Error Resume Next
     Dim ub : ub = UBound(arr)
     If (Err.Number = 0) And (ub >= 0) Then IsArrayDimmed = True
   End If  
 End Function

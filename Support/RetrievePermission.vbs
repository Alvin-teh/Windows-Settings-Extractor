OPTION EXPLICIT

'********************************************************************
'* Declare main variables
'********************************************************************
	
    CONST CONST_CurrentBuild          	= "5.2"
	
	Dim dirArray
	
	dirArray = Array (_
	"%SystemDrive%", _
	"%SystemRoot%\system32\at.exe", _
	"%SystemRoot%\system32\attrib.exe", _
	"%SystemRoot%\system32\cacls.exe", _
	"%SystemRoot%\system32\debug.exe", _
	"%SystemRoot%\system32\drwatson.exe", _
	"%SystemRoot%\system32\drwtsn.exe", _
	"%SystemRoot%\system32\edlin.exe", _
	"%SystemRoot%\system32\eventcreate.exe", _
	"%SystemRoot%\system32\eventtriggers.exe", _
	"%SystemRoot%\system32\ftp.exe", _
	"%SystemRoot%\system32\net.exe",_
	"%SystemRoot%\system32\net1.exe",_
	"%SystemRoot%\system32\netsh.exe",_
	"%SystemRoot%\system32\rcp.exe",_
	"%SystemRoot%\system32\reg.exe",_
	"%SystemRoot%\system32\regedit.exe",_
	"%SystemRoot%\system32\regedt32.exe",_
	"%SystemRoot%\system32\refsvr32.exe",_
	"%SystemRoot%\system32\rexec.exe",_
	"%SystemRoot%\system32\rsh.exe",_
	"%SystemRoot%\system32\runas.exe",_
	"%SystemRoot%\system32\sc.exe",_
	"%SystemRoot%\system32\subst.exe",_
	"%SystemRoot%\system32\telnet.exe",_
	"%SystemRoot%\system32\tftp.exe",_
	"%SystemRoot%\system32\tlntsvr.exe",_
	"%SystemRoot%\system32\cmd.exe",_
	"%SystemRoot%\system32\eventvwr.exe",_
	"%SystemRoot%\system32\runonce.exe",_
	"%SystemRoot%\system32\syskey.exe",_
	"%SystemRoot%\system32\wscript.exe",_
	"%SystemRoot%\system32\winmsd.exe",_
	"%SystemRoot%\repair\",_
	"c:\boot.ini",_
	"c:\ntdetect.com",_
	"c:\ntldr"_
	)
	
    Dim blnQuiet
    Dim filename_var
    Dim l_Used, q_Used, timewmi_Used, strDefaultDomain, strSystemDomainSid, strSystemDomainName, intPermUpdateCount
    Dim ObjTrustee_o_var_User, OldDaclParentObj(), strOldDaclParentPath, strOldDaclLastIsItAFolderValue, boolOldDACLParentRevokedUserFound
    Dim fso, InitialfilenameAbsPath, QryBaseNameHasWildcards, QryExtensionHasWildcards
    Dim objService, objLocalService, objLocator
    Dim strRemoteServerName, strRemoteShareName, strRemoteUserName, strRemotePassword
    Dim RemoteServer_Used, RemoteUserName_Used
    Dim DisplayDirPath, ActualDirPath
    Dim Global_bool_SID_Used
	
	Dim i, RegEx, strTemps, strTemp, WSHshell
    
    'When working with NTFS Security, we use constants that match the API documentation
    '********************* ControlFlags *********************
    CONST ALLOW_INHERIT  			= 33796		'Used in ControlFlag to turn on Inheritance
								'Same as: 
								'SE_SELF_RELATIVE + SE_DACL_AUTO_INHERITED + SE_DACL_PRESENT
    CONST DENY_INHERIT   			= 37892		'Used in ControlFlag to turn off Inheritance
								'Same as: 
								'SE_SELF_RELATIVE + SE_DACL_PROTECTED + SE_DACL_AUTO_INHERITED + SE_DACL_PRESENT
    Const SE_OWNER_DEFAULTED 			= 1		'A default mechanism, rather than the the original provider of the security 
								'descriptor, provided the security descriptor's owner security identifier (SID). 

    Const SE_GROUP_DEFAULTED 			= 2		'A default mechanism, rather than the the original provider of the security
								'descriptor, provided the security descriptor's group SID. 

    Const SE_DACL_PRESENT 				= 4		'Indicates a security descriptor that has a DACL. If this flag is not set, 
								'or if this flag is set and the DACL is NULL, the security descriptor allows 
								'full access to everyone.

    Const SE_DACL_DEFAULTED 			= 8		'Indicates a security descriptor with a default DACL. For example, if an 
								'object's creator does not specify a DACL, the object receives the default 
								'DACL from the creator's access token. This flag can affect how the system 
								'treats the DACL, with respect to ACE inheritance. The system ignores this 
								'flag if the SE_DACL_PRESENT flag is not set. 

    Const SE_SACL_PRESENT 				= 16		'Indicates a security descriptor that has a SACL. 

    Const SE_SACL_DEFAULTED 			= 32		'A default mechanism, rather than the the original provider of the security 
								'descriptor, provided the SACL. This flag can affect how the system treats 
								'the SACL, with respect to ACE inheritance. The system ignores this flag if 
								'the SE_SACL_PRESENT flag is not set. 

    Const SE_DACL_AUTO_INHERIT_REQ 	= 256		'Requests that the provider for the object protected by the security descriptor 
								'automatically propagate the DACL to existing child objects. If the provider 
								'supports automatic inheritance, it propagates the DACL to any existing child 
								'objects, and sets the SE_DACL_AUTO_INHERITED bit in the security descriptors 
								'of the object and its child objects.

    Const SE_SACL_AUTO_INHERIT_REQ 		= 512		'Requests that the provider for the object protected by the security descriptor 
								'automatically propagate the SACL to existing child objects. If the provider 
								'supports automatic inheritance, it propagates the SACL to any existing child 
								'objects, and sets the SE_SACL_AUTO_INHERITED bit in the security descriptors of 
								'the object and its child objects.

    Const SE_DACL_AUTO_INHERITED 		= 1024		'Windows 2000 only. Indicates a security descriptor in which the DACL is set up 
								'to support automatic propagation of inheritable ACEs to existing child objects. 
								'The system sets this bit when it performs the automatic inheritance algorithm 
								'for the object and its existing child objects. This bit is not set in security 
								'descriptors for Windows NT versions 4.0 and earlier, which do not support 
								'automatic propagation of inheritable ACEs.

    Const SE_SACL_AUTO_INHERITED 		= 2048		'Windows 2000: Indicates a security descriptor in which the SACL is set up to 
								'support automatic propagation of inheritable ACEs to existing child objects. 
								'The system sets this bit when it performs the automatic inheritance algorithm 
								'for the object and its existing child objects. This bit is not set in security 
								'descriptors for Windows NT versions 4.0 and earlier, which do not support automatic 
								'propagation of inheritable ACEs.

    Const SE_DACL_PROTECTED 			= 4096		'Windows 2000: Prevents the DACL of the security descriptor from being modified 
								'by inheritable ACEs. 

    Const SE_SACL_PROTECTED 				= 8192		'Windows 2000: Prevents the SACL of the security descriptor from being modified 
								'by inheritable ACEs. 

    Const SE_SELF_RELATIVE 				= 32768		'Indicates a security descriptor in self-relative format with all the security 
								'information in a contiguous block of memory. If this flag is not set, the security 
								'descriptor is in absolute format. For more information, see Absolute and 
								'Self-Relative Security Descriptors in the Platform SDK topic Low-Level Access-Control.

    '********************* ACE Flags *********************
    CONST OBJECT_INHERIT_ACE  			= 1 	'Noncontainer child objects inherit the ACE as an effective ACE. For child 
							'objects that are containers, the ACE is inherited as an inherit-only ACE 
							'unless the NO_PROPAGATE_INHERIT_ACE bit flag is also set.

    CONST CONTAINER_INHERIT_ACE 		= 2 	'Child objects that are containers, such as directories, inherit the ACE
							'as an effective ACE. The inherited ACE is inheritable unless the 
							'NO_PROPAGATE_INHERIT_ACE bit flag is also set.  

    CONST NO_PROPAGATE_INHERIT_ACE 	= 4 	'If the ACE is inherited by a child object, the system clears the 
							'OBJECT_INHERIT_ACE and CONTAINER_INHERIT_ACE flags in the inherited ACE. 
							'This prevents the ACE from being inherited by subsequent generations of objects.  

    CONST INHERIT_ONLY_ACE	 			= 8 	'Indicates an inherit-only ACE which does not control access to the object
							'to which it is attached. If this flag is not set, the ACE is an effective
							'ACE which controls access to the object to which it is attached. Both 
							'effective and inherit-only ACEs can be inherited depending on the state of
							'the other inheritance flags. 

    CONST INHERITED_ACE		 			= 16 	'Windows NT 5.0 and later, Indicates that the ACE was inherited. The system sets
							'this bit when it propagates an inherited ACE to a child object. 

    CONST ACEFLAG_VALID_INHERIT_FLAGS = 31 	'Indicates whether the inherit flags are valid.  


    'Two special flags that pertain only to ACEs that are contained in a SACL are listed below. 

    CONST SUCCESSFUL_ACCESS_ACE_FLAG 	= 64 	'Used with system-audit ACEs in a SACL to generate audit messages for successful
							'access attempts. 

    CONST FAILED_ACCESS_ACE_FLAG 		= 128 	'Used with system-audit ACEs in a SACL to generate audit messages for failed
							'access attempts. 

    '********************* ACE Types *********************
    CONST ACCESS_ALLOWED_ACE_TYPE 	= 0 	'Used with Win32_Ace AceTypes
    CONST ACCESS_DENIED_ACE_TYPE 		= 1 	'Used with Win32_Ace AceTypes
    CONST AUDIT_ACE_TYPE 				= 2 	'Used with Win32_Ace AceTypes


    '********************* Access Masks *********************

    Dim Perms_LStr, Perms_SStr, Perms_Const
    'Permission LongNames
    Perms_LStr=Array("Synchronize"			, _
		"Take Ownership"					, _
		"Change Permissions"				, _
		"Read Permissions"					, _
		"Delete"							, _
		"Write Attributes"					, _
		"Read Attributes"					, _
		"Delete Subfolders and Files"			, _
		"Traverse Folder / Execute File"		, _
		"Write Extended Attributes"			, _
		"Read Extended Attributes"			, _
		"Create Folders / Append Data"		, _
		"Create Files / Write Data"			, _
		"List Folder / Read Data"	)
    'Permission Single Character codes
    Perms_SStr=Array("E"		, _
		"D"		, _
		"C"		, _
		"B"		, _
		"A"		, _
		"9"		, _
		"8"		, _
		"7"		, _
		"6"		, _
		"5"		, _
		"4"		, _
		"3"		, _
		"2"		, _
		"1"		)
    'Permission Integer
    Perms_Const=Array(&H100000	, _
		&H80000		, _
		&H40000		, _
		&H20000		, _
		&H10000		, _
		&H100		, _
		&H80		, _
		&H40		, _
		&H20		, _
		&H10		, _
		&H8			, _
		&H4			, _
		&H2			, _
		&H1		)

   Dim OverLook_Perms_Const
    'Permission Integer
    OverLook_Perms_Const=Array(&H80000000			, _
    		&H40000000									, _
    		&H20000000									, _
    		&H10000000									)


			
'********************************************************************
'* Start of Main Script
'********************************************************************
    'FSO is used in several funcitons, so lets set it globally.
    Set fso = WScript.CreateObject("Scripting.FileSystemObject")
		
	For i = LBound(dirArray) to UBound(dirArray)
		InitialfilenameAbsPath = dirArray(i)
		Set WSHshell = CreateObject("WScript.Shell")
		InitialfilenameAbsPath =  WSHshell.ExpandEnvironmentStrings(InitialfilenameAbsPath)
	
	'Put statements in loop to be able to drop out and clear variables
		Do

		'Lets get the objService object which is used throughout the script
		If Not SetMainVars() then Exit Do
		
		If QryBaseNameHasWildcards or QryExtensionHasWildcards then
			Select Case DoesPathNameExist(fso.GetParentFolderName(InitialfilenameAbsPath))
			Case 1 'Directory
				
				Call DoTheWorkOnEverythingUnderDirectory(fso.GetParentFolderName(InitialfilenameAbsPath))
			End select
		Else
			
			'If a folder is found with the same name, then we work it as a folder and include files under it.
			Select Case DoesPathNameExist(InitialfilenameAbsPath)
			Case 0 'File does not exist
				FileDoesNotExist(InitialfilenameAbsPath)
			Case 1 'Directory
				Call DoTheWorkOnThisItem(InitialfilenameAbsPath, FALSE)
			Case 2 'File
				Call DoTheWorkOnThisItem(InitialfilenameAbsPath, FALSE)
			End select
		End if

		Exit Do	
		Loop
	Next

'********************************************************************
'* End of Main Script
'********************************************************************


'********************************************************************
'*
'* Sub DoTheWorkOnThisItem()
'* Purpose: Work on File/Folder passed to it, and pass to Work routine
'* Input:   ABSPath - Path to File/Folder
'* Output:  TRUE if Successful, FALSE if not
'*
'********************************************************************

Sub DoTheWorkOnThisItem(byval AbsPath, byval IsItAFolder)
    ON ERROR RESUME NEXT

	Call PrintMsg("")
	Call PrintMsg("**************************************************************************")
	
	If DisplayIt then 
		Call DisplayThisACL(AbsPath)
	End if
	Call PrintMsg("**************************************************************************")

End Sub

Sub FileDoesNotExist(byval AbsPath)
    ON ERROR RESUME NEXT

	Call PrintMsg("")
	Call PrintMsg("**************************************************************************")
	
	Call PrintMsg("Permission Settings for : " & AbsPath)
	Call PrintMsg("File/Folder Does Not Exist")
	
	Call PrintMsg("**************************************************************************")

End Sub
'********************************************************************
'*
'* Sub DisplayThisACL()
'* Purpose: Shows ACL's that are applied to strPath
'* Input:   strPath - string containing path of file or folder, ShowLong - If TRUE, permissions are in long form
'* Output:  prints the acls
'*
'********************************************************************

Sub DisplayThisACL(byval strPath)
	wscript.echo "Permission Settings for : " & strPath
    ON ERROR RESUME NEXT

    Dim objFileSecSetting, objOutParams, objSecDescriptor, objOwner, objDACL_Member
    Dim objtrustee, numAceFlags, strAceFlags, x, strAceType, numControlFlags, ReturnAceFlags, TempSECString
    ReDim arraystrACLS(0)

    'Put statements in loop to be able to drop out and clear variables
    Do
	set objFileSecSetting = objService.Get("Win32_LogicalFileSecuritySetting.Path=""" & Replace(strPath,"\","\\") & """")

	Set objOutParams = objFileSecSetting.ExecMethod_("GetSecurityDescriptor")
	
	set objSecDescriptor = objOutParams.Descriptor

	numControlFlags = objSecDescriptor.ControlFlags

	If IsArray(objSecDescriptor.DACL) then
		Call PrintMsg("")
		Call PrintMsg("Permissions:")
		Call PrintMsg( strPackString("Type", 9, 1, TRUE) & strPackString("Username", 35, 1, TRUE) & strPackString("Permissions", 22, 1, TRUE) & strPackString("Inheritance", 22, 1, TRUE))
		For Each objDACL_Member in objSecDescriptor.DACL
			TempSECString = ""
			ReturnAceFlags = 0
			Select Case objDACL_Member.AceType
			Case ACCESS_ALLOWED_ACE_TYPE
				strAceType = "Allowed"
			Case ACCESS_DENIED_ACE_TYPE
				strAceType = "Denied"
			Case else
				strAceType = "Unknown"
			End select
			Set objtrustee = objDACL_Member.Trustee
			numAceFlags = objDACL_Member.AceFlags
			strAceFlags = StringAceFlag(numAceFlags, numControlFlags, SE_DACL_AUTO_INHERITED, FALSE, ReturnAceFlags)
			TempSECString = SECString(objDACL_Member.AccessMask,TRUE)
			If ReturnAceFlags = 2 then
				If TempSECString = "Read and Execute" then
					TempSECString = "List Folder Contents"
				End if
			End if
			Call AddStringToArray(arraystrACLS, strPackString(strAceType, 9, 1, TRUE) & strPackString(objtrustee.Domain & "\" & objtrustee.Name, 35, 1, TRUE) & strPackString(TempSECString, 22, 1, TRUE) & strPackString(strAceFlags, 22, 1, TRUE),-1)
			Set objtrustee = Nothing
		Next
		For x = LBound(arraystrACLS) to UBound(arraystrACLS)
			Call PrintMsg(arraystrACLS(x))
		Next 
	Else
		Call PrintMsg("")
		Call PrintMsg("No Permissions set")
	End if

	Set objOwner = objSecDescriptor.Owner
	
	Call PrintMsg("")
	Call PrintMsg("Owner: " & objOwner.Domain & "\" & objOwner.Name)

	Exit Do		'We really didn't want to loop
    Loop
    'ClearObjects that could be set and aren't needed now
    Set objOwner = Nothing
    Set objSecDescriptor = Nothing
    Set objDACL_Member = Nothing
    Set objtrustee = Nothing
    Set objOutParams = Nothing
    Set objFileSecSetting = Nothing

End Sub


'********************************************************************
'*
'* Function SECString()
'* Purpose: Converts SEC bitmask to a string
'* Input:   intBitmask - integer and ReturnLong - Boolean
'* Output:  String Array
'*
'********************************************************************

Function SECString(byval intBitmask, byval ReturnLong)

    On Error Resume Next
    Dim LongName, X

    SECString = ""

    Do
		
	For X = LBound(Perms_LStr) to UBound(Perms_LStr)
    		If ((intBitmask And Perms_Const(X)) = Perms_Const(X)) then
			If Perms_SStr(X) <> "" then
				SECString = SECString & Perms_SStr(X)
			End if
    		End if
	Next

	Select Case SECString
	Case "DCBA987654321", "EDCBA987654321"
		SECString = "F"								'Full control
		LongName = "Full Control"	
	Case "BA98654321", "EBA98654321"
		SECString = "M"								'Modify
		LongName = "Modify"
	Case "B98654321", "EB98654321"
		SECString = "XW"								'Read, Write and Execute
		LongName = "Read, Write and Execute"
	Case "B9854321", "EB9854321"
		SECString = "RW"								'Read and Write
		LongName = "Read and Write"
	Case "B8641", "EB8641"
		SECString = "X"								'Read and Execute
		LongName = "Read and Execute"
	Case "B841", "EB841"
		SECString = "R"								'Read
		LongName = "Read"
	Case "9532", "E9532"
		SECString = "W"								'Write
		LongName = "Write"
	Case Else
		If SECString = "" then
			LongName = "Special (Unknown)"

		Else
			If LEN(SECString) = 1 then
				For X = LBound(Perms_SStr) to UBound(Perms_SStr)
					If StrComp(SECString,Perms_SStr(X),1) = 0 Then
						LongName = "Advanced (" & Perms_LStr(X) & ")"
						Exit For
					End if
				Next
			Else
				LongName = "Special (" & SECString & ")"
			End if
		End if
	End Select

	Exit Do
    Loop

    If ReturnLong Then SECString = LongName

End Function

'********************************************************************
'*
'* Function StringAceFlag()
'* Purpose: Changes the AceFlag into a string
'* Input:   numAceFlag =      This items ACEFlag
'*          numControlFlags = This Descriptors AceFlag
'*          FlagToCheck =     This lists Auto_Inherited bit to check for
'*          ReturnShort =     If True then we will return a short version
'*          ReturnAceFlags =  Final numAceFlags value after changes (leaves real one alone
'* Output:  String of our codes
'*
'********************************************************************

Function StringAceFlag(ByVal numAceFlags, ByVal numControlFlags, ByVal FlagToCheck, ByVal ReturnShort, ByRef ReturnAceFlags)

    On Error Resume Next

    Dim TempShort, TempLong

    Do
	If numAceFlags = 0 then 
		TempShort = "Implicit"
		TempLong = "This Folder Only"
		Exit Do
	End if
	If numAceFlags > FAILED_ACCESS_ACE_FLAG then
		numAceFlags = numAceFlags - FAILED_ACCESS_ACE_FLAG
	End if
	If numAceFlags > SUCCESSFUL_ACCESS_ACE_FLAG then
		numAceFlags = numAceFlags - SUCCESSFUL_ACCESS_ACE_FLAG
	End if
	If ((numAceFlags And INHERITED_ACE) = INHERITED_ACE) then
		TempShort = "Inherited"
		numAceFlags = numAceFlags - INHERITED_ACE
		TempLong = "Inherited"
	Else
		TempShort = "Implicit"
		TempLong = "Implicit"
	End If

	ReturnAceFlags = numAceFlags 

	If numControlFlags > DENY_INHERIT then
		numControlFlags = numControlFlags - DENY_INHERIT
	End if
	If numControlFlags > ALLOW_INHERIT then
		numControlFlags = numControlFlags - ALLOW_INHERIT
	End if

	Select Case numAceFlags 
	Case 0
		TempLong = "This Folder Only"
	Case 1							'OBJECT_INHERIT_ACE
		TempLong = "This Folder and Files"
	Case 2							'CONTAINER_INHERIT_ACE
		TempLong = "This Folder and Subfolders"
	Case 3
		TempLong = "This Folder, Subfolders and Files"
	Case 9
		TempLong = "Files Only"
	Case 10
		TempLong = "Subfolders only"
	Case 11
		TempLong = "Subfolders and Files only"
	Case Else
		If ((numControlFlags And FlagToCheck) = FlagToCheck) then
			TempShort = "Inherited"
			TempLong = "Inherited"
		End if
	End Select
	Exit Do
    Loop

    If ReturnShort then
	StringAceFlag = TempShort
    Else
	StringAceFlag = TempLong
    End if

End Function

'********************************************************************
'*
'* Function AddStringToArray()
'* Purpose: Adds a string to an array (allowing duplicates) and allows for a member index number
'* Input:   Array and Member
'* Output:  Returns Index Number
'* Notes:   If intUseIndex is -1 then we just want to ReDim the array to be 1 larger and use the
'*          last member. If its any other number than we want to use that number if available.
'*
'********************************************************************

Function AddStringToArray(ByRef theArray, byval theMember, byval intUseIndex)

    On Error Resume Next

    Dim UseThisNumber

    Do

	AddStringToArray = UBound(theArray)

	If intUseIndex <> -1 then
		If intUseIndex > AddStringToArray then
			AddStringToArray = intUseIndex 
		End if
		UseThisNumber = intUseIndex
	Else
		'We will always increment by 1 so the first member is 0 or blank
		AddStringToArray = AddStringToArray + 1
		UseThisNumber = AddStringToArray
	End if

	ReDim Preserve theArray(AddStringToArray)

	theArray(UseThisNumber) = theMember
	
	Exit Do
    Loop

End Function


'********************************************************************
'*
'* Function SetMainVars()
'* Purpose: Checks a FilePath for existance and sets Global Var's
'* Input:   Nothing
'* Output:  Boolean TRUE if worked, FALSE if failed
'* Notes:   None
'*
'********************************************************************

Function SetMainVars()
    On Error Resume Next

    Dim strTempServer, strTempShare, objFileShare

    Do

	SetMainVars = FALSE
	strTempServer = ""
	strTempShare = ""
	
	'Create Locator object to connect to remote CIM object manager
	Set objLocator = CreateObject("WbemScripting.SWbemLocator")

	Set objLocalService = objLocator.ConnectServer ("", "root/cimv2")

	'Connect to the namespace which is either local or remote
	If RemoteServer_Used then
		If RemoteUserName_Used then
			Set objService = objLocator.ConnectServer (strRemoteServerName, "root/cimv2", strRemoteUserName, strRemotePassword)
		Else
			Set objService = objLocator.ConnectServer (strRemoteServerName, "root/cimv2")
		End if
	Else
		Set objService = objLocator.ConnectServer ("", "root/cimv2")
	End if

	objLocalService.Security_.impersonationlevel = 3

	objLocalService.Security_.Privileges.AddAsString "SeSecurityPrivilege", TRUE

	ObjService.Security_.impersonationlevel = 3

	objService.Security_.Privileges.AddAsString "SeSecurityPrivilege", TRUE


	If fso.GetBaseName(filename_var) <> "" then
		QryBaseNameHasWildcards = HasWildcardCharacters(fso.GetBaseName(filename_var))
	Else
		QryBaseNameHasWildcards = FALSE
	End if
	If fso.GetExtensionName(filename_var) <> "" then
		QryExtensionHasWildcards = HasWildcardCharacters(fso.GetExtensionName(filename_var))
	Else
		QryExtensionHasWildcards = FALSE
	End if

	If strRemoteShareName <> "" Then
		set objFileShare = objService.Get("Win32_Share.Name=""" & strRemoteShareName & """")
		If objFileShare.Path <> "" then
			ActualDirPath = objFileShare.Path
			DisplayDirPath = "\\" & strRemoteServerName & "\" & strRemoteShareName
		Else
			Call PrintMsg("Error, Share """ & strRemoteShareName & """ does not have a Path set.")
			Call PrintMsg("Script can not continue.")
			Exit Do
		End if
	
		InitialfilenameAbsPath = fso.GetAbsolutePathName(Replace(filename_var, DisplayDirPath, ActualDirPath, 1, 1, 1))
	End if

	SetMainVars = TRUE
	Exit Do
    Loop

    'ClearObjects that could be set and aren't needed now
    Set objFileShare = Nothing

End Function

'********************************************************************
'*
'* Function DoesPathNameExist()
'* Purpose: Checks a FilePath for existance and what it is (file/folder)
'* Input:   File path string
'* Output:  Integer (0 for doesn't exist, 1 for Folder, 2 for File)
'* Notes:   None
'*
'********************************************************************

Function DoesPathNameExist(byVal strFilePath)

    On Error Resume Next

    Dim objFileSystemSet, objPath, strQuery

    Do
	DoesPathNameExist = 0
	If strFilePath = "" then Exit Do

	If RemoteServer_Used then
		strQuery = "Select Name, FileType from Cim_LogicalFile Where Name=""" & Replace(strFilePath,"\","\\") & """"
        	Set objFileSystemSet = objService.ExecQuery(strQuery,,0)
		
	    	for each objPath in objFileSystemSet
			If objPath.Name <> "" then
				Select Case UCase(objPath.FileType)
				Case "FILE FOLDER"
					DoesPathNameExist = 1
				Case Else
					DoesPathNameExist = 2
				End select
				Exit For
			End if
	    	next
	Else
		If fso.FolderExists(strFilePath) Then
			DoesPathNameExist = 1
		Else
			If fso.FileExists(strFilePath) Then
				DoesPathNameExist = 2
			End if
		End If
	End if
	Exit Do		'We really didn't want to loop
    Loop
    'ClearObjects that could be set and aren't needed now
    Set objPath = Nothing
    Set objFileSystemSet = Nothing

End Function

'********************************************************************
'*
'* Function GetThisArg()
'* Purpose: Gets the next argument, returns TRUE if there were no errors
'* Input:   ArgNumber of next argument
'* Output:  Returns String of next argument or blank if there was none, updates argnumber
'*
'********************************************************************

Function GetThisArg(ByRef intArgNumber)

    On Error Resume Next

    Dim BoolComplete, intLeftCharHex

    Do
	GetThisArg = ""
	If Wscript.arguments.count = 0 then                		'No arguments have been received
        	Exit Do
	End If

	If intArgNumber = (Wscript.arguments.count) then 		'No more to get
        	Exit Do
	End If

	BoolComplete = FALSE

	intLeftCharHex = ASC(Left(Wscript.arguments.Item(intArgNumber),1))
	GetThisArg = Wscript.arguments.Item(intArgNumber)
	Select Case intLeftCharHex
	Case 34, 145, 146, 147, 148	'Quotation marks (different kinds)
		If InStr(2, Wscript.arguments.Item(intArgNumber), Chr(intLeftCharHex),1) > 0 then
			'Then we know that the quotes is closed in the same argument.
		Else
			If intArgNumber < Wscript.arguments.count - 1 then
				While BoolComplete = FALSE
					intArgNumber = intArgNumber + 1
					GetThisArg = GetThisArg & " " & Wscript.arguments.Item(intArgNumber)
					If InStr(1, Wscript.arguments.Item(intArgNumber), Chr(intLeftCharHex),1) > 0 then
						'Then we found the quote pair, lets end it.
						BoolComplete = TRUE
					End if
				Wend 
			End if
		End if
	End Select

	Exit Do
	
    Loop

End Function

'********************************************************************
'*
'* Function strPackString()
'* Purpose: Attaches spaces to a string to increase the length to intWidth.
'* Input:   strString   a string
'*          intWidth   the intended length of the string
'*          blnAfter    specifies whether to add spaces after or before the string
'*          blnTruncate specifies whether to truncate the string or not if
'*                      the string length is longer than intWidth
'* Output:  strPackString is returned as the packed string.
'*
'********************************************************************

Function strPackString(byval strString, ByVal intWidth, byval blnAfter, byval blnTruncate)

    ON ERROR RESUME NEXT

    Do

	If intWidth > Len(strString) Then
        	If blnAfter Then
			strPackString = strString & Space(intWidth-Len(strString))
        	Else
			strPackString = Space(intWidth-Len(strString)) & strString & " "
        	End If
	Else
		If blnTruncate Then
			strPackString = Left(strString, intWidth-1) & " "
        	Else
			strPackString = strString & " "
		End If
	End If
	Exit Do
    Loop

End Function

'********************************************************************
'*
'* Sub PrintMsg()
'* Purpose: Prints a message on screen if blnQuiet = FALSE.
'* Input:   strMessage      the string to print
'* Output:  strMessage is printed on screen if blnQuiet = FALSE.
'*
'********************************************************************

Sub PrintMsg(byval strMessage)
    If Not blnQuiet then
		Wscript.Echo strMessage
    End If
End Sub


'********************************************************************
'*                                                                  *
'*                           End of File                            *
'*                                                                  *
'********************************************************************

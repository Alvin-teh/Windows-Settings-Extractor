Dim currentFileSize, previousFileSize, strComputerName, strCurrentDirectory

ii = 0
Const fsoForReading = 1
Const fsoForWriting = 2
Set WSHshell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

strCurrentDirectory = Wscript.Arguments.item(0)
strComputerName = Wscript.Arguments.item(1)

wscript.echo "======================= Extraction Status ======================="
'< ============================================================ Check that file extraction status ============================================================ >
' Tracking Mechanism to check that all extractions have been completed before proceeding to file compression
' Currently checking Registry and Windows Update only
Call ExtractionCompletionTracker(strCurrentDirectory & "\" & strComputerName &"\VB_Registry.txt")
Call ExtractionCompletionTracker(strCurrentDirectory & "\" & strComputerName &"\VB_WindowsUpdate.txt")
Call ExtractionCompletionTracker(strCurrentDirectory & "\" & strComputerName &"\VB_AccountPolicy.txt")

Call CombineOutputFile(strCurrentDirectory, strComputerName)

'********************************************************************
'*
'* Function: CombineOutputFile(sCurrentDirectory, strComputerName)
'* Purpose:  Combine all the output files into VB_FullConfig.txt
'* Input:    Strings
'* Output:   NIL
'* Notes:    NIL
'*
'********************************************************************
Sub CombineOutputFile(sCurrentDirectory, strComputerName)

	strOutputFolder = sCurrentDirectory & "\" & strComputerName
	strFullConfig = strOutputFolder & "\VB_FullConfig.txt"

	'Location of output files
	Set OutFile = fso.OpenTextFile(strFullConfig, fsoForWriting)
	Set folder = fso.GetFolder(strOutputFolder)
		For Each file in folder.Files
			
			wscript.echo file.path
			nameOutput = "#" & file.path
			Set testfile = fso.OpenTextFile(file.path)
			If testfile.AtEndOfStream Then
			   readFilePath = ""
			Else
				readFilePath = testfile.ReadAll
			End If
			testfile.close
			noOutput = "::1000" & ii + 1 & "::"
			OutFile.WriteLine nameOutput
			OutFile.WriteLine noOutput
			OutFile.WriteLine readFilePath
			Set readFilePath = Nothing
			ii = ii + 1
		Next
	OutFile.close
	wscript.echo 
	wscript.echo "Extraction has been completed"
	Set StdIn = WScript.StdIn
	Set StdOut = WScript.StdOut
	StdOut.WriteLine "Press enter to close"
	Do While Not WScript.StdIn.AtEndOfLine
		input = WScript.StdIn.Read(1)
		Exit Do
	Loop
	Call CloseCommandWindow(Wscript.Arguments.item(0), Wscript.Arguments.item(1))
End Sub

'********************************************************************
'*
'* Function: ExtractionCompletionTracker(fileDir)
'* Purpose:  Check the completion of the extraction process by attempting to open the file for appending. 
'* If extraction is still in progress, this method will catch the error and continue to loop until the extraction is completed. 
'* Input:    Full path of file to monitor (String)
'* Output:   NIL
'* Notes:    Method will continue looping until the file is released for appending.
'*
'********************************************************************
Sub ExtractionCompletionTracker(fileDir)

	' Strategy: Attempt to open the specified file in 'append' mode.
    ' Does not appear to change the 'modified' date on the file.
    ' Works with binary files as well as text files.

    ' Only 'ForAppending' is needed here. Define these constants
    ' outside of this function if you need them elsewhere in
    ' your source file.
	Const ForAppending = 8

    IsWriteAccessible = False

    Dim oFso : Set oFso = CreateObject("Scripting.FileSystemObject")

    On Error Resume Next

    Dim nErr : nErr = 0
    Dim sDesc : sDesc = ""
    Dim oFile : Set oFile = oFso.OpenTextFile(fileDir, ForAppending)
    If Err.Number = 0 Then
        oFile.Close
        If Err Then
            nErr = Err.Number
            sDesc = Err.Description
			wscript.echo nErr & " - " &  sDesc
        Else
            IsWriteAccessible = True
			' 
			'wscript.echo fileDir & " is write accessible"
        End if
    Else
        Select Case Err.Number
            Case 70
                ' Permission denied because:
                ' - file is open by another process
                ' - read-only bit is set on file, *or*
                ' - NTFS Access Control List settings (ACLs) on file
                '   prevents access
				Wscript.echo "Script is still running. Please wait ..."
				wscript.sleep 5000
				Call ExtractionCompletionTracker(fileDir)
            Case Else
                ' 52 - Bad file name or number
                ' 53 - File not found
                ' 76 - Path not found

                nErr = Err.Number
                sDesc = Err.Description
        End Select
    End If
    On Error GoTo 0

    If nErr Then
        Err.Raise nErr, , sDesc
    End If
End Sub


REM close all the background cmd
Sub CloseCommandWindow(currDirectory, computerName)
	sComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & sComputer & "\root\CIMV2") 
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Process Where Name ='cmd.exe'") 
	
	Call ArchiveFolder (currDirectory, computerName)
	wscript.echo "Closing all command prompt windows"
	
	For Each objItem in colItems
		objItem.Terminate()
		wscript.echo computerName
	Next
End Sub


'********************************************************************
'*
'* Function: ArchiveFolder (currDirectory, computerName)
'* Purpose:  Zip up the output into a zip file
'* Input:    Current Directory, hostname (computername)
'* Output:   NIL
'* Notes:    Method will zip up all the files in the directory named after the hostname
'*
'********************************************************************
Sub ArchiveFolder (currDirectory, computerName)
	wscript.echo "Compressing " & computerName & " to " &computerName & ".zip"
    With CreateObject("Scripting.FileSystemObject")
        sFolder = currDirectory & "\" & computerName
		zipName = currDirectory & "\" & computerName & ".zip"

        With .CreateTextFile(zipName, True)
            .Write Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, chr(0))
        End With
    End With

    With CreateObject("Shell.Application")
        .NameSpace(zipName).CopyHere .NameSpace(sFolder).Items

        Do Until .NameSpace(zipName).Items.Count = _
                 .NameSpace(sFolder).Items.Count
            WScript.Sleep 10000 
        Loop
    End With
	wscript.echo "Folder Compression has been completed"
	WScript.Sleep 5000 
End Sub
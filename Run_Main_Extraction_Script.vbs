Dim objShell, strCurrentDirectory, WSHshell
Set objShell = CreateObject("Shell.Application")
Set WSHshell = CreateObject("WScript.Shell")
strCurrentDirectory = chr(34) & WSHshell.currentdirectory & chr(34)

objShell.ShellExecute "cmd.exe", "/k cscript " & strCurrentDirectory & "\Support\" & "commandoutput.vbs " & strCurrentDirectory, "", "runas", 1

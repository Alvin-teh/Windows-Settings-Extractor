# Windows Settings Extractor

## Description

This is an extraction script I have developed for extracting configurations from Windows for the purpose of performing security configuration review against the Center for Internet Security (CIS) baselines.

## Why was VBScript (.vbs) chosen?

VBScript was chosen as it is supported by all default installation of Windows. Thus, there is no need to install any other programming languages to run this script.

## How do I run the script?

1) Log in using an account with Administrator privilege
2) Unzip the zip file. You will see the following files in the folder.
    > Root Directory
      - Run_Main_Extraction_Script.vbs
      - Read.txt
    > Support folder
      - combine.vbs
      - commandoutput.vbs
      - EnumerateRegistry.vbs
      - RetrievePermission.vbs
      - wusscan.dll
      - wsusscn2.cab.dat
      - wsusscn2.cab

3) Open a command prompt (cmd.exe) window with administrator privilege (shift + right click on cmd.exe)

4) Execute the following command
  > cd %temp%
  > cd "Windows_Script<Ver1-7>"
  > cscript Run_Main_Extraction_Script.vbs

5) A prompts will appear requesting for approval to run the script. Approve these request

6) The script will start running. 

7) A notification will appear When the script has completed the extraction

8) Press "Enter" when prompted to complete the extraction

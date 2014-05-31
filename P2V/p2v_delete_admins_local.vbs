' delete_p2vadmins.vbs
' ---------------------------------------------------------------' 
Option Explicit
Dim objFSO, objShell, objTextFile, objFile, objGroup, objUser, objComputer
Dim strDirectory, strFile, strText, strPeople, strUser, strComputer, strInputFilename, strLine, colAccounts
strInputFileName = InputBox("Enter full path of list of targets *.txt: ")
strFile = strInputFileName
' Create the File System Object to read text
Set objFSO = CreateObject("Scripting.FileSystemObject") 
' OpenTextFile Method needs a Const value
' ForAppending = 8 ForReading = 1, ForWriting = 2
Const ForReading = 1
Set objTextFile = objFSO.OpenTextFile _
(strFile, ForReading, True) 
' Start the Do .... Loop Until
Do 
strLine = objTextFile.ReadLine
strComputer = strLine
strUser = "p2vadmin"
Set objComputer = GetObject("WinNT://" & strComputer)
objComputer.Delete "user", strUser
Loop Until objTextFile.AtEndOfLine = true
' Loop ends, tidy up the text file.
objTextFile.Close

WScript.Echo "Done!"
WScript.Quit 
' End of VBScript 
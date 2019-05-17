'The StdIn, StdOut, and StdErr properties and methods work when running the script
'with the CScript.exe host executable file only.
'An "Invalid Handle" error is returned when run with WScript.exe.

ShowWelcomeBox()
ShowMode ("nomode")
UsrMode = SelectMode()




'###########################################################################
'This function takes input form user
Public Function ConsoleInput()
ConsoleInput = WScript.StdIn.ReadLine
End Function

'###########################################################################
'This Function controls output On command prompt
Public Sub ConsoleOutput (strMsg, strMode)
Select Case strMode
	Case LCase("logonly")
		WriteLogFile (strMsg)
	Case LCase("nolog")
		WScript.StdOut.WriteBlankLines(1)
		WScript.StdOut.WriteLine (strMsg)
	Case LCase ("verbose")
		WScript.StdOut.WriteBlankLines(1)
		WScript.StdOut.WriteLine (strMsg)
		WriteLogFile (strMsg)
End Select
End Sub

'###########################################################################
' This function show information about VBSX_VALIDATOR
Public Sub ShowWelcomeBox()

WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.Write "    "
WScript.StdOut.Write "**************************************************"
WScript.StdOut.WriteBlankLines(2)
WScript.StdOut.WriteLine VBTab & VBTab & "  " & "VBSX_VALIDATOR version 1.0.2"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine VBTab & VBTab & "     " & "BULK XML FILE VALIDATOR"
WScript.StdOut.WriteLine VBTab & vbTab & "   " & "Last Updated: November 2013"
WScript.StdOut.WriteLine VBTab & vbTab & "Tushar Sharma | www.testoxide.com"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.Write "    "
WScript.StdOut.Write "**************************************************"
WScript.StdOut.WriteBlankLines(2)

End Sub

'###########################################################################

Public Function WriteLogFile (strMsg)

Dim strDirName
Dim strFileName
strDirName = "C:\EuroDataGen"
strFileName = "C:\EuroGenLog" & "-" & Day(Date) & MonthName(Month(Date),True) & Right((Year(Date)),2) & ".txt"

Set ObjFSO = CreateObject("Scripting.FileSystemObject")
Set ObjTextFile = ObjFSO.OpenTextFile(strDirName & strFileName, 8, True)
ObjTextFile.WriteLine (strMsg & vbCrLf)

Set ObjTextFile = Nothing
Set ObjFSO = Nothing

End Function

'###########################################################################

'This function calls Readme.txt

Public Sub ShowReadMe()

Set ObjFSO = CreateObject("Scripting.FileSystemObject")
Set ObjTextFile = ObjFSO.OpenTextFile(GetCurrentDir() & "\ReadMe.txt", 1, False)
Set ObjTextFile = Nothing
Set ObjFSO = Nothing

End Sub


'###########################################################################
'This function sets the directory path of ReadMe.txt 

Public Function GetCurrentDir(strPath)

'GetCurrentDir = Left(strPath,InStrRev(strPath,"\"))
sCurrPath = Left(WScript.ScriptFullName,(Len(WScript.ScriptFullName)) - (Len(WScript.ScriptName)))
GetCurrentDir = sCurrPath

'The Other Methods -
'Dim sCurrPath
'sCurrPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")
'Msgbox sCurrPath

End Function


'###########################################################################
'This Function sets input values for operating modes 
Sub ShowMode (strMode)
WScript.StdOut.WriteBlankLines(2)
Select Case strMode
	Case LCase ("1")
		WScript.StdOut.WriteLine "MODE :- <INTERACTIVE SINGLE>"
	Case LCase ("2")
		WScript.StdOut.WriteLine "MODE :- <INTERACTIVE BATCH>"
	Case LCase ("3")
		WScript.StdOut.WriteLine "MODE :- <SILENT BATCH>"
	Case Else
		WScript.StdOut.WriteLine "MODE :- <NOT SET!>"
End Select
'WScript.StdOut.WriteBlankLines(1)
End Sub


'###########################################################################
'Displays operating modes And takes input
Public Function SelectMode()
Dim StrMode

WScript.StdOut.WriteLine "SELECT OPERATING MODE? [Example: Input 1 for 'INTERACTIVE SINGLE']"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine "1. INTERACTIVE SINGLE"
WScript.StdOut.WriteLine "2. INTERACTIVE BATCH"
WScript.StdOut.WriteLine "3. SILENT BATCH"
WScript.StdOut.WriteBlankLines(1)

strMode = ConsoleInput()
SelectMode = strMode 

End Function

'###########################################################################


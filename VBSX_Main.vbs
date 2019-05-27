'###################################################################################################
'# SCRIPT NAME: VBSX_Main.vbs
'#
'# DESCRIPTION:
'# Free script utility for silent XML/XSD validation of large sized files.
'# The VBSX_Validator is designed to validate large XML files.The project 
'# exposes the power and flexibility of VB Script language and demonstrates how it 
'# could be utilized for some specific XML related operations and automation.
'# 
'# 
'# NOTES:
'# 
'# 
'# 
'# 
'#
'# PLATFORM: Win7/8/Server | PRE-REQ: Script/Admin Privilege
'# LAST UPDATED: Wed, 25 May 2019 | AUTHOR: Tushar Sharma
'##################################################################################################



'If WScript.Arguments.length = 0 Then
'   Set objShell = CreateObject("Shell.Application")
'   objShell.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 3
'      WScript.Quit
'End If  

'###########################################################################

Dim LogHandle

ShowWelcomeBox()
ShowMode ("nomode")
strOpsMode = SelectMode()
ShowMode (strOpsMode)
Set LogHandle = CreateLogWriter()

If ValidateInput(strOpsMode) Then
	Select Case strOpsMode
		Case "1"
			Call SingleFileValidation()
		Case "2"
			Call BulkFileValidation()
	End Select
Else 
	ConsoleOutput "INVALID CHOICE!", "verbose", LogHandle
End If



Call ExitApp()

'###########################################################################

Public Sub SingleFileValidation()
Dim strFilePath, ObjSchemaCache

	ConsoleOutput "PROVIDE FULL PATH TO XML FILE (.xml) ? ", "verbose", LogHandle
	strFilePath = ConsoleInput()
	
	Set objXMLFile = LoadXML(strFilePath)
	ConsoleOutput "", "verbose", LogHandle
	
	ConsoleOutput "PROVIDE FULL PATH TO SCHEMA FILE (.xsd) ? ", "verbose", LogHandle
	strFilePath = ConsoleInput()	
	Set objXSDFile = LoadXML(strFilePath)
	
	Set ObjSchemaCache = LoadSchemaCache(objXSDFile)
	
	Call ValidateXML (objXMLFile, ObjSchemaCache)
	
	Set objXMLFile = Nothing
	Set objXSDFile = Nothing

End Sub	


'###########################################################################

Public Function LoadXML(strXmlPath)

Dim ObjParseErr, ObjXML
Dim IsWait

Set ObjXML = CreateObject ("MSXML2.DOMDocument.6.0")
    
'    WScript.Echo("Microsoft XML Core Services (MSXML) 6.0 is not installed.\n"
'          +"Download and install MSXML 6.0 from http://msdn.microsoft.com/xml\n"
'          +"before continuing.");
	
	With ObjXML
		'Set First Level DOM Properties
		.async = False
		.validateOnParse = False
		.resolveExternals = False
	End With
	
	ConsoleOutput "", "verbose", LogHandle
	ConsoleOutput "<INFO> Loading XML with First-Level XMLDOM Properties", "verbose", LogHandle
	ObjXML.Load (strXmlPath)
		
	If ObjXML.ParseError.errorCode <> 0 Then
		Call ParseError (ObjXML.parseError, ObjXML)
		Set LoadXML = False
	Else
		ConsoleOutput "<INFO> Configuring Second-Level XMLDOM Properties", "verbose", LogHandle
		'ConsoleOutput "<INFO> Setting Up XML Namespace Property ...", "verbose", LogHandle
		'ObjXML.setProperty "SelectionNamespaces", "xmlns:ns='" + ObjXML.documentElement.namespaceURI + "'"
		ConsoleOutput "<INFO> Setting Up XML Selection Language Property ... XPath", "verbose", LogHandle
		ObjXML.setProperty "SelectionLanguage", "XPath"
		ConsoleOutput "<INFO> Setting full parsing via MultipleErrorMessages Property", "verbose", LogHandle
		ObjXML.setProperty "MultipleErrorMessages", True
		ConsoleOutput "<INFO> Setting 'resolve externals' to false (disabled)", "verbose", LogHandle
		ObjXML.setProperty "ResolveExternals", False 
		ConsoleOutput "<INFO> Second-Level XMLDOM Properties configured successfully ..! ", "verbose", LogHandle
		ConsoleOutput "<INFO> File Loaded Successfully ..." & strXmlPath, "verbose", LogHandle
		Set LoadXML = ObjXML
	End If

End Function

	
'###########################################################################

Public Function LoadSchemaCache (objXSDFile)

Dim ObjSchemaCache, strNsURI
	
	Set ObjSchemaCache = CreateObject("MSXML2.XMLSchemaCache.6.0")
	ObjSchemaCache.validateOnload = False ' This method applies to only [Schema Cache] not (XSD or XML)
	
	ConsoleOutput "", "verbose", LogHandle
	ConsoleOutput "<INFO> Creating Schema Cache Collection", "verbose", LogHandle
	
	'Get targetNamespace property from XSD
	strNsURI = GetNamespaceURI (objXSDFile)
	
	'Load XSD from the Path
	ObjSchemaCache.Add strNsURI, objXSDFile
	ConsoleOutput "<INFO> Schema Cache Loaded Successfully from ... " & objXSDFile.url, "verbose", LogHandle
	
	Set LoadSchemaCache = ObjSchemaCache
	
End Function

'###########################################################################

Public Function GetNamespaceURI (ObjXML)

Dim strNsURI

strNsURI = ObjXML.documentElement.getAttribute("targetNamespace")

If strNsURI <> "" Then
	GetNamespaceURI = strNsURI
	ConsoleOutput "<INFO> Adding 'targetNamespace' " & strNsURI, "verbose", LogHandle
Else

End If

End Function

'###########################################################################

Public Function ValidateXML (ObjXMLDoc, ObjXSDDoc)

Set ObjXMLDoc.Schemas = ObjXSDDoc

ConsoleOutput "", "verbose", LogHandle
ConsoleOutput vbTab & "******************" & vbTab & "<STARTING VALIDATION> " & vbTab & " ******************" & vbCrLf , "verbose", LogHandle

If ObjXMLDoc.readystate = 4 Then
	Set ObjXParseErr = ObjXMLDoc.validate()
	Call ParseError (ObjXParseErr, ObjXMLDoc)
End If

End Function

'###########################################################################


Public Function ParseError (ByVal ObjParseErr, ObjXMLDoc)
Dim strResult, ErrFound
ErrFound = 0

Select Case ObjParseErr.errorCode
	Case 0
		ConsoleOutput "", "verbose", LogHandle
		strResult = "<INFO> XML SCHEMA VALIDATION: SUCCESS ! " & vbCrLf & ObjXMLDoc.url & vbCrLf '& ObjXSDDoc.url & vbCrLf
		ConsoleOutput strResult, "verbose", LogHandle
		ParseError = True
	Case Else
	   If (ObjParseErr.AllErrors.length > 1) Then
	      ConsoleOutput "<ERROR> VALIDATION FAILED WITH MULTIPLE ERRORS !" & vbCrLf, "verbose", LogHandle
	      For Each ErrorItem In ObjParseErr.AllErrors
			strResult = "[" & ErrFound+1 & "]" & " ERROR REASON :" & _
			vbCrLf & "    ------------" & vbCrLf & ErrorItem.reason & vbCrLf & _
			"Error Code: " & ErrorItem.errorCode & ", Line: " & _
							 ErrorItem.Line & ", Character: " & _		
							 ErrorItem.linepos & ", Source: " & _
							 Chr(34) & ErrorItem.srcText & vbCrLf & vbCrLf & _
							 "XPath Value : " & vbCrLf & ErrorItem.errorXPath & vbCrLf 
	      'ConsoleOutput ObjXMLDoc.url
	      ConsoleOutput strResult, "verbose", LogHandle
	     ErrFound = ErrFound + 1
	      Next
	   Else
			ConsoleOutput "<ERROR> VALIDATION FAILED WITH A SINGLE ERROR !" & vbCrLf, "verbose", LogHandle
			strResult = " ERROR REASON :" & _
			vbCrLf & " ------------" & vbCrLf & ObjParseErr.reason & vbCrLf & _
			"Error Code: " & ObjParseErr.errorCode & ", Line: " & _
							 ObjParseErr.Line & ", Character: " & _		
							 ObjParseErr.linepos & ", Source: " & _
							 Chr(34) & ObjParseErr.srcText & vbCrLf & vbCrLf & _
							 "XPath Value : " & vbCrLf & ObjParseErr.errorXPath & vbCrLf 
	      	ConsoleOutput strResult, "verbose", LogHandle
	      	ErrFound = ErrFound + 1
	   End If

End Select

If ErrFound > 0 Then
	ParseError = False
End If

End Function

'###########################################################################


'This function takes input form user
Public Function ConsoleInput()
ConsoleInput = WScript.StdIn.ReadLine
End Function

'###########################################################################

'This Function controls output On command prompt
Public Sub ConsoleOutput (strMsg, strMode, objFSOHandle)

Select Case strMode
	Case LCase("logonly")
		objFSOHandle.WriteLine (strMsg)
	Case LCase("nolog")
		WScript.StdOut.WriteLine (strMsg)
	Case LCase ("verbose")
		WScript.StdOut.WriteLine (strMsg)
		objFSOHandle.WriteLine (strMsg)
End Select

End Sub


'###########################################################################
'This function sets the directory path of ReadMe.txt 

Public Function CreateLogWriter()

sCurrPath = Left(WScript.ScriptFullName,(Len(WScript.ScriptFullName)) - (Len(WScript.ScriptName)))
strFileName = "vbsx-Validator" & "_" & Day(Date) & MonthName(Month(Date),True) & Right((Year(Date)),2) & ".txt"

Set ObjFSO = CreateObject("Scripting.FileSystemObject")
Set ObjTextFile = ObjFSO.OpenTextFile(sCurrPath & strFileName, 8, True)

Set CreateLogWriter = ObjTextFile 

Set ObjTextFile = Nothing
Set ObjFSO = Nothing

'The Other Methods -
'sCurrPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")

End Function


'###########################################################################

Function ValidateInput (strArgsIn)

Dim strValidInput, strArg, strFound
strFound = False
strValidInput = Array("1","2")

	For Each strArg In strValidInput
		If strArg = strArgsIn Then
			strFound = True
		End If
	Next
	
	If Not(strFound) Then
		ValidateInput = False
	Else 
		ValidateInput = True
	End If

End Function 

'###########################################################################

Function IsDocReady (ObjXML)

IsWait = True

Do While Not (ObjXML.readystate = 4)
	ConsoleOutput "Working on large size document, do you wish to continue (y/n)?", "nolog", LogHandle
	strResponse = UCase(ConsoleInput())
	If (strResponse = "N") Or (strResponse = "NO") Then
		IsWait = False
		Exit Do
	Else 
		WScript.Sleep(5000)
	End If
Loop 

If Not(IsWait) Then
	Call ExitApp()
End If

End Function


'###########################################################################
'This Function sets input values for operating modes 
Sub ShowMode (strMode)

WScript.StdOut.WriteBlankLines(2)
Select Case strMode
	Case LCase ("1")
		WScript.StdOut.WriteLine "OPERATING MODE :- <SINGLE FILE>"
	Case LCase ("2")
		WScript.StdOut.WriteLine "OPERATING MODE :- <BULK FILE>"
	Case Else
		WScript.StdOut.WriteLine "OPERATING MODE :- <NOT SET!>"
End Select


End Sub


'###########################################################################
'Displays operating modes And takes input
Public Function SelectMode()
Dim StrMode

WScript.StdOut.WriteLine "SELECT OPERATING MODE? [Eg. Input 1 for 'Single File Mode']"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine "1. SINGLE FILE [Validate one XML against one XSD] ?"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine "2. BULK FILE [Validate multiple XML against one XSD] ?"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine "Tip: Type a bullet number from above and hit Enter."
WScript.StdOut.WriteBlankLines(1)

strMode = ConsoleInput()
SelectMode = strMode 

End Function

'###########################################################################

Sub ExitApp()
	 WScript.StdOut.WriteBlankLines(1)
	 WScript.StdOut.WriteLine "Press 'Enter' key to exit ..."
	 ConsoleInput()
	 WScript.Quit
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

Public Sub BulkFileValidation()
Dim strFilePath
Dim sFormatChoice

ShowWelcomeBox("valonly")
ShowFileChoices()
sFormatChoice = ConsoleInput()

If CreateWorkingDir() Then
	Call ConsoleOutput ("PROVIDE FULL PATH TO TEMPLATE XML FILE (.xml) ? ", "verbose", LogHandle)
	strFilePath = ConsoleInput()
	Set MyXMLFile = LoadXML(strFilePath)
	
	Call ConsoleOutput ("PROVIDE FULL PATH TO THE SCHEMA FILE (.xsd) ? ", "verbose", LogHandle)
	Set MyXSDFile = LoadXSD(ConsoleInput(),GetNamespaceURI(MyXMLFile))
	
	ValidateXML MyXMLFile, MyXSDFile
	Set MyXSDFile = Nothing
	
	'Call ConsoleOutput ("SPECIFY NUMBER OF TRANSACTIONS", "verbose")
	'Call GenerateFile (sFormatChoice, MyXMLFile, ConsoleInput)
	'Call SaveXML (MyXMLFile)
	Set MyXMLFile = Nothing
End If

End Sub

'###########################################################################
'This function calls Readme.txt

'Public Sub ShowReadMe()

'Set ObjFSO = CreateObject("Scripting.FileSystemObject")
'Set ObjTextFile = ObjFSO.OpenTextFile(GetCurrentDir() & "\ReadMe.txt", 1, False)
'Set ObjTextFile = Nothing
'Set ObjFSO = Nothing

'End Sub

'###########################################################################

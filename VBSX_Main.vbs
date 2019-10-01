'###################################################################################################
'# SCRIPT NAME: VBSX_Main.vbs
'#
'# DESCRIPTION:
'# Free script utility for silent XML/XSD validation of large sized files.
'# The VBSX_Validator is designed to validate large XML files.The project 
'# exposes the power and flexibility of VB Script language and demonstrates how it 
'# could be utilized for some specific XML related operations and automation.
'# 
'# NOTES:
'# Dependency on MSXML6. Supports full multiple error parsing with offline log file output.
'# Also supports Batch (Multiple XML Files) Validation against a single specified XSD
'# The Parser does not resolve externals. It does not evaluate or resolve the schemaLocation 
'# or attributes specified in DocumentRoot. The parser validates strictly against the 
'# supplied XSD only without auto-resolving schemaLocation. The parser needs 
'# Namespace (targetNamespace) which is currently extracted from the supplied XSD.

'# PLATFORM: Win7/8/Server | PRE-REQ: Script/Admin Privilege
'# LAST UPDATED: May 2019 | AUTHOR: Tushar Sharma
'##################################################################################################



If WScript.Arguments.length = 0 Then
   Set objShell = CreateObject("Shell.Application")
   objShell.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 3
      WScript.Quit
End If  

'###########################################################################

Dim LogHandle, strLogPath
Dim strSchemaCacheFail, strCurrentFileName, oDictOutputFolders

Const strInvalid = "invalid"
Const strFile = "file"
Const strFolder = "folder"
Const strFileExtXSD = ".xsd"
Const strFileExtXML = ".xml"

Const strMainOutFolder = "MainOutputDir"
Const strValidOutFolder = "ValidOutputDir"
Const strInvalidOutFolder = "InvalidOutputDir"


Call StartVBSXMain()


'###########################################################################

Sub StartVBSXMain()

	ShowWelcomeBox()
	ShowMode ("nomode")
	strOpsMode = SelectMode()
	ShowMode (strOpsMode)
	
	strCurrentFileName = ""
	strSchemaCacheFail = False
	
	If Not(IsObject(LogHandle)) Then
		Set LogHandle = CreateLogWriter()
	End If
	
	If ValidateInput(strOpsMode) Then
		Select Case strOpsMode
			Case "1"
				Call SingleFileValidation()
			Case "2"
				Call BulkFileValidation()
		End Select
	Else 
		ConsoleOutput "INVALID CHOICE!", "verbose", LogHandle
		If IsReloadExit("") Then
			Call StartVBSXMain()
		Else
			ExitApp()
		End If
	End If
	
	Call ExitApp()

End Sub


'###########################################################################

Public Sub SingleFileValidation()

Dim ObjSchemaCache, objXMLFile
Dim strFilePath
	
	ConsoleOutput "PROVIDE FULL PATH TO XML FILE (e.g. C:\MyFile.xml) ? ", "verbose", LogHandle
	strFilePath = ConsoleInput()
	
	If IsXMLXSD(strFilePath) = strFileExtXML Then
		Set objXMLFile = LoadXML(strFilePath)
		ConsoleOutput "", "verbose", LogHandle
		
		Set ObjSchemaCache = GetSchemaCacheForXSDs()
		
		If Not(strSchemaCacheFail) Then
			Call ValidateXML (objXMLFile, ObjSchemaCache, False)
		Else
			ConsoleOutput "", "verbose", LogHandle
			ConsoleOutput "<ERROR> NO SCHEMA FILE FOUND OR INVALID INPUT!", "verbose", LogHandle
			ConsoleOutput "", "verbose", LogHandle
			If IsReloadExit("") Then
				Call StartVBSXMain()
			Else
				ExitApp()
			End If
		End If
		
	Else
		ConsoleOutput "", "verbose", LogHandle
		ConsoleOutput "<ERROR> INVALID FILE OR PATH! PLEASE TRY AGAIN ...", "verbose", LogHandle
		If IsReloadExit("") Then
			Call StartVBSXMain()
		Else
			ExitApp()
		End If
	
	End If 

	ConsoleOutput "Log File : " & strLogPath, "verbose", LogHandle
	
	If IsReloadExit("") Then
		Call StartVBSXMain()
	Else
		ExitApp()
	End If
	
	Set objXMLFile = Nothing
	Set ObjSchemaCache = Nothing

End Sub	

'###########################################################################

Public Sub BulkFileValidation()

Dim ObjSchemaCache, objXMLFile
Dim strFilePath, strFolderPath, strFileName, arrFileList

Set ObjFSOTemp = CreateObject("Scripting.FileSystemObject")
Set oDictOutputFolders = GetOutputFolders()
Set ObjSchemaCache = GetSchemaCacheForXSDs()

If Not(strSchemaCacheFail) Then

	ConsoleOutput "PROVIDE PATH TO FOLDER CONTAINING XML FILES (e.g. C:\MyXMLFiles) ? ", "verbose", LogHandle
	strFolderPath = ConsoleInput()

	arrFileList = GetFolderFiles(strFolderPath, strFileExtXML)
	If IsArray(arrFileList) Then
		For Each strFileName In arrFileList
			strCurrentFileName = ObjFSOTemp.GetFileName(strFileName)
			Set objXMLFile = LoadXML(strFileName)
			Call ValidateXML (objXMLFile, ObjSchemaCache, True)
		Next
	Else
		ConsoleOutput "", "verbose", LogHandle
		ConsoleOutput "<ERROR> NO XML FILES FOUND OR INVALID INPUT!", "verbose", LogHandle
		ConsoleOutput "", "verbose", LogHandle
		If IsReloadExit("") Then
			Call StartVBSXMain()
		Else
			ExitApp()
		End If
	End If

Else

	ConsoleOutput "", "verbose", LogHandle
	ConsoleOutput "<ERROR> NO SCHEMA FILE FOUND OR INVALID INPUT!", "verbose", LogHandle
	ConsoleOutput "", "verbose", LogHandle
	If IsReloadExit("") Then
		Call StartVBSXMain()
	Else
		ExitApp()
	End If

End If
	
	If IsReloadExit("") Then
		Call StartVBSXMain()
	Else
		ExitApp()
	End If
	
	Set objXMLFile = Nothing
	Set ObjSchemaCache = Nothing

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
		Call ParseLoadErrors (ObjXML.parseError)
		If IsReloadExit("") Then
			Call StartVBSXMain()
		Else
			ExitApp()
		End If
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

Public Function LoadSchemaCache (objSchemaColl, objXSDFile)

Dim strNsURI
	
	ConsoleOutput "", "verbose", LogHandle
	ConsoleOutput "<INFO> Creating Schema Cache Collection", "verbose", LogHandle
	
	'Get targetNamespace property from XSD
	strNsURI = GetNamespaceURI (objXSDFile)
	
	'Load XSD from the Path
	objSchemaColl.Add strNsURI, objXSDFile
	ConsoleOutput "<INFO> Schema Cache Loaded Successfully from ... " & objXSDFile.url, "verbose", LogHandle
	
	Set LoadSchemaCache = objSchemaColl
	
End Function

'###########################################################################

Public Function GetNamespaceURI (ObjXML)

Dim strNsURI

strNsURI = ObjXML.documentElement.getAttribute("targetNamespace")

If strNsURI <> "" Then
	GetNamespaceURI = strNsURI
	ConsoleOutput "<INFO> Adding 'targetNamespace' " & strNsURI, "verbose", LogHandle
Else
	strNsURI = ObjXML.namespaceURI
	GetNamespaceURI = strNsURI
	ConsoleOutput "<INFO> Adding 'targetNamespace' " & strNsURI, "verbose", LogHandle
End If

End Function

'###########################################################################

Public Function ValidateXML (ObjXMLDoc, ObjXSDDoc, bIsSaveFile)

Dim bValResult, oDictFolders
Dim strValidDirPath, strInvalidDirPath

Set ObjXMLDoc.Schemas = ObjXSDDoc

ConsoleOutput "", "verbose", LogHandle
ConsoleOutput vbTab & "******************" & vbTab & "<STARTING VALIDATION> " & vbTab & " ******************" & vbCrLf , "verbose", LogHandle

If ObjXMLDoc.readystate = 4 Then
	
	Set ObjXParseErr = ObjXMLDoc.validate()
	bValResult = ParseValidationError (ObjXParseErr, ObjXMLDoc)
	
	If (bIsSaveFile) Then
		strValidDirPath = oDictOutputFolders.Item(strValidOutFolder)
		strInvalidDirPath = oDictOutputFolders.Item(strInvalidOutFolder)
		Select Case bValResult
			Case True
				ObjXMLDoc.Save(strValidDirPath & "\" & strCurrentFileName)
			Case False
				ObjXMLDoc.Save(strInvalidDirPath & "\" & strCurrentFileName)
		End Select
	
	End If
	
End If

ConsoleOutput "", "verbose", LogHandle
ConsoleOutput vbTab & "******************" & vbTab & "<COMPLETED VALIDATION> " & vbTab & " ******************" & vbCrLf , "verbose", LogHandle
ConsoleOutput "", "verbose", LogHandle

End Function

'###########################################################################

' This ParseError property is for errors and warnings during 'Load'method.
' Applies to IXMLDOMParseError interface.

Public Function ParseLoadErrors (ByVal ObjParseErr)

Dim strResult

strResult = vbCrLf & "<ERROR> INVALID XML! FAILED WELL-FORMED (STRUCTURE) CHECK ! " & _
vbCrLf & ObjParseErr.reason & vbCr & _
"Error Code: " & ObjParseErr.errorCode & ", Line: " & _
				 ObjParseErr.Line & ", Character: " & _		
				 ObjParseErr.linepos & ", Source: " & _
				 Chr(34) & ObjParseErr.srcText & _
				 Chr(34) & " - " & vbCrLf & _
				 vbCrLf & "CORRECT THE FILE BEFORE CONTINUING XSD VALIDATION !" & vbCrLf 

ConsoleOutput "Log File : " & strLogPath, "verbose", LogHandle
ConsoleOutput "", "verbose", LogHandle
ConsoleOutput strResult, "verbose", LogHandle
ParseLoadErrors = False

End Function

'###########################################################################

' The .AllErrors contains errors and warnings found DURING validation. Not valid for Load errors.
' Applies to IXMLDOMParseError2 interface which extends the IXMLDOMParseError interface

Public Function ParseValidationError (ByVal ObjParseErr, ObjXMLDoc)
Dim strResult, ErrFound
ErrFound = 0

Select Case ObjParseErr.errorCode
	Case 0
		ConsoleOutput "", "verbose", LogHandle
		strResult = "<INFO> XML SCHEMA VALIDATION: SUCCESS ! " & vbCrLf & ObjXMLDoc.url & vbCrLf '& ObjXSDDoc.url
		ConsoleOutput strResult, "verbose", LogHandle
		ParseError = True
	Case Else
	   If (ObjParseErr.AllErrors.length > 1) Then	'.AllErrors contains errors and warnings found DURING validation. Not valid property for Load errors
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
	ParseValidationError = False
Else 
	ParseValidationError = True
End If

End Function

'###########################################################################


'This function takes input form user
Public Function ConsoleInput()
Dim strIn

strIn = WScript.StdIn.ReadLine

If (Right(strIn,1) = Chr(34)) And (Left(strIn,1) = Chr(34)) Then
	strIn = Replace(strIn,Chr(34),"")
End If

ConsoleInput = strIn

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


Function GetSchemaCacheForXSDs()

Dim strInput, iFound, strFileName, iInvalidCount
Dim ObjSchemaCache , objXSDFile, arrXSDFiles

Set ObjSchemaCache = CreateObject("MSXML2.XMLSchemaCache.6.0")
'Indicates whether the schema will be compiled and validated when it is loaded into the schema cache
ObjSchemaCache.validateOnload = False ' This method applies to only [Schema Cache] not (XSD or XML)

ConsoleOutput "PROVIDE FULL PATH TO SCHEMA FILE/S OR FOLDER CONTAINING XSDs (e.g. C:\MySchema.xsd OR C:\SchemaFileFolder) ? ", "verbose", LogHandle
strInput = ConsoleInput()

iFound = 0
iInvalidCount = 0
Do	
	Select Case IsFolderFile(strInput)
		Case strFile
			If (IsXMLXSD(strInput) = strFileExtXSD) Then
				Set objXSDFile = LoadXML(strInput)
				Set ObjSchemaCache = LoadSchemaCache(ObjSchemaCache, objXSDFile)
				iFound = iFound + 1
			Else 
				strInput = strInvalid
				iInvalidCount = iInvalidCount + 1
				ConsoleOutput "", "verbose", LogHandle
				ConsoleOutput "<ERROR> INVALID INPUT! SUPPLIED FILE IS NOT XSD ('.xsd'). PLEASE TRY AGAIN ...", "verbose", LogHandle
			End If
		Case strFolder
			arrXSDFiles = GetFolderFiles(strInput, strFileExtXSD)
			If IsArray(arrXSDFiles) Then
				For Each strFileName In arrXSDFiles
					Set objXSDFile = LoadXML(strFileName)
					Set ObjSchemaCache = LoadSchemaCache(ObjSchemaCache, objXSDFile)
					iFound = iFound + 1
				Next
			End If
		Case strInvalid
			strInput = strInvalid
			iInvalidCount = iInvalidCount + 1
			ConsoleOutput "", "verbose", LogHandle
			ConsoleOutput "<ERROR> INVALID INPUT! PLEASE TRY AGAIN ...", "verbose", LogHandle
	End Select
	
	If (iInvalidCount >= 3) Then
		ConsoleOutput "", "verbose", LogHandle
		ConsoleOutput "<ERROR> TOO MANY INVALID ATTEMPTS! EXIT OR RE-LOAD APPLICATION ...", "verbose", LogHandle
		strInput = ""
		Exit Do
	Else 
		ConsoleOutput "", "verbose", LogHandle
		ConsoleOutput "PROVIDE MORE XSDs IF ASSOCIATED WITH XML ELSE JUST PRESS ENTER (e.g. C:\MySchema.xsd OR C:\SchemaFileFolder) ? ", "verbose", LogHandle
		strInput = ConsoleInput()
	End If
	
Loop While (strInput <> "")


If (iFound > 0) And Not(iInvalidCount >= 3) Then	
	Set GetSchemaCacheForXSDs = ObjSchemaCache
	strSchemaCacheFail = False
Else
	Set GetSchemaCacheForXSDs = ObjSchemaCache
	strSchemaCacheFail = True
End If


End Function

'###########################################################################

Function IsFolderFile(strPathInput)

Set objFSO = CreateObject("Scripting.FileSystemObject") 

If objFSO.FileExists(strPathInput) Then 
    IsFolderFile = strFile
ElseIf objFSO.FolderExists(strPathInput) Then
	IsFolderFile = strFolder
Else 
	IsFolderFile = strInvalid
End If

End Function

'###########################################################################

Function IsXMLXSD(strFilePath)

Dim objFSO, strFileExt

Set objFSO = CreateObject("Scripting.FileSystemObject") 

If IsFolderFile(strFilePath) = strFile Then
	strFileExt = objFSO.GetFileName(strFilePath)
	
	Select Case Right(strFileExt,4)
		Case ".xml"
			IsXMLXSD = strFileExtXML	
		Case ".xsd"
			IsXMLXSD = strFileExtXSD
		Case Else
			IsXMLXSD = strInvalid
	End Select

Else
	IsXMLXSD = strInvalid
End If

End Function

'###########################################################################
'This function sets the directory path of ReadMe.txt 

Public Function CreateLogWriter()

sCurrPath = Left(WScript.ScriptFullName,(Len(WScript.ScriptFullName)) - (Len(WScript.ScriptName)))
strFileName = "vbsx-Validator" & "_" & Day(Date) & MonthName(Month(Date),True) & Right((Year(Date)),2) & ".txt"

Set ObjFSO = CreateObject("Scripting.FileSystemObject")
Set ObjTextFile = ObjFSO.OpenTextFile(sCurrPath & strFileName, 8, True)
strLogPath = sCurrPath & strFileName

Set CreateLogWriter = ObjTextFile 

Set ObjTextFile = Nothing
Set ObjFSO = Nothing

'The Other Methods -
'sCurrPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")

End Function


'###########################################################################


Function GetOutputFolders ()

sCurrPath = Left(WScript.ScriptFullName,(Len(WScript.ScriptFullName)) - (Len(WScript.ScriptName)))
strMainFolderName = "Output" & "_" & Day(Date) & MonthName(Month(Date),True) & Right((Year(Date)),2)
strSubFolderName = Hour(Now) & Minute(Now) & Second(Now)

Set ObjFSO = CreateObject("Scripting.FileSystemObject")
Set ObjFolderDict = CreateObject("Scripting.Dictionary")

If Not (ObjFSO.FolderExists(sCurrPath & strMainFolderName)) Then
	Set ObjOutputDir = ObjFSO.CreateFolder(sCurrPath & strMainFolderName)
	Set ObjValidDir = ObjFSO.CreateFolder(sCurrPath & strMainFolderName & "\" & "Valid" & "_" & strSubFolderName)
	Set ObjInvalidDir = ObjFSO.CreateFolder(sCurrPath & strMainFolderName & "\" & "Invalid" & "_" & strSubFolderName)
Else
	Set ObjOutputDir = ObjFSO.GetFolder(sCurrPath & strMainFolderName)
	Set ObjValidDir = ObjFSO.CreateFolder(sCurrPath & strMainFolderName & "\" & "Valid" & "_" & strSubFolderName)
	Set ObjInvalidDir = ObjFSO.CreateFolder(sCurrPath & strMainFolderName & "\" & "Invalid" & "_" & strSubFolderName)
End if

ObjFolderDict.Add strMainOutFolder,ObjOutputDir.Path
ObjFolderDict.Add strValidOutFolder,ObjValidDir.Path
ObjFolderDict.Add strInvalidOutFolder,ObjInvalidDir.Path

Set GetOutputFolders = ObjFolderDict

End Function


'###########################################################################

Function ValidateInput (strArgsIn)

Dim strValidInput, strArg, strFound
strFound = False
strValidNumIn = Array("1","2")
strValidStrIn = Array("Y","N","YES","NO")

If IsNumeric(strArgsIn) Then
	For Each strArg In strValidNumIn
		If (StrComp(strArg, strArgsIn) = 0) Then
			strFound = True
			Exit For
		End If
	Next
Else
	For Each strArg In strValidStrIn
		If (StrComp(UCase(strArg), strArgsIn) = 0) Then
			strFound = True
			Exit For
		End If
	Next
End If
	
	
	If Not(strFound) Then
		ValidateInput = False
	Else 
		ValidateInput = True
	End If

End Function 

'###########################################################################

Function IsReloadExit (ObjXML)
IsWait = True

If IsObject(ObjXML) Then
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
End If

ConsoleOutput "", "nolog", LogHandle
ConsoleOutput "RE-LOAD THE PROGRAM OR EXIT (y=Reload / n=Exit) ?", "nolog", LogHandle
strResponse = UCase(ConsoleInput())

If ValidateInput(strResponse) Then
	Select Case strResponse
	    Case "Y"
	    	IsReloadExit = True
	    Case "N"
	    	IsReloadExit = False
	End Select
Else
	ConsoleOutput "INVALID CHOICE!", "verbose", LogHandle
End If

If Not(IsWait) Then
	Call ExitApp()
End If

End Function

'###########################################################################

Function GetFolderFiles(strFolderPath,strFileExtType)

Dim ObjFSO, ObjFolder, ObjFiles, strCurPath
Dim iCount, arrObjFiles(), strFile

iCount = 0

Set ObjFSO = CreateObject("Scripting.FileSystemObject")

If (ObjFSO.FolderExists(strFolderPath)) Then
	Set ObjFolder = ObjFSO.GetFolder(strFolderPath)
	strCurPath = ObjFolder.Path
	Set ObjFiles = ObjFolder.Files
	If ObjFiles.Count > 0 Then
		ConsoleOutput "", "verbose", LogHandle
		ConsoleOutput "<INFO> LOADING FILES FROM FOLDER " & strCurPath, "verbose", LogHandle
		For Each strFile In ObjFiles
			If (Right(strFile.Path,4) = strFileExtType) Then
				ConsoleOutput "", "verbose", LogHandle
				ConsoleOutput "<INFO> Found File : " & strFile.Path, "verbose", LogHandle
				ReDim Preserve arrObjFiles (iCount)
				arrObjFiles(iCount) = strFile.Path
				iCount = iCount + 1
			End If
		Next
		If IsArray(arrObjFiles) Then
			GetFolderFiles = arrObjFiles
		Else
			GetFolderFiles = False
		End If	
	Else 
		ConsoleOutput "<ERROR> NO FILES FOUND IN THE SPECIFIED FOLDER !", "verbose", LogHandle
		GetFolderFiles = False
		If IsReloadExit("") Then
			Call StartVBSXMain()
		Else
			ExitApp()
		End If		
	End If
Else
	ConsoleOutput "<ERROR> FOLDER NOT FOUND !", "nolog", LogHandle
	GetFolderFiles = False
	If IsReloadExit("") Then
		Call StartVBSXMain()
	Else
		ExitApp()
	End If

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
WScript.StdOut.WriteLine "1. SINGLE FILE [Single XML against XSD/s] ?"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine "2. BULK FILE [Multiple XMLs against XSD/s] ?"
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

Public Sub ShowWelcomeBox()

WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine "      " & "****************************************************************"
WScript.StdOut.WriteLine "      " & "----------------------------------------------------------------"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine VBTab & vbTab & VBTab & "   " & "VBSX_VALIDATOR version v2.0.1"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine VBTab & " Free and Fast XML/XSD Validator. Supports Bulk Validation"
WScript.StdOut.WriteLine vbTab & "  " & "Full XML Error Parsing (MSXML6) with Offline Log Output"
WScript.StdOut.WriteLine VBTab & "    " & "Platform: Win7/8 | Pre-Req: Script/Admin Privilege"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine VBTab & "   " & "Updated: May 2019 | Tushar Sharma | www.testoxide.com"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine "      " & "****************************************************************"
WScript.StdOut.WriteLine "      " & "----------------------------------------------------------------"
WScript.StdOut.WriteBlankLines(2)

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

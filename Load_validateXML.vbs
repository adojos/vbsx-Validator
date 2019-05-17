Option Explicit
Dim input
Dim inputxsdns
Dim xml 'to store path of xml file
Dim xsd ' To store path of xsd file
' Call ChooseFile to select XML file from local drive
WScript.Echo "Selected xml  file: " & ChooseFile( )

' Call ChooseFile1 to select XSD file from local drive

WScript.Echo "Selected xsd file: " & ChooseFile1( )

'Launch Input box to accept namespace of XSD file
inputxsdns = InputBox("NameSpace of XSD file entered : ", "Enter XSD namespace URI")

'Verify that XML file path is correct
If Not FileExists(xml) Then
 Msgbox "Specified XML file does not exists",vbCritical
 WScript.Quit
End If

'Verify that XSD file path is correct
If Not FileExists(xsd) Then
 Msgbox "Specified XSD file does not exists",vbCritical
 WScript.Quit
End If
'Call LoadXMLDocument function
Call LoadXMLDocument(xml,xsd, inputxsdns)

'******************************
'Sub Routines / Functions
'******************************
' Function to choose xml file from local drive
Function ChooseFile( )
 Dim objFSO, objShell, objTempFolder, strTempFileName, strFullTempFileName, objOpenFile, objTextFile, strTempTextFileName
 Const TemporaryFolder = 2
 Const ForReading = 1
 strTempFileName = "OpenFile.hta"
 strTempTextFileName = "OpenFile.txt"
 Set objFSO= CreateObject("Scripting.FileSystemObject")
 Set objTempFolder = objFSO.GetSpecialFolder(TemporaryFolder)
 strFullTempFileName=objTempFolder.Path & "\" & strTempFileName
 Set objOpenFile = objFSO.CreateTextFile(strFullTempFileName,True)
 objOpenFile.writeline("<html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1252"">")
 objOpenFile.writeline("<title>Open File</title>")
 objOpenFile.writeline("<script language=""vbscript"">")
 objOpenFile.writeline("Sub Window_Onload")
 objOpenFile.writeline("FileName.click")
 objOpenFile.writeline("WriteFile FileName.value")
 objOpenFile.writeline("Self.Close()")
 objOpenFile.writeline("End Sub")
 objOpenFile.writeline("Sub WriteFile(strFileName)")
 objOpenFile.writeline("Dim objFSO, objTempFolder, strTempFileName, strFullTempFileName, objOpenFile")
 objOpenFile.writeline("Const TemporaryFolder = 2")
 objOpenFile.writeline("strTempFileName = ""OpenFile.txt""")
 objOpenFile.writeline("Set objFSO=CreateObject(""Scripting.FileSystemObject"")")
 objOpenFile.writeline("Set objTempFolder = objFSO.GetSpecialFolder(TemporaryFolder)")
 objOpenFile.writeline("strFullTempFileName=objTempFolder.Path & ""\"" & strTempFileName")
 objOpenFile.writeline("Set objOpenFile = objFSO.CreateTextFile(strFullTempFileName,True)")
 objOpenFile.writeline("objOpenFile.writeline(strFileName)")
 objOpenFile.writeline("objOpenFile.Close")
 objOpenFile.writeline("Set objFSO=Nothing")
 objOpenFile.writeline("Set objTempFolder=Nothing")
 objOpenFile.writeline("Set objSleepFile=Nothing")
 objOpenFile.writeline("Set objShell=Nothing")
 objOpenFile.writeline("End Sub")
 objOpenFile.writeline("</script>")
 objOpenFile.writeline("<hta:application applicationname=""Open File"" border=""dialog"" borderstyle=""normal"" caption=""Open File"" contextmenu=""no"" maximizebutton=""no"" minimizebutton=""no"" navigable=""no"" scroll=""no"" selection=""no"" showintaskbar=""no"" singleinstance=""yes"" sysmenu=""no"" version=""1.0"" windowstate=""minimize"">")
 objOpenFile.writeline("</head>")
 objOpenFile.writeline("<body>")
 objOpenFile.writeline("<input Application=""True"" type=""file"" id=""FileName"" />")
 objOpenFile.writeline("</body>")
 objOpenFile.writeline("</html>")
 objOpenFile.Close
 Set objShell = CreateObject("WScript.Shell")
 objShell.Run "mshta.exe " & strFullTempFileName,0,True
 objFSO.DeleteFile strFullTempFileName, True
 Set objShell=Nothing
 Set objOpenFile=Nothing
 strFullTempFileName = objTempFolder.Path & "\" & strTempTextFileName
 
 Set objTextFile=objFSO.OpenTextFile(strFullTempFileName, ForReading)
 ChooseFile = objTextFile.ReadLine 
 xml= ChooseFile
 objTextFile.Close
 objFSO.DeleteFile strFullTempFileName, True
 Set objTextFile=Nothing
 Set objFSO=Nothing
 Set objTempFolder=Nothing
End Function

' Function to choose xsd file from local drive
Function ChooseFile1( )
 Dim objFSO1, objShell1, objTempFolder1, strTempFileName1, strFullTempFileName1, objOpenFile1, objTextFile1, strTempTextFileName1
 Const TemporaryFolder1 = 2
 Const ForReading1 = 1
 strTempFileName1 = "OpenFile.hta"
 strTempTextFileName1 = "OpenFile.txt"
 Set objFSO1= CreateObject("Scripting.FileSystemObject")
 Set objTempFolder1 = objFSO1.GetSpecialFolder(TemporaryFolder1)
 strFullTempFileName1=objTempFolder1.Path & "\" & strTempFileName1
 Set objOpenFile1 = objFSO1.CreateTextFile(strFullTempFileName1,True)
 objOpenFile1.writeline("<html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1252"">")
 objOpenFile1.writeline("<title>Open File</title>")
 objOpenFile1.writeline("<script language=""vbscript"">")
 objOpenFile1.writeline("Sub Window_Onload")
 objOpenFile1.writeline("FileName.click")
 objOpenFile1.writeline("WriteFile FileName.value")
 objOpenFile1.writeline("Self.Close()")
 objOpenFile1.writeline("End Sub")
 objOpenFile1.writeline("Sub WriteFile(strFileName)")
 objOpenFile1.writeline("Const TemporaryFolder = 2")
 objOpenFile1.writeline("strTempFileName = ""OpenFile.txt""")
 objOpenFile1.writeline("Set objFSO=CreateObject(""Scripting.FileSystemObject"")")
 objOpenFile1.writeline("Set objTempFolder = objFSO.GetSpecialFolder(TemporaryFolder)")
 objOpenFile1.writeline("strFullTempFileName=objTempFolder.Path & ""\"" & strTempFileName")
 objOpenFile1.writeline("Set objOpenFile = objFSO.CreateTextFile(strFullTempFileName,True)")
 objOpenFile1.writeline("objOpenFile.writeline(strFileName)")
 objOpenFile1.writeline("objOpenFile.Close")
 objOpenFile1.writeline("Set objFSO=Nothing")
 objOpenFile1.writeline("Set objTempFolder=Nothing")
 objOpenFile1.writeline("Set objSleepFile=Nothing")
 objOpenFile1.writeline("Set objShell=Nothing")
 objOpenFile1.writeline("End Sub")
 objOpenFile1.writeline("</script>")
 objOpenFile1.writeline("<hta:application applicationname=""Open File"" border=""dialog"" borderstyle=""normal"" caption=""Open File"" contextmenu=""no"" maximizebutton=""no"" minimizebutton=""no"" navigable=""no"" scroll=""no"" selection=""no"" showintaskbar=""no"" singleinstance=""yes"" sysmenu=""no"" version=""1.0"" windowstate=""minimize"">")
 objOpenFile1.writeline("</head>")
 objOpenFile1.writeline("<body>")
 objOpenFile1.writeline("<input Application=""True"" type=""file"" id=""FileName"" />")
 objOpenFile1.writeline("</body>")
 objOpenFile1.writeline("</html>")
 objOpenFile1.Close
 Set objShell1 = CreateObject("WScript.Shell")
 objShell1.Run "mshta.exe " & strFullTempFileName1,0,True
 objFSO1.DeleteFile strFullTempFileName1, True
 Set objShell1=Nothing
 Set objOpenFile1=Nothing
 strFullTempFileName1 = objTempFolder1.Path & "\" & strTempTextFileName1
 Set objTextFile1=objFSO1.OpenTextFile(strFullTempFileName1, ForReading1)
 ChooseFile1 = objTextFile1.ReadLine 
 xsd= ChooseFile1
 objTextFile1.Close
 objFSO1.DeleteFile strFullTempFileName1, True
 Set objTextFile1=Nothing
 Set objFSO1=Nothing
 Set objTempFolder1=Nothing
End Function

' Function to verify file exist or not

Function FileExists(filename)
 Dim objFSO2
 
 'Create Object for FileSystem
 Set objFSO2 = CreateObject("Scripting.FileSystemObject")
 
 'Verify the specified file exists
 If objFSO2.FileExists(filename) Then
  FileExists = True
 Else 
  FileExists = False
 End If
 
 Set objFSO2 = Nothing
End Function

' Function to load and validate xml against XSD

Function LoadXMLDocument(fileName, xsd, xsdns)
 Dim objDOM, objSchema, loadStatus
Dim xmlParseErr, objXMLDocSchemaCache, xmlXSDError
 
 'Create Object for MSXML
 Set objDOM = CreateObject("MSXML2.DOMDocument.6.0") 
 Set objXMLDocSchemaCache = CreateObject("Msxml2.XMLSchemaCache.6.0")
 
 'Add the XSD file to SchemCache object
 objXMLDocSchemaCache.add xsdns, xsd 
 
 'Set the DOM variables
 objDOM.async = False
 objDOM.validateOnParse = True
 objDOM.resolveExternals = True
 
 'Set the loaded XSD to DOM and load XML to DOM
 Set objDOM.schemas = objXMLDocSchemaCache 
 loadStatus = objDOM.Load(fileName) 
 Set xmlParseErr = objDOM.parseError
 
 'verify that XML is validated aganist Schema and loaded successfully
 If xmlParseErr.errorCode <> 0 Then
  Msgbox  xmlParseErr.reason, vbCritical
 Else  
  Set xmlXSDError = objDOM.validate
  If xmlXSDError.errorCode <> 0 Then
   MsgBox xmlXSDError.reason, vbCritical
  Else
   Msgbox "Successfully validated the XML aganist XSD", vbInformation 
  End If
 End If
 
 Set objDOM = Nothing
 Set objXMLDocSchemaCache = Nothing
End Function
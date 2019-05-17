Option Explicit
Dim input
Dim inputxsdns
'Launch Input box
'InputBox ( "title", "prompt" [, "default" [, "password char" [, width [, height [, left [, top [, timeout [, hwnd]]]]]]]] )
input = InputBox("Enter values with comma separated ", "Enter XML and XSD")
'input = InputBox("Question", "Where were you born?")
input = Split(input,",")

'Launch Input box
inputxsdns = InputBox("NameSpace of XSD file entered : ", "Enter XSD namespace URI")

'Verify that XML file path is correct
If Not FileExists(Trim(input(0))) Then
 Msgbox "Specified XML file does not exists",vbCritical
 WScript.Quit
End If

'Verify that XSD file path is correct
If Not FileExists(Trim(input(1))) Then
 Msgbox "Specified XSD file does not exists",vbCritical
 WScript.Quit
End If
'Call LoadXMLDocument function
Call LoadXMLDocument(Trim(input(0)), Trim(input(1)), inputxsdns)

'******************************
'Sub Routines / Functions
'******************************

Function FileExists(filename)
 Dim objFSO
 
 'Create Object for FileSystem
 Set objFSO = CreateObject("Scripting.FileSystemObject")
 
 'Verify the specified file exists
 If objFSO.FileExists(filename) Then
  FileExists = True
 Else 
  FileExists = False
 End If
 
 Set objFSO = Nothing
End Function

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
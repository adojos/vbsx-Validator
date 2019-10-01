
Const strMainOutFolder = "MainOutputDir"
Const strValidOutFolder = "ValidOutputDir"
Const strInvalidOutFolder = "InvalidOutputDir"

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

MsgBox "Main Folder " & ObjFolderDict.Item(strMainOutFolder)
MsgBox "Main Folder " & ObjFolderDict.Item(strValidOutFolder)
MsgBox "Main Folder " & ObjFolderDict.Item(strInvalidOutFolder)

Dim ObjFSO, ObjFolder, ObjFiles, strCurPath
Dim iCount, arrObjFiles()

strFolderPath = "E:\PROGRAMMING\XMLs\XMLDOM-PROJECTS\VBSX-VALIDATOR\examples\Set3_Multi_XSD_Demo\ipo"
strFileExtXSD = ".xsd"

iCount = 0

Set ObjFSO = CreateObject("Scripting.FileSystemObject")

If (ObjFSO.FolderExists(strFolderPath)) Then
	Set ObjFolder = ObjFSO.GetFolder(strFolderPath)
	strCurPath = ObjFolder.Path
	MsgBox strCurPath
	Set ObjFiles = ObjFolder.Files
	If ObjFiles.Count > 0 Then
'		ConsoleOutput "", "verbose", LogHandle
'		ConsoleOutput "<INFO> Loading Files from folder " & strCurPath, "verbose", LogHandle
		For Each strFile In ObjFiles
		MsgBox strFile.Path
			If (Right(strFile.Path,4) = strFileExtType) Then
'				ConsoleOutput "<INFO> Found File " & strFile.Path, "verbose", LogHandle
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
'		ConsoleOutput "<ERROR> NO FILES FOUND IN THE SPECIFIED FOLDER !", "verbose", LogHandle
'		GetFolderFiles = False
'		If IsReloadExit("") Then
'			Call StartVBSXMain()
'		Else
'			ExitApp()
'		End If		
	End If
Else
'	ConsoleOutput "<ERROR> FOLDER NOT FOUND !", "nolog", LogHandle
'	GetFolderFiles = False
'	If IsReloadExit("") Then
'		Call StartVBSXMain()
'	Else
'		ExitApp()
'	End If

End If

	
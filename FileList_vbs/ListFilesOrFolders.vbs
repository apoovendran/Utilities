'------------------------------------------------------------------------------------------
'Name	:	ListFilesOrfolders_0.1.vbs
'Desc	:	List files and folders in csv
'Remarks:	ListFilesOrFolders - is the entry point
'			Parameter 1 - Path to list the files or folders eg : C:\tmp
'			Parameter 2	- If only folder List give "Folder"
'						- If Folders and Files -  then "Files"
'Output	:	It will return Output.csv file in the same folder
'------------------------------------------------------------------------------------------


Dim fso,RecCnt
Dim ObjOutFile

Set fso = CreateObject("Scripting.FileSystemObject") 
Set ObjOutFile = fso.CreateTextFile("E_2TB.csv") 
RecCnt = 0
ObjOutFile.WriteLine("RecCnt,Folder,File,Size,CreatedOn,DateLastModified,DateLastAccessed,FileCnt") 
RecCnt = RecCnt+1
ObjOutFile.WriteLine(RecCnt & ",E:\2TB") 
Call ListFilesOrFolders("E:\2TB", "Folder")


Function ListFilesOrFolders(strTopFolder, FolderOrFiles)
	Dim ObjFolder, bFirstTimeEntry
	Set ObjFolder = fso.GetFolder(strTopFolder) 
	
	Dim objFiles
	Set objFiles = objFolder.Files  
	Dim ObjSubFolders
	Set ObjSubFolders = ObjFolder.SubFolders 
	
	Dim ofldr
	For Each ofldr in ObjSubFolders
		Set objFiles = ofldr.Files  
		RecCnt = RecCnt+1
		ObjOutFile.WriteLine(RecCnt & "," & ofldr.path & ",Folder," &  ofldr.Size & "," & ofldr.DateCreated & "," & ofldr.DateLastModified & "," & ofldr.DateLastAccessed & "," & objFiles.Count)
		If FolderOrFiles = "Files" then
			Call ListFiles(ofldr)
		end if
		call ListFilesOrFolders(ofldr, FolderOrFiles)
	Next
End Function

Function ListFiles(objFolder)
	Dim objFiles
	Set objFiles = objFolder.Files  
	For each folderIdx In objFiles  
		RecCnt = RecCnt+1
		ObjOutFile.WriteLine(RecCnt & "," & folderIdx.path & ",File," &  folderIdx.Size & "," & folderIdx.DateCreated & "," & folderIdx.DateLastModified & "," & folderIdx.DateLastAccessed)
	Next

End Function

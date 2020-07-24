'------------------------------------------------------------------------------------------
'Name	:	ListFilesOrfolders.vbs
'Desc	:	List files and folders in txt - open in excel "/" delimited 
'Date	:	01 Feb 2014
'Remarks:	strPath and FileOrFolder to be given

	'------------------------------------------------------------------------------------------
'Version:	1.1
'			a) .svn folders removed from listing
'			b) comma used in folders so delimiter changed to "/"
'			c) File extension added in FileCnt column
	'------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------

Dim strPath
Dim FileOrFolder

strPath = "C:\Poov\Cloud\"
FileOrFolder = "Files"
' FileOrFolder = "Folder"

'------------------------------------------------------------------------------------------

Dim OutputFileName
Dim fso,RecCnt
Dim ObjOutFile

OutputFileName = Replace(strPath & FileOrFolder ,":","_")
OutputFileName = Replace(OutputFileName ,"\","_")
OutputFileName = OutputFileName & ".txt"

on error resume next

Set fso = CreateObject("Scripting.FileSystemObject") 
Set ObjOutFile = fso.CreateTextFile(OutputFileName) 
RecCnt = 0
ObjOutFile.WriteLine("RecCnt~Folder~File~Size~CreatedOn~DateLastModified~DateLastAccessed~FileCnt") 
RecCnt = RecCnt+1
ObjOutFile.WriteLine(RecCnt & "~" & strPath) 

Call ListFilesOrFolders(strPath, FileOrFolder)

Function ListFilesOrFolders(strTopFolder, FolderOrFiles)

on error resume next
	Dim ObjFolder, bFirstTimeEntry
	Set ObjFolder = fso.GetFolder(strTopFolder) 
	
	Dim objFiles
	Set objFiles = objFolder.Files  
	Dim ObjSubFolders
	Set ObjSubFolders = ObjFolder.SubFolders 
	
	Dim ofldr
	For Each ofldr in ObjSubFolders
		if InStr(1,ofldr,".svn") > 0  then
			'do nothing
		else
	
			Set objFiles = ofldr.Files  
			RecCnt = RecCnt+1
			ObjOutFile.WriteLine(RecCnt & "~" & ofldr.path & "~Folder~" &  ofldr.Size & "~" & ofldr.DateCreated & "~" & ofldr.DateLastModified & "~" & ofldr.DateLastAccessed & "~" & objFiles.Count)
			If FolderOrFiles = "Files" then
				Call ListFiles(ofldr)
			end if
			call ListFilesOrFolders(ofldr, FolderOrFiles)
		end if
	Next
End Function

Function ListFiles(objFolder)
on error resume next
	Dim objFiles
	Set objFiles = objFolder.Files  
	For each folderIdx In objFiles  
		RecCnt = RecCnt+1
		ObjOutFile.WriteLine(RecCnt & "~" & folderIdx.path & "~File~" &  folderIdx.Size & "~" & folderIdx.DateCreated & "~" & folderIdx.DateLastModified & "~" & folderIdx.DateLastAccessed & "~" &  fso.GetExtensionName(folderIdx.path))
	Next

End Function

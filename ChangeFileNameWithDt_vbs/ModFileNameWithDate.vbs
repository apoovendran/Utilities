' REQUIREMENT: - Rename  all / extension wise files with date suffix - folders and subfolders
'
'USER MANUAL:-
'		Initialize Inputs
'
' TECHNICAL DOCUMENT:
' 1. Initialize Inputs
' 2. Rename the parent Folder's file
' 3. Loop subfolders and call function to rename the files



'''''''''''''''''''''''''''''''''''''''''
' <1. Initialize Inputs starts 
Dim strFolder, strExtensions, strKeepNameAsPrefix, IsDebugLog, bResult
strFolder = "C:\APTP\"			'update
strExtensions = "All" 			' "All" OR "jpg,png,gif"	
strKeepNameAsPrefix = "Yes"
IsDebugLog = "Yes"				'either "Yes" or "No"
' 1. Initialize Inputs starts>

'Date and Log file creation 
Dim fso,iRowCnt
Dim ObjOutFile
Dim OutputFileName
Dim ObjFolder
	
Set fso = CreateObject("Scripting.FileSystemObject") 
iRowCnt = 0

'--- set output file
OutputFileName = Replace(strFolder,":","_")
OutputFileName = Replace(OutputFileName ,"\","_")
OutputFileName = "ChangeFileNameLog_" & OutputFileName & "_" & GetDateStr() & ".txt"
Set ObjOutFile = fso.CreateTextFile(OutputFileName) 
'objOutFile.WriteLine(GetDateStr() & " started")
objOutFile.WriteLine("RowNo~Date~OldName~NewName~Remarks")
objOutFile.WriteLine(iRowCnt & "~" & GetDateStr()& "~" & strFolder & "~strExtensions=" & strExtensions & "-strKeepNameAsPrefix=" & strKeepNameAsPrefix & "~Started")
	
	
Set ObjFolder = fso.GetFolder(strFolder) 
Call ListFiles(objFolder)
Call ListFolder(objFolder)

Function ListFolder(objFolder)
	Dim ofldr
	Dim ObjSubFolders
	Set ObjSubFolders = objFolder.SubFolders 
	
	'Call ListFiles(objFolder)
	
	For Each ofldr in ObjSubFolders
		objOutFile.WriteLine(iRowCnt & "~" & GetDateStr() & "~" & ofldr.name & "~Folder" & "~")
		Call ListFiles(ofldr)
		call ListFolder(ofldr)
	Next
	

End Function



Function ListFiles(objFolder)
	Dim objFiles
	Set objFiles = objFolder.Files  
	'RecCnt =RecCnt+1
	For each folderIdx In objFiles 
		iSubFldrRowCnt = iSubFldrRowCnt+1
		iRowCnt = iRowCnt+1
		If strExtensions = "All" then
			if strKeepNameAsPrefix = "Yes" then 
					objOutFile.WriteLine(iRowCnt & "~" & GetDateStr() & "~" & folderIdx.name & "~" & folderIdx.name & "_" & GetDateStr() & "-" & iSubFldrRowCnt & "~Renamed")
				fso.movefile folderIdx.path, objFolder.path & "\" & folderIdx.name & "_" & GetDateStr() & "-" & iSubFldrRowCnt
			Else
						objOutFile.WriteLine(iRowCnt & "~" & GetDateStr() & "~" & folderIdx.name & "~" & folderIdx.name & "_" & GetDateStr() & "-" & iSubFldrRowCnt & "~Renamed")
				fso.movefile folderIdx.path, objFolder.path & "\" & GetDateStr() & "-" & iSubFldrRowCnt
			End If
		else
			If instr(1,strExtensions,right(folderIdx.path,3)) > 0 then
			
				if strKeepNameAsPrefix = "Yes" then 
						objOutFile.WriteLine(iRowCnt & "~" & GetDateStr() & "~" & folderIdx.name & "~" & folderIdx.name & "_" & GetDateStr() & "-" & iSubFldrRowCnt & "~Renamed")
					fso.movefile folderIdx.path, objFolder.path & "\" & folderIdx.name & "_" & GetDateStr() & "-" & iSubFldrRowCnt
				else
						objOutFile.WriteLine(iRowCnt & "~" & GetDateStr() & "~" & folderIdx.name & "~" & folderIdx.name & "_" & GetDateStr() & "-" & iSubFldrRowCnt & "~Renamed")
					fso.movefile folderIdx.path, objFolder.path & "\" & GetDateStr() & "-" & iSubFldrRowCnt
				end if
			end if 
		end if 

	Next
	iSubFldrRowCnt = 0


End Function


Function GetDateStr()
	Dim strMonth, strDay, strHour, strMinute, strSeconds
	
	If len(month(now())) = 1 then
		strMonth = "0" & month(now())
	else
		strMonth =  month(now())
	end if
	
	If len(day(now())) = 1 then
		strDay = "0" & day(now())
	else
		strDay =  day(now())
	end if	
	
	If len(hour(now())) = 1 then
		strHour = "0" & hour(now())
	else
		strHour =  hour(now())
	end if	
	
	If len(minute(now())) = 1 then
		strMinute = "0" & minute(now())
	else
		strMinute =  minute(now())
	end if	
	
	If len(Second(now()) ) = 1 then
		strSeconds = "0" & Second(now()) 
	else
		strSeconds =  Second(now()) 
	end if	
	
	GetDateStr = year(now()) & strMonth & strDay & "_" & strHour  & "_" & strMinute & "_" &  strSeconds
End Function

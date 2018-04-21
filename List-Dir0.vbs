Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(".")
Set colSubfolders = objFolder.Subfolders

For Each objSubfolder in colSubfolders
    Wscript.Echo objSubfolder.Name, objSubfolder.Size
Next
	
	
' **************************************************************************
' * Ficheiro:          cps.vbs                                             *
' * Função:            copia um ficheiro para várias sub-directorias       *
' * Sintase:           cps.vbs ficheiro dir_base dir_sub                   *
' * Autor:             Fernando Oliveira                                   *
' * Data:              2007-09-04                                          *
' **************************************************************************

On Error Resume Next

nDirOK = 0
nDirErr = 0

If Wscript.Arguments.Count < 3 Then
    Wscript.Echo "Sintase: cscript cps.vbs nome_ficheiro dir_base dir_sub"
    Wscript.Quit
Else
    ficheiro = Wscript.Arguments.Item(0)
    dirbase = Wscript.Arguments.Item(1)
    dirsub = Wscript.Arguments.Item(2)
End If

Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FileExists(ficheiro) Then
    Wscript.Echo "OK   : " & ficheiro & " existe"
Else
    Wscript.Echo "Erro : " & ficheiro & ": ficheiro não existe"
    Wscript.Quit
End If

If objFSO.FolderExists(dirbase) Then
    Wscript.Echo "OK   : " & dirbase & " existe"
Else
    Wscript.Echo "Erro : " & dirbase & ": directoria não existe"
    Wscript.Quit
End If
	
Set objFolder = objFSO.GetFolder(dirbase)
Set colSubfolders = objFolder.Subfolders
For Each objSubfolder in colSubfolders
    subdir = dirbase & "\" & objSubfolder.Name & "\" & dirsub
    if objFSO.FolderExists(subdir) Then
         objFSO.CopyFile ficheiro, subdir & "\", OverwriteExisting
         Wscript.Echo "OK   : " & ficheiro & " copiado para " & subdir
         nDirOK = nDirOK + 1
    Else
         Wscript.Echo "Aviso: " & subdir & " não existe a directoria"
         nDirErr = nDirErr + 1
    End If
Next

Wscript.Echo
Wscript.Echo nDirOK & " copia(s) do ficheiro"
Wscript.Echo nDirErr & " directoria(s) inexistente(s)"

' *****************************************************************
' * Ficheiro:          Dirp.vbs                                   *
' * Fun��o:            lista propriedades de uma directoria       *
' * Sintase:           Dirp.FO.vbs [nome_da_directoria]           *
' * Autor:             Fernando Oliveira                          *
' * Data:              2007-05-16                                 *
' *****************************************************************

On Error Resume Next

atribp = False
separap= " / " 

If Wscript.Arguments.Count > 0 Then
    chkpasta = Wscript.Arguments.Item(0)
Else
    chkpasta = "."
End If

Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FolderExists(chkPasta) Then
   Set objFolder = objFSO.GetFolder(chkPasta)
   Wscript.echo vbCrLf & "Nome              : " & objFolder.Name
   Wscript.echo "Nome curto        : " & objFolder.ShortName
   if objFolder.IsRootFolder then
       eraiz="Sim"
   else
       eraiz="N�o"
   end if
   Wscript.echo "� directoria raiz : " & eraiz
   Wscript.echo "Data cria��o      : " & objFolder.DateCreated
   Wscript.echo "�ltimo acesso     : " & objFolder.DateLastAccessed
   Wscript.echo "�ltima modifica��o: " & objFolder.DateLastModified
   Wscript.echo "Drive             : " & objFolder.Drive
   Wscript.echo "Directoria pai    : " & objFolder.ParentFolder
   Wscript.echo "Caminho           : " & objFolder.Path
   Wscript.echo "Caminho curto     : " & objFolder.ShortPath
   Wscript.echo "Tamanho           : " & objFolder.Size
   Wscript.echo "Tipo              : " & objFolder.Type
   wscript.stdout.write "Atributos         : "
   objA=objFolder.Attributes
   If objA AND 2 Then
       wscript.stdout.write "Escondido "
       atribp = True
   End If    
   If objA AND 4 Then
       if atribp then
           wscript.stdout.write separap  
       end if
       wscript.stdout.write "Sistema "
       atribp = True
   End If    
   If objA AND 16 Then
       if atribp then
           wscript.stdout.write separap  
       end if
       wscript.stdout.write "Directorio "
       atribp = True
   End If  
   If objA AND 32 Then
       if atribp then
           wscript.stdout.write separap  
       end if
       wscript.stdout.write "bit de arquivo "
       atribp = True
   End If
   If objA AND 2048 Then
       if atribp then
           wscript.stdout.write separap  
       end if
       wscript.stdout.write "Comprimido "
   End If
else
   wscript.echo "Directoria n�o existe"
End If

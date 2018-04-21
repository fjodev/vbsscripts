' *****************************************************************
' * Ficheiro:          Dirp.vbs                                   *
' * Função:            lista propriedades de uma directoria       *
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
       eraiz="Não"
   end if
   Wscript.echo "É directoria raiz : " & eraiz
   Wscript.echo "Data criação      : " & objFolder.DateCreated
   Wscript.echo "Último acesso     : " & objFolder.DateLastAccessed
   Wscript.echo "Última modificação: " & objFolder.DateLastModified
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
       wscript.stdout.write "Directório "
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
   wscript.echo "Directoria não existe"
End If

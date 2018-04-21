' *****************************************************************
' * Ficheiro:          DelTempFO.vbs                              *
' * Fun��o:            Remover a directoria TEMPFO                *
' * Sintase:           DelTempFO.vbs [Caminho directoria]         *
' * Autor:             Fernando Oliveira                          *
' * Data:              2006-06-05                                 *
' *****************************************************************

On Error Resume Next

'**********************************
'* Definicao constantes e variaveis
'**********************************

bytesPasta = 0
varPasta = "TEMPFO"
msgBTitulo = "Ficheiros tempor�rios"

'************************************************************************************
'* Pasta a apagar
'* Se � dado argumentos, a pasta a apagar � 1� argumento
'* Se n�o n�o s�o fornecidos argumentos, a pasta a apagar � da variavel de utilizador 
'*  definida em varPasta
'************************************************************************************

If Wscript.Arguments.Count > 0 Then
    chkpasta = Wscript.Arguments.Item(0)
Else
    Set objShell = WScript.CreateObject("WScript.Shell")
    Set colUserEnvVars = objShell.Environment("System")
    chkPasta = colUserEnvVars(varPasta)
End If

'******************************************************************
'* Se n�o foi definida variavel a pasta a est� vazia o programa sai
'******************************************************************

If chkPasta ="" Then
    Wscript.Quit
End If

'*************************************************
'* Se a pasta existir e tiver ficheiros � removida
'*************************************************

Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FolderExists(chkPasta) Then
   Set objFolder = objFSO.GetFolder(chkPasta)
   bytesPasta = objFolder.Size
   if bytesPasta > 0 then
      resposta=msgbox("Deseja remover a pasta " & _ 
       chkPasta & " ?", vbYesNo, msgBTitulo)
       If resposta = vbYes Then
          objFSO.DeleteFolder(chkPasta)
       End If
   End If
End If

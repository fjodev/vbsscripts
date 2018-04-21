' *****************************************************************
' * Ficheiro:          CreateTempFO.vbs                           *
' * Função:            Cria a directoria da variavel TEMPFO       *
' * Sintase:           CreateTempFO.vbs                           *
' * Autor:             Fernando Oliveira                          *
' * Data:              2006-06-07                                 *
' *****************************************************************

On Error Resume Next

'**********************************
'* Definicao constantes e variaveis
'**********************************

varPasta = "TEMPFO"

Set objShell = WScript.CreateObject("WScript.Shell")
Set colUserEnvVars = objShell.Environment("System")
chkPasta = colUserEnvVars(varPasta)

'******************************************************
'* Se existir a variavel de pasta e a pasta não existir
'*  então é criada uma pasta
'******************************************************

If chkpasta ="" Then
    MsgBox "Variavel de pasta temporária não definida", _
     vbCritical, "Criação pasta temporária - " & varPasta
    Wscript.Quit
Else
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    If Not objFSO.FolderExists(chkPasta) Then
        objFSO.CreateFolder(chkPasta)
    End If
End If
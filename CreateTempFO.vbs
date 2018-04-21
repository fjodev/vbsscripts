' *****************************************************************
' * Ficheiro:          CreateTempFO.vbs                           *
' * Fun��o:            Cria a directoria da variavel TEMPFO       *
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
'* Se existir a variavel de pasta e a pasta n�o existir
'*  ent�o � criada uma pasta
'******************************************************

If chkpasta ="" Then
    MsgBox "Variavel de pasta tempor�ria n�o definida", _
     vbCritical, "Cria��o pasta tempor�ria - " & varPasta
    Wscript.Quit
Else
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    If Not objFSO.FolderExists(chkPasta) Then
        objFSO.CreateFolder(chkPasta)
    End If
End If
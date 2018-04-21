' *****************************************************************
' * Ficheiro:          Notify.vbs                                 *
' * Função:            Produz som                                 *
' * Sintase:           Notify.vbs                                 *
' * Autor:             Fernando Oliveira                          *
' * Data:              2007-06-26                                 *
' *****************************************************************

On Error Resume Next

'**********************************
'* Definicao constantes e variaveis
'**********************************

strSoundFile = "C:\windows\Media\Notify.wav"
strCommand = "sndrec32 /play /close " & chr(34) & strSoundFile & chr(34)

'**********************************
'* Execução do som
'**********************************

Set objShell = WScript.CreateObject("WScript.Shell")
objShell.Run strCommand, 0, False
Wscript.Sleep 1000

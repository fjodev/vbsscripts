' *****************************************************************
' * Ficheiro:          saida.vbs                                  *
' * Função:            Mostra uma mensagem com acs                *
' * Sintase:           msgacs.vbs Messagem [caracter]             *
' * Autor:             Fernando Oliveira                          *
' * Data:              2006-02-06                                 *
' *****************************************************************

On Error Resume Next

strAgentName = "Merlin"
strAgentMsg  = "Hora do almoço!"

Set objShell = WScript.CreateObject("WScript.Shell")
Set colSystemEnvVars = objShell.Environment("Process")
strWINDIR = colSystemEnvVars("SystemRoot")

strAgentPath = strWINDIR & "\msagent\chars\"& strAgentName & ".acs"

Set objAgent = CreateObject("Agent.Control.2")

objAgent.Connected = TRUE
objAgent.Characters.Load strAgentName, strAgentPath
Set objCharacter = objAgent.Characters.Character(strAgentName)

objCharacter.Show

objCharacter.Play "GetAttention"
objCharacter.Play "LookDown"
objCharacter.MoveTo 500,400
objCharacter.Play "Announce"
objCharacter.Speak strAgentMsg
objCharacter.Play "Pleased"
objCharacter.Play "Idle1_1"
objCharacter.Hide

Do While objCharacter.Visible = TRUE
    Wscript.Sleep 100
Loop

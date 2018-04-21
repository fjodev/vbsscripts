strAgentName = "Merlin"
strAgentPath = "c:\winnt\msagent\chars\" & strAgentName & ".acs"
Set objAgent = CreateObject("Agent.Control.2")

objAgent.Connected = TRUE
objAgent.Characters.Load strAgentName, strAgentPath
Set objCharacter = objAgent.Characters.Character(strAgentName)

objCharacter.Show

objCharacter.Play "GetAttention"
objCharacter.Speak "Olá, tudo bem?"
objCharacter.Play "LookDown"
objCharacter.Think "Vou dar uma volta ...."
objCharacter.MoveTo 500,400
objCharacter.Play "Pleased"
objCharacter.Speak "Adeus."
objCharacter.Hide

Do While objCharacter.Visible = TRUE
    Wscript.Sleep 100
Loop

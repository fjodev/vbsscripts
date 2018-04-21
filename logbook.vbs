' *****************************************************************
' * Ficheiro:          LogBook.vbs                                *
' * Função:            Registo de eventos em ficheiro             *
' * Sintase:           LogBook.vbs [[nome_ficheiro] texto]        *
' * Autor:             Fernando Oliveira                          *
' * Data:              23-01-2006                                 *
' *****************************************************************

On Error Resume Next

'**********************************
'* Definicao constantes e variaveis
'**********************************

Const ForAppending = 8

dtmLog = now

Set objNetwork = WScript.CreateObject("WScript.Network")
strUser = objNetwork.UserName

If Wscript.Arguments.Count < 2 Then
    strFicheiro = "LogBook.log"
    strTexto = "** vazio **"
Else
    strFicheiro = Wscript.Arguments.Item(0)
    strTexto = Wscript.Arguments.Item(1)
    
End If

'******************************************************
'* Verifica se o ficheiro existe e se necess�rio cria-o
'******************************************************

Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FileExists(strFicheiro) Then
    set objTF1 = objFSO.OpenTextFile(strFicheiro, ForAppending)
Else
    set objTF1 = objFSO.CreateTextFile(strFicheiro, True)
    objTF1.WriteLine "Data" & vbTab & "Utilizador" & vbTab & "Descrição"
End If

'******************************
'* COloca o registo no ficheiro
'******************************

objTF1.WriteLine dtmLog & vbTab & strUser & vbTab & strTexto
objTF1.Close
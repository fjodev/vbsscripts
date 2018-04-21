' *****************************************************************
' * Ficheiro:          RIEP.vbs                                   *
' * Função:            Faz inventário das definições proxy do IE  *
' * Sintase:           cscript RIEP.vbs [fich_lista_computadores] *
' * Autor:             Fernando Oliveira                          *
' * Data:              21-06-2007                                 *
' *****************************************************************

on error resume next

' // Variaveis

const ForAppend = 8
const ForReading = 1

strOutFile = "RIEPxls.txt"
strInFile = "Computer.lst"
datalog = now
doLoop = True

' // Ficheiro de Input ***********************************************************

Set objFSO = CreateObject("Scripting.FileSystemObject")

If Wscript.Arguments.Count > 0 Then
    strInFile = Wscript.Arguments.Item(0)
    If objFSO.FileExists(strInFile) Then
        InFileOK = True
        set objTextFile1 = objFSO.OpenTextFile(strInFile, ForReading)
        strComputer = objTextFile1.Readline
    Else
        Wscript.Echo "Ficheiro " & strInFile & " inexistente"
        Wscript.Quit
   End If
Else
   InFileOK = False
   Set objNetwork = CreateObject("Wscript.Network")
   strComputer = objNetwork.ComputerName
End If

' // Nome do utilizador **********************************************************

Wscript.stdout.write("A analisar " & strComputer & " ")
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
Wscript.StdOut.Write(".")

Set colComputer = objWMIService.ExecQuery _
    ("Select * from Win32_ComputerSystem")
Wscript.StdOut.Write(".")

 
For Each objComputer in colComputer
    UserName = objComputer.UserName
    Wscript.StdOut.Write(".")
Next

' // Ficheiro de Output	**********************************************************

If objFSO.FileExists(strOutFile) Then
    Set objTextFile2 = objFSO.OpenTextFile(strOutFile, ForAppend, True)
else
    Set objTextFile2 = objFSO.CreateTextFile(strOutFile, True)
    objTextFile2.WriteLine "Data" & vbTab & "Computer" & vbTab & "User" & vbTab & _
     "Proxy" & vbTab & "Server" & vbTab & "Override"
End If
Wscript.StdOut.Write(".")

' // Ciclo de Leitura da defini��es do proxy do IE ***************************************

Do While doLoop
    Set objWMIService = GetObject("winmgmts:" & _
     "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2\Applications\MicrosoftIE")
    Wscript.StdOut.Write(".")
    If Err.Number =  0 Then
        ComputerOK = True
        Wscript.StdOut.Write(".")
    else
        ComputerOK = False
            Wscript.StdOut.Write(" KO" & vbCrLf)
    End If
 
    if ComputerOK Then
        Set colIESettings = objWMIService.ExecQuery ("Select * from MicrosoftIE_LANSettings")
        Wscript.StdOut.Write(".")

        For Each strIESetting in colIESettings
            datalog = now
            Wscript.StdOut.Write(".")
            objTextFile2.WriteLine datalog & vbTab & strComputer & vbTab & UserName & vbTab & _ 
             strIESetting.Proxy & vbTab & strIESetting.ProxyServer & vbTab & _ 
             strIESetting.ProxyOverride
            Wscript.StdOut.Write(".")
        Next
        Wscript.StdOut.Write(". OK" & vbCrLf)
    End If

    If InFileOK Then
        If objTextFile1.AtEndOfStream Then
            doLoop = False
        Else
	    StrComputer = objTextFile1.Readline
        End If
    Else
        doLoop = False
    End If
Loop

If InFileOK Then
    objTextFile1.Close
End If
objTextFile2.Close
Wscript.Echo "Terminado"
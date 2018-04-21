' *****************************************************************
' * Ficheiro:          IEProxy.vbs                                *
' * Fun��o:            Faz invent�rio das defini��es proxy do IE  *
' * Sintase:           cscript IEProxy.vbs [-n]                   *
' * Autor:             Fernando Oliveira                          *
' * Data:              19-06-2007                                 *
' *****************************************************************

const ForWrite = 2
const ForAppend = 8

titulo = "          DEFINIÇÕES PROXY DO INTERNET EXPLORER"
linha = "-----------------------------------------------------------"
strFile1 = "IEproxy.log"
strFile2 = "IEproxyxls.txt"
datalog = now

if Wscript.Arguments.Count > 0 Then
    if Wscript.Arguments.Item(0) = "-n" or Wscript.Arguments.Item(0) = "-N" Then
        OutFile = False
    Else
        OutFile = True
        Wscript.Echo "--> Argumento desconhecido"
    End If
Else
    OutFile = True
End If

Set objNetwork = CreateObject("Wscript.Network")
strComputer = objNetwork.ComputerName

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 

Set colComputer = objWMIService.ExecQuery _
    ("Select * from Win32_ComputerSystem")
 
For Each objComputer in colComputer
    UserName = objComputer.UserName
Next
	
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\cimv2\Applications\MicrosoftIE")

Set colIESettings = objWMIService.ExecQuery _
    ("Select * from MicrosoftIE_LANSettings")

Set objFSO = CreateObject("Scripting.FileSystemObject")

If OutFile Then
    If objFSO.FileExists(strFile1) Then
        Set objTextFile1 = objFSO.OpenTextFile(strFile1, ForAppend, True)
    else
        Set objTextFile1 = objFSO.CreateTextFile(strFile1, True)
    End If

    If objFSO.FileExists(strFile2) Then
        Set objTextFile2 = objFSO.OpenTextFile(strFile2, ForAppend, True)
    else
        Set objTextFile2 = objFSO.CreateTextFile(strFile2, True)
        objTextFile2.WriteLine "Data" & vbTab & "Computer" & vbTab & "User" & vbTab & _
         "Proxy" & vbTab & "Server" & vbTab & "Override"
    End If
    Wscript.Echo "--> Resultado registado nos ficheiros " & strFile1 & " e " & strFile2
Else
    Wscript.Echo "--> Resultado não registado"
End If

Wscript.Echo
wscript.Echo titulo
wscript.Echo linha
wscript.Echo "Data                    : " & dataLog
wscript.Echo "User name               : " & Username
wscript.Echo "Computer name           : " & strComputer
For Each strIESetting in colIESettings
    wscript.Echo linha
    Wscript.Echo "Autocfg proxy           : " & strIESetting.AutoConfigProxy
    Wscript.Echo "Autocfg URL             : " & strIESetting.AutoConfigURL
    Wscript.Echo "Autocfg Proxy det. mode : " & strIESetting.AutoProxyDetectMode
    wscript.Echo
    Wscript.Echo "Proxy                   : " & strIESetting.Proxy
    Wscript.Echo "Proxy server            : " & strIESetting.ProxyServer
    Wscript.Echo "Proxy override          : " & strIESetting.ProxyOverride

    If OutFile Then
        objTextFile1.WriteLine linha
        objTextFile1.WriteLine "Data hora recolha " & vbTab & dataLog
        objTextFile1.WriteLine "User name         " & vbTab & UserName
        objTextFile1.WriteLine "Computer name     " & vbTab & strComputer
        objTextFile1.WriteLine "Autocfg proxy     " & vbTab & strIESetting.AutoConfigProxy
        objTextFile1.WriteLine "Autocfg URL       " & vbTab & strIESetting.AutoConfigURL
        objTextFile1.WriteLine "Autocfg det. mode " & vbTab & strIESetting.AutoProxyDetectMode
        objTextFile1.WriteLine "Proxy             " & vbTab & strIESetting.Proxy
        objTextFile1.WriteLine "Proxy server      " & vbTab & strIESetting.ProxyServer
        objTextFile1.WriteLine "Proxy override    " & vbTab & strIESetting.ProxyOverride
        objTextFile2.WriteLine datalog & vbTab & strComputer & vbTab & UserName & vbTab & _ 
         strIESetting.Proxy & vbTab & strIESetting.ProxyServer & vbTab & _ 
         strIESetting.ProxyOverride
    End If
Next
wscript.Echo linha
If OutFile Then
	objTextFile1.Close
	objTextFile2.Close
End If
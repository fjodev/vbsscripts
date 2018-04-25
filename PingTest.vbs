' ********************************************************************
' ** Ficheiro:          PingTest.vbs                                **
' ** Função:            Testa ligação por ping                      **
' ** Sintase:           cscript PingTest.vbs [fich_log] [fich_end]  **
' ** Autor:             Fernando Oliveira                           **
' ** Data:              07-03-2008                                  **
' ********************************************************************

' ************************************************
' ** Variaveis, constantes e outras declarações **
' ************************************************

const ForReading = 1
const ForWrite = 2
const ForAppend = 8

Select Case Wscript.Arguments.Count
    Case 0
        strFile1 = "PingTest.log"
        strFile2 = "PingTest.lst"
    case 1
        strFile1 = Wscript.Arguments.Item(0)
        strFile2 = "PingTest.lst"
    case Else
        strFile1 = Wscript.Arguments.Item(0)
        strFile2 = Wscript.Arguments.Item(1)
End Select

iReadFile = 0

Dim aMachinesName()
Dim aMachinesIP()

Set objFSO = CreateObject("Scripting.FileSystemObject")

' ******************************************
' ** Obt�m a lista de IP's a testar       **
' ******************************************

if objFSO.FileExists(strFile2) Then
    Set objTextFile2 = objFSO.OpenTextFile(strFile2, ForReading)
    Do Until objTextFile2.AtEndOfStream
        strNextLine = objTextFile2.Readline
        aListMachines = Split(strNextLine , ",")
        ReDim Preserve aMachinesName(iReadFile)
        aMachinesName(iReadFile) = aListMachines(0)
        ReDim Preserve aMachinesIP(iReadFile)
        aMachinesIP(iReadFile) = aListMachines(1)
        iReadFile = iReadFile + 1
    Loop
Else
    Wscript.Echo "Ficheiro " & strFile2 & " inexistente!" _ 
      & vbCrLf & "Programa vai terminar."
    Wscript.Quit
End If

' ******************************************
' ** Verifica se o ficheiro de Log existe **
' ** e em caso negativo cria-o            **
' ******************************************

If objFSO.FileExists(strFile1) Then
    Set objTextFile1 = objFSO.OpenTextFile(strFile1, ForAppend, True)
else
    Set objTextFile1 = objFSO.CreateTextFile(strFile1, True)
    objTextFile1.WriteLine "# " & strFile1
    ObjtextFile1.WriteLine "# Campos: Data Hora Maquina_origem ip Maquina_destino ip resultado"
    ObjtextFile1.WriteLine
End If

' *******************************************
' ** Determina o nome da maquina de origem **
' *******************************************

Set objNetwork = CreateObject("Wscript.Network")
strComputer = objNetwork.ComputerName

' *******************************************
' ** Determina o End. IP da Maquina Origem **
' *******************************************

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set IPConfigSet = objWMIService.ExecQuery _
    ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")

' **********************************************************
' ** Ciclos de ping's                                     **
' ** por cada ip de origem � realizado um ciclo de ping   **
' ** para cada ip da lista de destinos                    ** 
' **********************************************************

For Each IPConfig in IPConfigSet
    If Not IsNull(IPConfig.IPAddress) Then 
        For iSIP = LBound(IPConfig.IPAddress) to UBound(IPConfig.IPAddress)
            For iDIP = 0 to Ubound(aMachinesIP) 
                dataPing = now
                Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}")._
                  ExecQuery("select * from Win32_PingStatus where address = '"_
                  & aMachinesIP(iDIP) & "'")
                For Each objStatus in objPing
                    ' intPadding = 20 - len(aMachinesName(iDIP))
                    strLineFile = dataPing & vbTab & strComputer & vbTab & _ 
                      IPConfig.IPAddress(iSIP) & vbTab & _
                      rtrim(aMachinesName(iDIP)) & vbTab & aMachinesIP(iDIP)
                    If IsNull(objStatus.StatusCode) or objStatus.StatusCode<>0 Then 
                        strResult = "FALHA"
                    Else
                        strResult = "OK"
                    End If
                    objTextFile1.WriteLine strLineFile & vbTab & strResult
                    strDisplay = strDisplay & vbCrLf & _ 
                      aMachinesName(iDIP) & vbTab & strResult
                Next    
            Next
       Next
    End If
Next        
objTextFile1.Close

' ***********************************
' ** Mostra o resultado do teste   **
' ***********************************

Wscript.Echo strDisplay

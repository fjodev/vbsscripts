On Error Resume Next

Const ForReading = 1
Const InFile = "computer.lst"

Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FileExists(InFile) Then
    set objTextFile = objFSO.OpenTextFile(InFile, ForReading)

    Wscript.Echo "Total de memória por computador listado em " & InFile & ":" & vbCrLf

    Do Until objTextFile.AtEndOfStream
        StrComputer = objTextFile.Readline
        Set objSWbemServices = GetObject("winmgmts:\\" & strComputer)
        Set colSWbemObjectSet = _
        objSWbemServices.InstancesOf("Win32_LogicalMemoryConfiguration")

        For Each objSWbemObject In colSWbemObjectSet
           Wscript.Echo strComputer & " = " & vbTab & _
            objSWbemObject.TotalPhysicalMemory
        Next
    Loop
else
    StrComputer="."
    Set objSWbemServices = GetObject("winmgmts:\\" & strComputer)
    Set colSWbemObjectSet = _
        objSWbemServices.InstancesOf("Win32_LogicalMemoryConfiguration")

    For Each objSWbemObject In colSWbemObjectSet
        Wscript.Echo "Memória local = " & vbTab & _
         objSWbemObject.TotalPhysicalMemory
    Next
End If
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colPageFiles = objWMIService.ExecQuery("Select * from Win32_PageFileUsage")

wscript.echo

For Each objPageFile in colPageFiles
    Wscript.Echo "Allocated Base Size: " & objPageFile.AllocatedBaseSize
    Wscript.Echo "Current Usage      : " & objPageFile.CurrentUsage
    Wscript.Echo "Description        : " & objPageFile.Description
    InstDate = WMIDateStringToDate(objPageFile.InstallDate)
    Wscript.Echo "Install Date       : " & InstDate
    Wscript.Echo "Name               : " & objPageFile.Name   
    Wscript.Echo "Peak Usage         : " & objPageFile.PeakUsage 
Next

Function WMIDateStringToDate(dtmInstallDate)
    WMIDateStringToDate = CDate(Mid(dtmInstallDate, 7, 2) & "/" & _
        Mid(dtmInstallDate, 5, 2) & "/" & Left(dtmInstallDate, 4) _
            & " " & Mid (dtmInstallDate, 9, 2) & ":" & _
                Mid(dtmInstallDate, 11, 2) & ":" & Mid(dtmInstallDate, _
                    13, 2))
End Function

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colSettings = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

For Each objOS in colSettings 
    strDisplay = "OS Name: " & vbTab & objOS.caption & vbCrLf & _
     "Version: " & objOS.Version & vbCrLf & _
     "Service Pack               : " & objOS.ServicePackMajorVersion _
            & "." & objOS.ServicePackMinorVersion & vbCrLf & _
     "OS Manufacturer: " & vbTab & objOS.Manufacturer & vbCrLf & _
     "Windows Directory: " & vbTab & objOS.WindowsDirectory & vbCrLf & _
     "Locale: " & vbTab & objOS.Locale & vbCrLf & _
     "Available Physical Memory: " & vbTab & objOS.FreePhysicalMemory & vbCrLf & _
     "Total Virtual Memory: " & vbTab & objOS.TotalVirtualMemorySize & vbCrLf & _
     "Available Virtual Memory: " & vbTab & objOS.FreeVirtualMemory & vbCrLf & _
     "Size stored in paging files: " & vbTab & objOS.SizeStoredInPagingFiles
Next

Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")

For Each objComputer in colSettings 
    strDisplay = strDisplay & vbCrLf & "System Name: " & vbTab & objComputer.Name & _ 
       vbCrLf & _
     "System Manufacturer: " & vbTab & objComputer.Manufacturer & vbCrLf & _
     "System Model: " & vbTab & objComputer.Model & vbCrLf & _
     "Time Zone: " & vbTab & objComputer.CurrentTimeZone & vbCrLf & _
     "Total Physical Memory: " & vbTab & objComputer.TotalPhysicalMemory 
Next

Set colSettings = objWMIService.ExecQuery _
    ("Select * from Win32_Processor")

For Each objProcessor in colSettings 
    strDisplay = strDisplay & vbCrLf & "System Type                : " & objProcessor.Architecture & _
      vbCrLf & "Processor: " & vbTab & objProcessor.Description
Next

Set colSettings = objWMIService.ExecQuery("Select * from Win32_BIOS")

For Each objBIOS in colSettings 
    strDisplay = strDisplay & vbCrLf & "BIOS Version: " & vbTab & objBIOS.Version
Next
Wscript.Echo StrDisplay
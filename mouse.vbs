strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colMice = objWMIService.ExecQuery _
    ("Select * from Win32_PointingDevice")

For Each objMouse in colMice
    Wscript.Echo "Hardware Type: " & objMouse.HardwareType
    Wscript.Echo "Number of Buttons: " & objMouse.NumberOfButtons    
    Wscript.Echo "Status: " & objMouse.Status
    Wscript.Echo "PNP Device ID: " & objMouse.PNPDeviceID
Next
	
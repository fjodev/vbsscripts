Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colOperatingSystems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

For Each objOperatingSystem in colOperatingSystems
    dtmConvertedDate.Value = objOperatingSystem.InstallDate
    dtmInstallDate = dtmConvertedDate.GetVarDate
     Wscript.Echo "Caption          : " & objOperatingSystem.Caption & vbCrLf & _
     "Version          : " & objOperatingSystem.Version & vbCrLf & _
     "Service Pack     : " & objOperatingSystem.ServicePackMajorVersion  _
        & "." & objOperatingSystem.ServicePackMinorVersion & vbCrLf & _
     "Build Number     : " & objOperatingSystem.BuildNumber & vbCrLf & _
     "Build Type       : " & objOperatingSystem.BuildType & vbCrLf & _
     "Install Date     : " & dtmInstallDate  & vbCrLf & _
     "Serial Number    : " & objOperatingSystem.SerialNumber & vbCrLf & _
     "Code Set         : " & objOperatingSystem.CodeSet & vbCrLf & _
     "Country Code     : " & objOperatingSystem.CountryCode & vbCrLf & _
     "Debug            : " & objOperatingSystem.Debug & vbCrLf & _
     "Encryption Level : " & objOperatingSystem.EncryptionLevel & vbCrLf & _
     "Organization     : " & objOperatingSystem.Organization & vbCrLf & _
     "Registered User  : " & objOperatingSystem.RegisteredUser & vbCrLf & _
     "Licensed Users   : " & objOperatingSystem.NumberOfLicensedUsers & vbCrLf & _
     "OS Language      : " & objOperatingSystem.OSLanguage & vbCrLf & _
     "OS Product Suite : " & objOperatingSystem.OSProductSuite & vbCrLf & _
     "OS Type          : " & objOperatingSystem.OSType & vbCrLf & _
     "Primary          : " & objOperatingSystem.Primary & vbCrLf & _
     "Boot Device      : " & objOperatingSystem.BootDevice
Next

	
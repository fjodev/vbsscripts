On Error Resume Next

Wscript.Echo vbCrLf & "Running. Wait please... " & vbCrLf

dim OutFile1
dim OutFile2

OutFile1 = "lssw.tsv"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.CreateTextFile(OutFile1, True)

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
 & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

SecTitulo = "*** Operation System Properties"
Wscript.Echo "STEP 1 " & SecTitulo & vbCrLf
objTextFile.WriteLine SecTitulo & String(3,"*") & vbCrLf

Set colOSes = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

For Each objOS in colOSes
    objTextFile.WriteLine "Computer Name" & vbTab & objOS.CSName
    OutFile2 = objOS.CSName & "_sw.tsv"
    objTextFile.WriteLine "caption" & vbTab & objOS.Caption
    objTextFile.WriteLine "Version" & vbTab & objOS.Version
    objTextFile.WriteLine "Build Number" & vbTab & objOS.BuildNumber 
    objTextFile.WriteLine "Build Type" & vbTab & objOS.BuildType
    objTextFile.WriteLine "OS Type" & vbTab & objOS.OSType
    objTextFile.WriteLine "Other Type Description" & vbTab & objOS.OtherTypeDescription
    objTextFile.WriteLine "Service Pack" & vbTab & objOS.ServicePackMajorVersion & "." & _ 
     objOS.ServicePackMinorVersion

    dtmConvertedDate.Value = objOS.InstallDate
    dtmInstallDate = dtmConvertedDate.GetVarDate

    objTextFile.WriteLine "Boot Device" & vbTab & objOS.BootDevice
    objTextFile.WriteLine "Code Set" & vbTab & objOS.CodeSet 
    objTextFile.WriteLine "Country Code" & vbTab & objOS.CountryCode 
    objTextFile.WriteLine "Debug" & vbTab & objOS.Debug 
    objTextFile.WriteLine "Encryption Level" & vbTab & objOS.EncryptionLevel 
    objTextFile.WriteLine "Install Date" & vbTab & objOS.InstallDate & vbTab & dtmInstallDate 
    objTextFile.WriteLine "Licensed Users" & vbTab & objOS.NumberOfLicensedUsers 
    objTextFile.WriteLine "Organization" & vbTab & objOS.Organization 
    objTextFile.WriteLine "OS Language" & vbTab & objOS.OSLanguage 
    objTextFile.WriteLine "OS Product Suite" & vbTab & objOS.OSProductSuite 
    objTextFile.WriteLine "OS Type" & vbTab & objOS.OSType 
    objTextFile.WriteLine "Primary" & vbTab & objOS.Primary 
    objTextFile.WriteLine "Registered User" & vbTab & objOS.RegisteredUser 
    objTextFile.WriteLine "Serial Number" & vbTab & objOS.SerialNumber 
    objTextFile.WriteLine "Version" & vbTab & objOS.Version & vbCrLf
    wscript.Echo "Computer Name - " & objOS.CSName
    Wscript.Echo "OS Caption    - "  & objOS.Caption
    wscript.Echo "Service Pack  - " & objOS.ServicePackMajorVersion & "." & _ 
     objOS.ServicePackMinorVersion & vbCrLf
Next

SecTitulo = "*** Installed Software "
Wscript.Echo "STEP 2 " & SecTitulo & vbCrLf
objTextFile.WriteLine SecTitulo & String(3,"*") & vbCrLf

Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
strComputer = "."
strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
strEntry1a = "DisplayName"
strEntry1b = "QuietDisplayName"
strEntry2 = "InstallDate"
strEntry3 = "VersionMajor"
strEntry4 = "VersionMinor"
strEntry5 = "EstimatedSize"

Set objReg = GetObject("winmgmts://" & strComputer & _
 "/root/default:StdRegProv")
objReg.EnumKey HKLM, strKey, arrSubkeys

objTextFile.WriteLine "Display Name" & vbTab & _
 "Install Date" & vbTab & _ 
 "Version" & vbTab & _ 
 "Estimated Size (MB)"

contador = 0

For Each strSubkey In arrSubkeys
    intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntry1a, strValue1)
    If intRet1 <> 0 Then
        objReg.GetStringValue HKLM, strKey & strSubkey, strEntry1b, strValue1
    End If
    If strValue1 <> "" Then
        objReg.GetStringValue HKLM, strKey & strSubkey, strEntry2, strValue2  
        objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntry3, intValue3
        objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntry4, intValue4
        objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntry5, intValue5

        objTextFile.WriteLine strValue1 & vbTab & _ 
         strValue2 & vbTab & _
         intValue3 & "." & intValue4 & vbTab & _
         Round(intValue5/1024, 3)
        contador = contador + 1
        Wscript.Echo "App" & contador & " - " & strValue1
    End If
Next
objTextFile.Close
objFSO.DeleteFile(OutFile2)
objFSO.MoveFile OutFile1, OutFile2
Wscript.Echo vbCrLf & "Finish. Result on " & OutFile2 & vbCrLf
' *****************************************************************
' * Ficheiro:          rlssw.vbs                                  *
' * Função:            Faz inventário do software dum computador  *
' * Sintase:           rlssw.vbs [nome_ficheiro]                  *
' *                    Se omitido o ficheiro faz inventário local *
' * Autor:             Fernando Oliveira                          *
' * Data:              24-01-2006                                 *
' *****************************************************************


On Error Resume Next

Const ForReading = 1
doLoop = True
dim strOutFile

Set objFSO = CreateObject("Scripting.FileSystemObject")

If Wscript.Arguments.Count > 0 Then
    strInFile = Wscript.Arguments.Item(0)
    If objFSO.FileExists(strInFile) Then
        strInFileOK = True
        set objTextFile1 = objFSO.OpenTextFile(strInFile, ForReading)
        strComputer = objTextFile1.Readline
        strOutFileP1 = strComputer
    Else
        Wscript.Echo "Ficheiro " & strInFile & " inexistente"
        Wscript.Quit
   End If
Else
   strInFileOK = False
   strComputer = "."
   strOutFileP1 = "local"   
End If

Do While doLoop

    Wscript.StdOut.Write "A analisar o software do computador " & strOutFileP1 & ". Pf aguarde..."
    Set objWMIService = GetObject("winmgmts:" _
     & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

    strOutFile = strOutFileP1 & "_sw.tsv"
    Set objTextFile2 = objFSO.CreateTextFile(strOutFile, True)

    SecTitulo = "*** Operation System Properties"
    objTextFile2.WriteLine SecTitulo & String(3,"*") & vbCrLf
    Set colOSes = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

    For Each objOS in colOSes
        Wscript.StdOut.Write(".")
        objTextFile2.WriteLine "Computer Name" & vbTab & objOS.CSName
        objTextFile2.WriteLine "caption" & vbTab & objOS.Caption
        objTextFile2.WriteLine "Version" & vbTab & objOS.Version
        objTextFile2.WriteLine "Build Number" & vbTab & objOS.BuildNumber 
        objTextFile2.WriteLine "Build Type" & vbTab & objOS.BuildType
        objTextFile2.WriteLine "OS Type" & vbTab & objOS.OSType
        objTextFile2.WriteLine "Other Type Description" & vbTab & objOS.OtherTypeDescription
        objTextFile2.WriteLine "Service Pack" & vbTab & objOS.ServicePackMajorVersion & "." & _ 
        objOS.ServicePackMinorVersion

        dtmConvertedDate.Value = objOS.InstallDate
        dtmInstallDate = dtmConvertedDate.GetVarDate

        objTextFile2.WriteLine "Boot Device" & vbTab & objOS.BootDevice
        objTextFile2.WriteLine "Code Set" & vbTab & objOS.CodeSet 
        objTextFile2.WriteLine "Country Code" & vbTab & objOS.CountryCode 
        objTextFile2.WriteLine "Debug" & vbTab & objOS.Debug 
        objTextFile2.WriteLine "Encryption Level" & vbTab & objOS.EncryptionLevel 
        objTextFile2.WriteLine "Install Date" & vbTab & objOS.InstallDate & vbTab & dtmInstallDate 
        objTextFile2.WriteLine "Licensed Users" & vbTab & objOS.NumberOfLicensedUsers 
        objTextFile2.WriteLine "Organization" & vbTab & objOS.Organization 
        objTextFile2.WriteLine "OS Language" & vbTab & objOS.OSLanguage 
        objTextFile2.WriteLine "OS Product Suite" & vbTab & objOS.OSProductSuite 
        objTextFile2.WriteLine "OS Type" & vbTab & objOS.OSType 
        objTextFile2.WriteLine "Primary" & vbTab & objOS.Primary 
        objTextFile2.WriteLine "Registered User" & vbTab & objOS.RegisteredUser 
        objTextFile2.WriteLine "Serial Number" & vbTab & objOS.SerialNumber 
        objTextFile2.WriteLine "Version" & vbTab & objOS.Version & vbCrLf
    Next

    SecTitulo = "*** Installed Software "
    objTextFile2.WriteLine SecTitulo & String(3,"*") & vbCrLf

    Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
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

    objTextFile2.WriteLine "Display Name" & vbTab & _
     "Install Date" & vbTab & _ 
     "Version" & vbTab & _ 
     "Estimated Size (MB)"

    For Each strSubkey In arrSubkeys
        Wscript.StdOut.Write(".")
        intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntry1a, strValue1)
        If intRet1 <> 0 Then
            objReg.GetStringValue HKLM, strKey & strSubkey, strEntry1b, strValue1
        End If
        If strValue1 <> "" Then
            objReg.GetStringValue HKLM, strKey & strSubkey, strEntry2, strValue2  
            objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntry3, intValue3
            objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntry4, intValue4
            objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntry5, intValue5

            objTextFile2.WriteLine strValue1 & vbTab & _ 
             strValue2 & vbTab & _
             intValue3 & "." & intValue4 & vbTab & _
             Round(intValue5/1024, 3)
        End If
    Next
    objTextFile2.Close

    Wscript.Echo vbCrLf & "Analise completa. Resultado em " & _ 
     strOutFile & vbCrLf

    If strInFileOK Then
        If objTextFile1.AtEndOfStream Then
            doLoop = False
        Else
	    StrComputer = objTextFile1.Readline
            strOutFileP1 = strComputer
        End If
    Else
        doLoop = False
    End If
Loop
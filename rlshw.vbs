' *****************************************************************
' * Ficheiro:          rlshw.vbs                                  *
' * Função:            Faz inventário do hardware dum computador  *
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
    Wscript.StdOut.Write "A analisar o hardware do computador " & strOutFileP1 & ". Pf aguarde..."
    Set objWMIService = GetObject("winmgmts:" _
     & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

    strOutFile = strOutFileP1 & "_hw.tsv"
    Set objTextFile2 = objFSO.CreateTextFile(strOutFile, True)

    Set objWMIService = GetObject("winmgmts:" _
     & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

    SecTitulo = "*** BIOS "
    objTextFile2.WriteLine SecTitulo & String(3,"*") & vbCrLf

    Set colBIOS = objWMIService.ExecQuery("Select * from Win32_BIOS")

    For each objBIOS in colBIOS
        Wscript.StdOut.Write(".")
        objTextFile2.WriteLine "Build Number" & vbTab & objBIOS.BuildNumber
        objTextFile2.WriteLine "Current Language" & vbTab & objBIOS.CurrentLanguage
        objTextFile2.WriteLine "Installable Languages" & vbTab & objBIOS.InstallableLanguages
        objTextFile2.WriteLine "Manufacturer" & vbTab & objBIOS.Manufacturer
        objTextFile2.WriteLine "Name" & vbTab & objBIOS.Name
        objTextFile2.WriteLine "Primary BIOS" & vbTab & objBIOS.PrimaryBIOS
        objTextFile2.WriteLine "Release Date" & vbTab & objBIOS.ReleaseDate
        objTextFile2.WriteLine "Serial Number" & vbTab & objBIOS.SerialNumber
        objTextFile2.WriteLine "SMBIOS Version" & vbTab & objBIOS.SMBIOSBIOSVersion
        objTextFile2.WriteLine "SMBIOS Major Version" & vbTab & objBIOS.SMBIOSMajorVersion
        objTextFile2.WriteLine "SMBIOS Minor Version" & vbTab & objBIOS.SMBIOSMinorVersion
        objTextFile2.WriteLine "SMBIOS Present" & vbTab & objBIOS.SMBIOSPresent
        objTextFile2.WriteLine "Status" & vbTab & objBIOS.Status
        objTextFile2.WriteLine "Version" & vbTab & objBIOS.Version
        objTextFile2.WriteLine
    Next

    SecTitulo = "*** Processor "
    objTextFile2.WriteLine SecTitulo & String(3,"*") & vbCrLf

    Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor")

    For Each objItem in colItems
        Wscript.StdOut.Write(".")
        objTextFile2.WriteLine "Address Width" & vbTab & objItem.AddressWidth
        objTextFile2.WriteLine "Architecture" & vbTab & objItem.Architecture
        objTextFile2.WriteLine "Availability" & vbTab & objItem.Availability
        objTextFile2.WriteLine "CPU Status" & vbTab & objItem.CpuStatus
        objTextFile2.WriteLine "Current Clock Speed" & vbTab & objItem.CurrentClockSpeed
        objTextFile2.WriteLine "Data Width" & vbTab & objItem.DataWidth
        objTextFile2.WriteLine "Description" & vbTab & objItem.Description
        objTextFile2.WriteLine "Device ID" & vbTab & objItem.DeviceID
        objTextFile2.WriteLine "External Clock" & vbTab & objItem.ExtClock
        objTextFile2.WriteLine "Family" & vbTab & objItem.Family
        objTextFile2.WriteLine "L2 Cache Size" & vbTab & objItem.L2CacheSize
        objTextFile2.WriteLine "L2 Cache Speed" & vbTab & objItem.L2CacheSpeed
        objTextFile2.WriteLine "Level" & vbTab & objItem.Level
        objTextFile2.WriteLine "Load Percentage" & vbTab & objItem.LoadPercentage
        objTextFile2.WriteLine "Manufacturer" & vbTab & objItem.Manufacturer
        objTextFile2.WriteLine "Maximum Clock Speed" & vbTab & objItem.MaxClockSpeed
        objTextFile2.WriteLine "Name" & vbTab & objItem.Name
        objTextFile2.WriteLine "Processor ID" & vbTab & objItem.ProcessorId
        objTextFile2.WriteLine "Processor Type" & vbTab & objItem.ProcessorType
        objTextFile2.WriteLine "Revision" & vbTab & objItem.Revision
        objTextFile2.WriteLine "Socket Designation" & vbTab & objItem.SocketDesignation
        objTextFile2.WriteLine "Status Information" & vbTab & objItem.StatusInfo
        objTextFile2.WriteLine "Stepping" & vbTab & objItem.Stepping
        objTextFile2.WriteLine "Unique Id" & vbTab & objItem.UniqueId
        objTextFile2.WriteLine "Upgrade Method" & vbTab & objItem.UpgradeMethod
        objTextFile2.WriteLine "Version" & vbTab & objItem.Version
        objTextFile2.WriteLine "Voltage Caps" & vbTab & objItem.VoltageCaps
        objTextFile2.WriteLine
    Next

    SecTitulo = "*** Motherboard Device "
    objTextFile2.WriteLine SecTitulo & String(3,"*") & vbCrLf

    Set colItems = objWMIService.ExecQuery("Select * from Win32_MotherboardDevice")

    For Each objItem in colItems
        Wscript.StdOut.Write(".")
        objTextFile2.WriteLine "Device ID" & vbTab & objItem.DeviceID
        objTextFile2.WriteLine "Primary Bus Type" & vbTab & objItem.PrimaryBusType
        objTextFile2.WriteLine "Secondary Bus Type" & vbTab & objItem.SecondaryBusType
        objTextFile2.WriteLine
    Next

    SecTitulo = "*** Onboard Device "
    objTextFile2.WriteLine SecTitulo & String(3,"*") & vbCrLf

    Set colItems = objWMIService.ExecQuery("Select * from Win32_OnBoardDevice")

    For Each objItem in colItems
        Wscript.StdOut.Write(".")
        objTextFile2.WriteLine "Description" & vbTab & objItem.Description
        objTextFile2.WriteLine "Device Type" & vbTab & objItem.DeviceType
        objTextFile2.WriteLine "Model" & vbTab & objItem.Model
        objTextFile2.WriteLine "Name" & vbTab & objItem.Name
        objTextFile2.WriteLine "Tag" & vbTab & objItem.Tag
        objTextFile2.WriteLine "Version" & vbTab & objItem.Version
        objTextFile2.WriteLine
    Next

    SecTitulo = "*** Phisical Memory "
    objTextFile2.WriteLine SecTitulo & String(3,"*") & vbCrLf

    Set objSWbemServices = GetObject("winmgmts:\\" & strComputer)
    Set colSWbemObjectSet = _
    objSWbemServices.InstancesOf("Win32_LogicalMemoryConfiguration")

    For Each objSWbemObject In colSWbemObjectSet
        objTextFile2.WriteLine "Total Memory" & vbTab & objSWbemObject.TotalPhysicalMemory _
         & vbCrLf     
    Next

    Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemoryArray")

    For Each objItem in colItems
        Wscript.StdOut.Write(".")
        objTextFile2.WriteLine "Description" & vbTab & objItem.Description
        objTextFile2.WriteLine "Maximum Capacity" & vbTab & objItem.MaxCapacity
        objTextFile2.WriteLine "Memory Devices" & vbTab & objItem.MemoryDevices
        objTextFile2.WriteLine "Memory Error Correction" & vbTab & objItem.MemoryErrorCorrection
        objTextFile2.WriteLine
    Next

    SecTitulo = "*** Memory Configuration "
    objTextFile2.WriteLine SecTitulo & String(3,"*") & vbCrLf

    Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory")

    For Each objItem in colItems
        Wscript.StdOut.Write(".")
        objTextFile2.WriteLine "Tag" & vbTab & objItem.Tag
        objTextFile2.WriteLine "Bank Label" & vbTab & objItem.BankLabel
        objTextFile2.WriteLine "Capacity" & vbTab & objItem.Capacity
        objTextFile2.WriteLine "Data Width" & vbTab & objItem.DataWidth
        objTextFile2.WriteLine "Description" & vbTab & objItem.Description
        objTextFile2.WriteLine "Device Locator" & vbTab & objItem.DeviceLocator
        objTextFile2.WriteLine "Form Factor" & vbTab & objItem.FormFactor
        objTextFile2.WriteLine "Hot Swappable" & vbTab & objItem.HotSwappable
        objTextFile2.WriteLine "Manufacturer" & vbTab & objItem.Manufacturer
        objTextFile2.WriteLine "Memory Type" & vbTab & objItem.MemoryType
        objTextFile2.WriteLine "Name" & vbTab & objItem.Name
        objTextFile2.WriteLine "Part Number" & vbTab & objItem.PartNumber
        objTextFile2.WriteLine "Position In Row" & vbTab & objItem.PositionInRow
        objTextFile2.WriteLine "Speed" & vbTab & objItem.Speed
        objTextFile2.WriteLine "Type Detail" & vbTab & objItem.TypeDetail
        objTextFile2.WriteLine
    Next

    SecTitulo = "*** Sound Card Properties "
    objTextFile2.WriteLine SecTitulo & String(3,"*") & vbCrLf

    Set colItems = objWMIService.ExecQuery("Select * from Win32_SoundDevice")

    For Each objItem in colItems
        Wscript.StdOut.Write(".")
        objTextFile2.WriteLine "Description" & vbTab & objItem.Description
        objTextFile2.WriteLine "Device ID" & vbTab & objItem.DeviceID
        objTextFile2.WriteLine "DMA Buffer Size" & vbTab & objItem.DMABufferSize
        objTextFile2.WriteLine "Manufacturer" & vbTab & objItem.Manufacturer
        objTextFile2.WriteLine "MPU 401 Address" & vbTab & objItem.MPU401Address
        objTextFile2.WriteLine "Name" & vbTab & objItem.Name
        objTextFile2.WriteLine "PNP Device ID" & vbTab & objItem.PNPDeviceID
        objTextFile2.WriteLine "Product Name" & vbTab & objItem.ProductName
        objTextFile2.WriteLine "Status Information" & vbTab & objItem.StatusInfo
        objTextFile2.WriteLine
    Next

    SecTitulo = "*** Video Adapter Information "
    objTextFile2.WriteLine SecTitulo & String(3,"*") & vbCrLf

    Set colItems = objWMIService.ExecQuery _
     ("Select * from Win32_DisplayControllerConfiguration")

    For Each objItem in colItems
        Wscript.StdOut.Write(".")
        objTextFile2.WriteLine "Bits Per Pixel" & vbTab & objItem.BitsPerPixel
        objTextFile2.WriteLine "Color Planes" & vbTab & objItem.ColorPlanes
        objTextFile2.WriteLine "Device Entries in a Color vbTable" & _
         vbTab & objItem.DeviceEntriesInAColorvbTable
        objTextFile2.WriteLine "Device Specific Pens" & vbTab & objItem.DeviceSpecificPens
        objTextFile2.WriteLine "Horizontal Resolution" & vbTab & objItem.HorizontalResolution
        objTextFile2.WriteLine "Name" & vbTab & objItem.Name
        objTextFile2.WriteLine "Refresh Rate" & vbTab & objItem.RefreshRate
        objTextFile2.WriteLine "Setting ID" & vbTab & objItem.SettingID
        objTextFile2.WriteLine "Vertical Resolution" & vbTab & objItem.VerticalResolution
        objTextFile2.WriteLine "Video Mode" & vbTab & objItem.VideoMode
        objTextFile2.WriteLine
    Next

    SecTitulo = "*** Drive Types "
    objTextFile2.WriteLine SecTitulo & String(3,"*") & vbCrLf

    Set colDisks = objWMIService.ExecQuery _
     ("Select * from Win32_LogicalDisk")

    DriveT=""

    For Each objDisk in colDisks
        Wscript.StdOut.Write(".")
        Select Case objDisk.DriveType
            Case 1
            DriveT="No root directory. Drive type could not be " _
                & "determined."
            Case 2
            DriveT="Removable drive."
    
            Case 3
            DriveT="Local hard disk."
        
            Case 4
            DriveT="Network disk."      

            Case 5
            DriveT="Compact disk."      
    
            Case 6
            DriveT="RAM disk."   

            Case Else
            DriveT="Drive type could not be determined."
        End Select
        objTextFile2.WriteLine "DeviceID"& vbTab & objDisk.DeviceID & " " & DriveT 
    Next
    objTextFile2.WriteLine

    SecTitulo = "*** Physical Disk Properties "
    objTextFile2.WriteLine SecTitulo & String(3,"*") & vbCrLf

    Set colDiskDrives = objWMIService.ExecQuery _    
     ("Select * from Win32_DiskDrive")

    For each objDiskDrive in colDiskDrives
        Wscript.StdOut.Write(".")
        objTextFile2.WriteLine "Caption" & vbTab &  objDiskDrive.Caption
        objTextFile2.WriteLine "Device ID" & vbTab &  objDiskDrive.DeviceID
        objTextFile2.WriteLine "Index" & vbTab &  objDiskDrive.Index
        objTextFile2.WriteLine "Interface Type" & vbTab & objDiskDrive.InterfaceType
        objTextFile2.WriteLine "Manufacturer" & vbTab & objDiskDrive.Manufacturer
        objTextFile2.WriteLine "Media Loaded" & vbTab  & objDiskDrive.MediaLoaded
        objTextFile2.WriteLine "Media Type" & vbTab &  objDiskDrive.MediaType
        objTextFile2.WriteLine "Model" & vbTab &  objDiskDrive.Model
        objTextFile2.WriteLine "Name" & vbTab &  objDiskDrive.Name
        objTextFile2.WriteLine "Partitions" & vbTab & objDiskDrive.Partitions
        objTextFile2.WriteLine "SCSI Bus" & vbTab &  objDiskDrive.SCSIBus
        objTextFile2.WriteLine "SCSI Logical Unit" & vbTab &  _
         objDiskDrive.SCSILogicalUnit
        objTextFile2.WriteLine "SCSI Port" & vbTab &  objDiskDrive.SCSIPort
        objTextFile2.WriteLine "SCSI TargetId" & vbTab &  objDiskDrive.SCSITargetId    
        objTextFile2.WriteLine "Sectors Per Track" & vbTab &  _
         objDiskDrive.SectorsPerTrack        
        objTextFile2.WriteLine "Signature" & vbTab &  objDiskDrive.Signature          
        objTextFile2.WriteLine "Size" & vbTab &  objDiskDrive.Size     
        objTextFile2.WriteLine "Status" & vbTab &  objDiskDrive.Status         
        objTextFile2.WriteLine "Total Cylinders" & vbTab &  _
         objDiskDrive.TotalCylinders         
        objTextFile2.WriteLine "Total Heads" & vbTab &  objDiskDrive.TotalHeads    
        objTextFile2.WriteLine "Total Sectors" & vbTab &  objDiskDrive.TotalSectors
        objTextFile2.WriteLine "Total Tracks" & vbTab &  objDiskDrive.TotalTracks
        objTextFile2.WriteLine "Tracks Per Cylinder" & vbTab &  _
         objDiskDrive.TracksPerCylinder
        objTextFile2.WriteLine
    Next

    SecTitulo = "*** Logical Disk Drive Properties "
    objTextFile2.WriteLine SecTitulo & String(3,"*") & vbCrLf

    Set colDisks = objWMIService.ExecQuery _
     ("Select * from Win32_LogicalDisk")

    For each objDisk in colDisks
        Wscript.StdOut.Write(".")
        objTextFile2.WriteLine "Compressed" & vbTab & objDisk.Compressed  
        objTextFile2.WriteLine "Description" & vbTab & objDisk.Description       
        objTextFile2.WriteLine "DeviceID" & vbTab & objDisk.DeviceID      
        objTextFile2.WriteLine "DriveType" & vbTab & objDisk.DriveType    
        objTextFile2.WriteLine "FileSystem" & vbTab & objDisk.FileSystem  
        objTextFile2.WriteLine "FreeSpace" & vbTab & objDisk.FreeSpace    
        objTextFile2.WriteLine "Name" & vbTab & objDisk.Name      
        objTextFile2.WriteLine "Size" & vbTab & objDisk.Size      
        objTextFile2.WriteLine "VolumeName" & vbTab & objDisk.VolumeName  
        objTextFile2.WriteLine "VolumeSerialNumber" & _
         vbTab & objDisk.VolumeSerialNumber      
        objTextFile2.WriteLine
    Next

    SecTitulo = "*** Network Adapter "
    objTextFile2.WriteLine SecTitulo & String(3,"*") & vbCrLf

    Set colAdapters = objWMIService.ExecQuery _
     ("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
 
    n = 1
 
    For Each objAdapter in colAdapters
        Wscript.StdOut.Write(".")
        objTextFile2.WriteLine "Description" & vbTab & objAdapter.Description
        objTextFile2.WriteLine "Physical (MAC) address" & vbTab & objAdapter.MACAddress
        objTextFile2.WriteLine "Host name:              " & vbTab & objAdapter.DNSHostName
 
        If Not IsNull(objAdapter.IPAddress) Then
           For i = 0 To UBound(objAdapter.IPAddress)
              objTextFile2.WriteLine "IP address:" & vbTab & objAdapter.IPAddress(i)
           Next
        End If
 
        If Not IsNull(objAdapter.IPSubnet) Then
           For i = 0 To UBound(objAdapter.IPSubnet)
              objTextFile2.WriteLine "Subnet:" & vbTab & objAdapter.IPSubnet(i)
           Next
        End If
 
        If Not IsNull(objAdapter.DefaultIPGateway) Then
           For i = 0 To UBound(objAdapter.DefaultIPGateway)
              objTextFile2.WriteLine "Default gateway:" & vbTab & _
               objAdapter.DefaultIPGateway(i)
           Next
        End If
 
        If Not IsNull(objAdapter.DNSServerSearchOrder) Then
           For i = 0 To UBound(objAdapter.DNSServerSearchOrder)
              objTextFile2.WriteLine "DNS " & i & ":" & vbTab & _
               objAdapter.DNSServerSearchOrder(i)
           Next
        End If
 
        objTextFile2.WriteLine "DNS domain:" & vbTab & objAdapter.DNSDomain
 
        If Not IsNull(objAdapter.DNSDomainSuffixSearchOrder) Then
           For i = 0 To UBound(objAdapter.DNSDomainSuffixSearchOrder)
              objTextFile2.WriteLine "DNS suffix search list:" & vbTab & _
               objAdapter.DNSDomainSuffixSearchOrder(i)
           Next
        End If
 
        objTextFile2.WriteLine "DHCP enabled" & vbTab & objAdapter.DHCPEnabled
        objTextFile2.WriteLine "DHCP server:  " & vbTab & objAdapter.DHCPServer
 
        If Not IsNull(objAdapter.DHCPLeaseObtained) Then
           utcLeaseObtained = objAdapter.DHCPLeaseObtained
           strLeaseObtained = WMIDateStringToDate(utcLeaseObtained)
        Else
           strLeaseObtained = ""
        End If
        objTextFile2.WriteLine "DHCP lease obtained" & vbTab & strLeaseObtained
 
        If Not IsNull(objAdapter.DHCPLeaseExpires) Then
           utcLeaseExpires = objAdapter.DHCPLeaseExpires
           strLeaseExpires = WMIDateStringToDate(utcLeaseExpires)
        Else
           strLeaseExpires = ""
        End If
        objTextFile2.WriteLine "DHCP lease expires:" & vbTab & strLeaseExpires
        objTextFile2.WriteLine "Primary WINS server:" & vbTab & objAdapter.WINSPrimaryServer
        objTextFile2.WriteLine "Secondary WINS server" & vbTab & objAdapter.WINSSecondaryServer
        objTextFile2.WriteLine
 
        n = n + 1
 
    Next

    SecTitulo = "*** Network Shares "
    objTextFile2.WriteLine SecTitulo & String(3,"*") & vbCrLf

    Set colShares = objWMIService.ExecQuery("Select * from Win32_Share")

    For each objShare in colShares
        Wscript.StdOut.Write(".")
        objTextFile2.WriteLine "Allow Maximum" & vbTab & objShare.AllowMaximum   
        objTextFile2.WriteLine "Caption" & vbTab & objShare.Caption   
        objTextFile2.WriteLine "Maximum Allowed" & vbTab & objShare.MaximumAllowed
        objTextFile2.WriteLine "Name" & vbTab & objShare.Name   
        objTextFile2.WriteLine "Path" & vbTab & objShare.Path   
        objTextFile2.WriteLine "Type" & vbTab & objShare.Type   
        objTextFile2.WriteLine
    Next

    SecTitulo = "*** Printers "
    objTextFile2.WriteLine SecTitulo & String(3,"*") & vbCrLf

    Set colInstalledPrinters =  objWMIService.ExecQuery _
     ("Select * from Win32_Printer")

    For Each objPrinter in colInstalledPrinters
        Wscript.StdOut.Write(".")
        objTextFile2.WriteLine "Name" & vbTab & objPrinter.Name
        objTextFile2.WriteLine "Location" & vbTab & objPrinter.Location
        objTextFile2.WriteLine "Default" & vbTab & objPrinter.Default
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

Function WMIDateStringToDate(utcDate)
   WMIDateStringToDate = CDate(Mid(utcDate, 5, 2)  & "/" & _
    Mid(utcDate, 7, 2)  & "/" & _
    Left(utcDate, 4)    & " " & _
    Mid (utcDate, 9, 2) & ":" & _
    Mid(utcDate, 11, 2) & ":" & _
    Mid(utcDate, 13, 2))
End Function
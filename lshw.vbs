On Error Resume Next

Wscript.Echo vbCrLf & "Running. Wait please... " & vbCrLf

const OutFile1 = "lshw.tsv"
step = 0
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.CreateTextFile(OutFile1, True)

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

SecTitulo = "*** BIOS "
step = step + 1
Wscript.Echo "STEP " & step & " " & SecTitulo
objTextFile.WriteLine SecTitulo & String(3,"*") & vbCrLf

Set colBIOS = objWMIService.ExecQuery _
    ("Select * from Win32_BIOS")

For each objBIOS in colBIOS
    objTextFile.WriteLine "Build Number" & vbTab & objBIOS.BuildNumber
    objTextFile.WriteLine "Current Language" & vbTab & objBIOS.CurrentLanguage
    objTextFile.WriteLine "Installable Languages" & vbTab & objBIOS.InstallableLanguages
    objTextFile.WriteLine "Manufacturer" & vbTab & objBIOS.Manufacturer
    objTextFile.WriteLine "Name" & vbTab & objBIOS.Name
    objTextFile.WriteLine "Primary BIOS" & vbTab & objBIOS.PrimaryBIOS
    objTextFile.WriteLine "Release Date" & vbTab & objBIOS.ReleaseDate
    objTextFile.WriteLine "Serial Number" & vbTab & objBIOS.SerialNumber
    objTextFile.WriteLine "SMBIOS Version" & vbTab & objBIOS.SMBIOSBIOSVersion
    objTextFile.WriteLine "SMBIOS Major Version" & vbTab & objBIOS.SMBIOSMajorVersion
    objTextFile.WriteLine "SMBIOS Minor Version" & vbTab & objBIOS.SMBIOSMinorVersion
    objTextFile.WriteLine "SMBIOS Present" & vbTab & objBIOS.SMBIOSPresent
    objTextFile.WriteLine "Status" & vbTab & objBIOS.Status
    objTextFile.WriteLine "Version" & vbTab & objBIOS.Version
    objTextFile.WriteLine
Next

SecTitulo = "*** Processor "
step = step + 1
Wscript.Echo "STEP " & step & " " & SecTitulo
objTextFile.WriteLine SecTitulo & String(3,"*") & vbCrLf

Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor")

For Each objItem in colItems
    objTextFile.WriteLine "Address Width" & vbTab & objItem.AddressWidth
    objTextFile.WriteLine "Architecture" & vbTab & objItem.Architecture
    objTextFile.WriteLine "Availability" & vbTab & objItem.Availability
    objTextFile.WriteLine "CPU Status" & vbTab & objItem.CpuStatus
    objTextFile.WriteLine "Current Clock Speed" & vbTab & objItem.CurrentClockSpeed
    objTextFile.WriteLine "Data Width" & vbTab & objItem.DataWidth
    objTextFile.WriteLine "Description" & vbTab & objItem.Description
    objTextFile.WriteLine "Device ID" & vbTab & objItem.DeviceID
    objTextFile.WriteLine "External Clock" & vbTab & objItem.ExtClock
    objTextFile.WriteLine "Family" & vbTab & objItem.Family
    objTextFile.WriteLine "L2 Cache Size" & vbTab & objItem.L2CacheSize
    objTextFile.WriteLine "L2 Cache Speed" & vbTab & objItem.L2CacheSpeed
    objTextFile.WriteLine "Level" & vbTab & objItem.Level
    objTextFile.WriteLine "Load Percentage" & vbTab & objItem.LoadPercentage
    objTextFile.WriteLine "Manufacturer" & vbTab & objItem.Manufacturer
    objTextFile.WriteLine "Maximum Clock Speed" & vbTab & objItem.MaxClockSpeed
    objTextFile.WriteLine "Name" & vbTab & objItem.Name
    objTextFile.WriteLine "Processor ID" & vbTab & objItem.ProcessorId
    objTextFile.WriteLine "Processor Type" & vbTab & objItem.ProcessorType
    objTextFile.WriteLine "Revision" & vbTab & objItem.Revision
    objTextFile.WriteLine "Socket Designation" & vbTab & objItem.SocketDesignation
    objTextFile.WriteLine "Status Information" & vbTab & objItem.StatusInfo
    objTextFile.WriteLine "Stepping" & vbTab & objItem.Stepping
    objTextFile.WriteLine "Unique Id" & vbTab & objItem.UniqueId
    objTextFile.WriteLine "Upgrade Method" & vbTab & objItem.UpgradeMethod
    objTextFile.WriteLine "Version" & vbTab & objItem.Version
    objTextFile.WriteLine "Voltage Caps" & vbTab & objItem.VoltageCaps
    objTextFile.WriteLine
Next

SecTitulo = "*** Motherboard Device "
step = step + 1
Wscript.Echo "STEP " & step & " " & SecTitulo
objTextFile.WriteLine SecTitulo & String(3,"*") & vbCrLf

Set colItems = objWMIService.ExecQuery("Select * from Win32_MotherboardDevice")

For Each objItem in colItems
    objTextFile.WriteLine "Device ID" & vbTab & objItem.DeviceID
    objTextFile.WriteLine "Primary Bus Type" & vbTab & objItem.PrimaryBusType
    objTextFile.WriteLine "Secondary Bus Type" & vbTab & objItem.SecondaryBusType
    objTextFile.WriteLine
Next

SecTitulo = "*** Onboard Device "
step = step + 1
Wscript.Echo "STEP " & step & " " & SecTitulo
objTextFile.WriteLine SecTitulo & String(3,"*") & vbCrLf

Set colItems = objWMIService.ExecQuery("Select * from Win32_OnBoardDevice")

For Each objItem in colItems
    objTextFile.WriteLine "Description" & vbTab & objItem.Description
    objTextFile.WriteLine "Device Type" & vbTab & objItem.DeviceType
    objTextFile.WriteLine "Model" & vbTab & objItem.Model
    objTextFile.WriteLine "Name" & vbTab & objItem.Name
    objTextFile.WriteLine "Tag" & vbTab & objItem.Tag
    objTextFile.WriteLine "Version" & vbTab & objItem.Version
    objTextFile.WriteLine
Next

SecTitulo = "*** Phisical Memory "
step = step + 1
Wscript.Echo "STEP " & step & " " & SecTitulo
objTextFile.WriteLine SecTitulo & String(3,"*") & vbCrLf

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_PhysicalMemoryArray")

For Each objItem in colItems
    objTextFile.WriteLine "Description" & vbTab & objItem.Description
    objTextFile.WriteLine "Maximum Capacity" & vbTab & objItem.MaxCapacity
    objTextFile.WriteLine "Memory Devices" & vbTab & objItem.MemoryDevices
    objTextFile.WriteLine "Memory Error Correction" & vbTab & objItem.MemoryErrorCorrection
    objTextFile.WriteLine
Next

SecTitulo = "*** Memory Configuration "
step = step + 1
Wscript.Echo "STEP " & step & " " & SecTitulo
objTextFile.WriteLine SecTitulo & String(3,"*") & vbCrLf

Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory")

For Each objItem in colItems
    objTextFile.WriteLine "Bank Label" & vbTab & objItem.BankLabel
    objTextFile.WriteLine "Capacity" & vbTab & objItem.Capacity
    objTextFile.WriteLine "Data Width" & vbTab & objItem.DataWidth
    objTextFile.WriteLine "Description" & vbTab & objItem.Description
    objTextFile.WriteLine "Device Locator" & vbTab & objItem.DeviceLocator
    objTextFile.WriteLine "Form Factor" & vbTab & objItem.FormFactor
    objTextFile.WriteLine "Hot Swappable" & vbTab & objItem.HotSwappable
    objTextFile.WriteLine "Manufacturer" & vbTab & objItem.Manufacturer
    objTextFile.WriteLine "Memory Type" & vbTab & objItem.MemoryType
    objTextFile.WriteLine "Name" & vbTab & objItem.Name
    objTextFile.WriteLine "Part Number" & vbTab & objItem.PartNumber
    objTextFile.WriteLine "Position In Row" & vbTab & objItem.PositionInRow
    objTextFile.WriteLine "Speed" & vbTab & objItem.Speed
    objTextFile.WriteLine "Tag" & vbTab & objItem.Tag
    objTextFile.WriteLine "Type Detail" & vbTab & objItem.TypeDetail
    objTextFile.WriteLine
Next

SecTitulo = "*** Sound Card Properties "
step = step + 1
Wscript.Echo "STEP " & step & " " & SecTitulo
objTextFile.WriteLine SecTitulo & String(3,"*") & vbCrLf

Set colItems = objWMIService.ExecQuery("Select * from Win32_SoundDevice")

For Each objItem in colItems
    objTextFile.WriteLine "Description" & vbTab & objItem.Description
    objTextFile.WriteLine "Device ID" & vbTab & objItem.DeviceID
    objTextFile.WriteLine "DMA Buffer Size" & vbTab & objItem.DMABufferSize
    objTextFile.WriteLine "Manufacturer" & vbTab & objItem.Manufacturer
    objTextFile.WriteLine "MPU 401 Address" & vbTab & objItem.MPU401Address
    objTextFile.WriteLine "Name" & vbTab & objItem.Name
    objTextFile.WriteLine "PNP Device ID" & vbTab & objItem.PNPDeviceID
    objTextFile.WriteLine "Product Name" & vbTab & objItem.ProductName
    objTextFile.WriteLine "Status Information" & vbTab & objItem.StatusInfo
    objTextFile.WriteLine
Next

SecTitulo = "*** Video Adapter Information "
step = step + 1
Wscript.Echo "STEP " & step & " " & SecTitulo
objTextFile.WriteLine SecTitulo & String(3,"*") & vbCrLf

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_DisplayControllerConfiguration")

For Each objItem in colItems
    objTextFile.WriteLine "Bits Per Pixel" & vbTab & objItem.BitsPerPixel
    objTextFile.WriteLine "Color Planes" & vbTab & objItem.ColorPlanes
    objTextFile.WriteLine "Device Entries in a Color vbTable" & _
         vbTab & objItem.DeviceEntriesInAColorvbTable
    objTextFile.WriteLine "Device Specific Pens" & vbTab & objItem.DeviceSpecificPens
    objTextFile.WriteLine "Horizontal Resolution" & vbTab & objItem.HorizontalResolution
    objTextFile.WriteLine "Name" & vbTab & objItem.Name
    objTextFile.WriteLine "Refresh Rate" & vbTab & objItem.RefreshRate
    objTextFile.WriteLine "Setting ID" & vbTab & objItem.SettingID
    objTextFile.WriteLine "Vertical Resolution" & vbTab & objItem.VerticalResolution
    objTextFile.WriteLine "Video Mode" & vbTab & objItem.VideoMode
    objTextFile.WriteLine
Next


SecTitulo = "*** Drive Types "
step = step + 1
Wscript.Echo "STEP " & step & " " & SecTitulo
objTextFile.WriteLine SecTitulo & String(3,"*") & vbCrLf

Set colDisks = objWMIService.ExecQuery _
    ("Select * from Win32_LogicalDisk")

DriveT=""

For Each objDisk in colDisks
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
    objTextFile.WriteLine "DeviceID"& vbTab & objDisk.DeviceID & " " & DriveT 
Next
objTextFile.WriteLine

SecTitulo = "*** Physical Disk Properties "
step = step + 1
Wscript.Echo "STEP " & step & " " & SecTitulo
objTextFile.WriteLine SecTitulo & String(3,"*") & vbCrLf

Set colDiskDrives = objWMIService.ExecQuery _    
    ("Select * from Win32_DiskDrive")

For each objDiskDrive in colDiskDrives
    objTextFile.WriteLine "Caption" & vbTab &  objDiskDrive.Caption
    objTextFile.WriteLine "Device ID" & vbTab &  objDiskDrive.DeviceID
    objTextFile.WriteLine "Index" & vbTab &  objDiskDrive.Index
    objTextFile.WriteLine "Interface Type" & vbTab & objDiskDrive.InterfaceType
    objTextFile.WriteLine "Manufacturer" & vbTab & objDiskDrive.Manufacturer
    objTextFile.WriteLine "Media Loaded" & vbTab  & objDiskDrive.MediaLoaded
    objTextFile.WriteLine "Media Type" & vbTab &  objDiskDrive.MediaType
    objTextFile.WriteLine "Model" & vbTab &  objDiskDrive.Model
    objTextFile.WriteLine "Name" & vbTab &  objDiskDrive.Name
    objTextFile.WriteLine "Partitions" & vbTab & objDiskDrive.Partitions
    objTextFile.WriteLine "SCSI Bus" & vbTab &  objDiskDrive.SCSIBus
    objTextFile.WriteLine "SCSI Logical Unit" & vbTab &  _
        objDiskDrive.SCSILogicalUnit
    objTextFile.WriteLine "SCSI Port" & vbTab &  objDiskDrive.SCSIPort
    objTextFile.WriteLine "SCSI TargetId" & vbTab &  objDiskDrive.SCSITargetId    
    objTextFile.WriteLine "Sectors Per Track" & vbTab &  _
        objDiskDrive.SectorsPerTrack        
    objTextFile.WriteLine "Signature" & vbTab &  objDiskDrive.Signature          
    objTextFile.WriteLine "Size" & vbTab &  objDiskDrive.Size     
    objTextFile.WriteLine "Status" & vbTab &  objDiskDrive.Status         
    objTextFile.WriteLine "Total Cylinders" & vbTab &  _
        objDiskDrive.TotalCylinders         
    objTextFile.WriteLine "Total Heads" & vbTab &  objDiskDrive.TotalHeads    
    objTextFile.WriteLine "Total Sectors" & vbTab &  objDiskDrive.TotalSectors
    objTextFile.WriteLine "Total Tracks" & vbTab &  objDiskDrive.TotalTracks
    objTextFile.WriteLine "Tracks Per Cylinder" & vbTab &  _
        objDiskDrive.TracksPerCylinder
    objTextFile.WriteLine
Next

SecTitulo = "*** Logical Disk Drive Properties "
step = step + 1
Wscript.Echo "STEP " & step & " " & SecTitulo
objTextFile.WriteLine SecTitulo & String(3,"*") & vbCrLf

Set colDisks = objWMIService.ExecQuery _
    ("Select * from Win32_LogicalDisk")

For each objDisk in colDisks
    objTextFile.WriteLine "Compressed" & vbTab & objDisk.Compressed  
    objTextFile.WriteLine "Description" & vbTab & objDisk.Description       
    objTextFile.WriteLine "DeviceID" & vbTab & objDisk.DeviceID      
    objTextFile.WriteLine "DriveType" & vbTab & objDisk.DriveType    
    objTextFile.WriteLine "FileSystem" & vbTab & objDisk.FileSystem  
    objTextFile.WriteLine "FreeSpace" & vbTab & objDisk.FreeSpace    
    objTextFile.WriteLine "Name" & vbTab & objDisk.Name      
    objTextFile.WriteLine "Size" & vbTab & objDisk.Size      
    objTextFile.WriteLine "VolumeName" & vbTab & objDisk.VolumeName  
    objTextFile.WriteLine "VolumeSerialNumber" & _
         vbTab & objDisk.VolumeSerialNumber      
    objTextFile.WriteLine
Next

SecTitulo = "*** Network Adapter "
step = step + 1
Wscript.Echo "STEP " & step & " " & SecTitulo
objTextFile.WriteLine SecTitulo & String(3,"*") & vbCrLf

Set colAdapters = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
 
n = 1
 
For Each objAdapter in colAdapters
    objTextFile.WriteLine "Description" & vbTab & objAdapter.Description
 
    objTextFile.WriteLine "Physical (MAC) address" & vbTab & objAdapter.MACAddress
    objTextFile.WriteLine "Host name:              " & vbTab & objAdapter.DNSHostName
 
    If Not IsNull(objAdapter.IPAddress) Then
       For i = 0 To UBound(objAdapter.IPAddress)
          objTextFile.WriteLine "IP address:" & vbTab & objAdapter.IPAddress(i)
       Next
    End If
 
    If Not IsNull(objAdapter.IPSubnet) Then
       For i = 0 To UBound(objAdapter.IPSubnet)
          objTextFile.WriteLine "Subnet:" & vbTab & objAdapter.IPSubnet(i)
       Next
    End If
 
    If Not IsNull(objAdapter.DefaultIPGateway) Then
       For i = 0 To UBound(objAdapter.DefaultIPGateway)
          objTextFile.WriteLine "Default gateway:" & vbTab & _
              objAdapter.DefaultIPGateway(i)
       Next
    End If
 
    If Not IsNull(objAdapter.DNSServerSearchOrder) Then
       For i = 0 To UBound(objAdapter.DNSServerSearchOrder)
          objTextFile.WriteLine "DNS " & i & ":" & vbTab & objAdapter.DNSServerSearchOrder(i)
       Next
    End If
 
    objTextFile.WriteLine "DNS domain:" & vbTab & objAdapter.DNSDomain
 
    If Not IsNull(objAdapter.DNSDomainSuffixSearchOrder) Then
       For i = 0 To UBound(objAdapter.DNSDomainSuffixSearchOrder)
          objTextFile.WriteLine "DNS suffix search list:" & vbTab & _
              objAdapter.DNSDomainSuffixSearchOrder(i)
       Next
    End If
 
    objTextFile.WriteLine "DHCP enabled" & vbTab & objAdapter.DHCPEnabled
    objTextFile.WriteLine "DHCP server:  " & vbTab & objAdapter.DHCPServer
 
    If Not IsNull(objAdapter.DHCPLeaseObtained) Then
       utcLeaseObtained = objAdapter.DHCPLeaseObtained
       strLeaseObtained = WMIDateStringToDate(utcLeaseObtained)
    Else
       strLeaseObtained = ""
    End If
    objTextFile.WriteLine "DHCP lease obtained" & vbTab & strLeaseObtained
 
    If Not IsNull(objAdapter.DHCPLeaseExpires) Then
       utcLeaseExpires = objAdapter.DHCPLeaseExpires
       strLeaseExpires = WMIDateStringToDate(utcLeaseExpires)
    Else
       strLeaseExpires = ""
    End If
    objTextFile.WriteLine "DHCP lease expires:" & vbTab & strLeaseExpires
 
    objTextFile.WriteLine "Primary WINS server:" & vbTab & objAdapter.WINSPrimaryServer
    objTextFile.WriteLine "Secondary WINS server" & vbTab & objAdapter.WINSSecondaryServer
    objTextFile.WriteLine
 
    n = n + 1
 
Next

SecTitulo = "*** Network Shares "
step = step + 1
Wscript.Echo "STEP " & step & " " & SecTitulo
objTextFile.WriteLine SecTitulo & String(3,"*") & vbCrLf

Set colShares = objWMIService.ExecQuery("Select * from Win32_Share")

For each objShare in colShares
    objTextFile.WriteLine "Allow Maximum" & vbTab & objShare.AllowMaximum   
    objTextFile.WriteLine "Caption" & vbTab & objShare.Caption   
    objTextFile.WriteLine "Maximum Allowed" & vbTab & objShare.MaximumAllowed
    objTextFile.WriteLine "Name" & vbTab & objShare.Name   
    objTextFile.WriteLine "Path" & vbTab & objShare.Path   
    objTextFile.WriteLine "Type" & vbTab & objShare.Type   
    objTextFile.WriteLine
Next

SecTitulo = "*** Printers "
step = step + 1
Wscript.Echo "STEP " & step & " " & SecTitulo
objTextFile.WriteLine SecTitulo & String(3,"*") & vbCrLf

Set colInstalledPrinters =  objWMIService.ExecQuery _
    ("Select * from Win32_Printer")

For Each objPrinter in colInstalledPrinters
    objTextFile.WriteLine "Name" & vbTab & objPrinter.Name
    objTextFile.WriteLine "Location" & vbTab & objPrinter.Location
    objTextFile.WriteLine "Default" & vbTab & objPrinter.Default
Next
objTextFile.Close

Set colOSes = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
For Each objOS in colOSes
    OutFile2 = objOS.CSName & "_hw.tsv"
Next

objFSO.DeleteFile(OutFile2)
objFSO.MoveFile OutFile1, OutFile2
Wscript.Echo vbCrLf & "Finish. Result on " & OutFile2 & vbCrLf

Function WMIDateStringToDate(utcDate)
   WMIDateStringToDate = CDate(Mid(utcDate, 5, 2)  & "/" & _
    Mid(utcDate, 7, 2)  & "/" & _
    Left(utcDate, 4)    & " " & _
    Mid (utcDate, 9, 2) & ":" & _
    Mid(utcDate, 11, 2) & ":" & _
    Mid(utcDate, 13, 2))
End Function
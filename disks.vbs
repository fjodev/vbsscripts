strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colDiskPartitions = objWMIService.ExecQuery _
    ("Select * from Win32_DiskPartition")
separador = "------------------------------------------------------"
For each objPartition in colDiskPartitions
    Wscript.Echo separador
    Wscript.Echo "Name               : " & objPartition.Name
    Wscript.Echo "Device ID          : " & objPartition.DeviceID       
    Wscript.Echo "Description        : " & objPartition.Description
    Wscript.Echo "Type               : " & objPartition.Type 
    Wscript.Echo "Disk Index         : " & objPartition.DiskIndex     
    Wscript.Echo "Index              : " & objPartition.Index       
    Wscript.Echo "Size               : " & objPartition.Size 
    Wscript.Echo "Block Size         : " & objPartition.BlockSize     
    Wscript.Echo "Bootable           : " & objPartition.Bootable 
    Wscript.Echo "Boot Partition     : " & objPartition.BootPartition  
    Wscript.Echo "Number Of Blocks   : " & _
        objPartition.NumberOfBlocks     
    Wscript.Echo "Primary Partition  : " & _
        objPartition.PrimaryPartition   
    Wscript.Echo "Starting Offset    : " & _
        objPartition.StartingOffset     

Next
	
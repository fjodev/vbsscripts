strComputer = "."
strDirChk = Wscript.Arguments.Item(0)
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colSubfolders = objWMIService.ExecQuery _
    ("Associators of {Win32_Directory.Name='c:\foliveira-docs\scripts'} " _
        & "Where AssocClass = Win32_Subdirectory " _
            & "ResultRole = PartComponent")

For Each objFolder in colSubfolders
    Wscript.Echo objFolder.Name
Next
	
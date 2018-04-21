strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_UTCTime")

wscript.echo

For Each objItem in colItems
    Wscript.Echo "Ano           : " & objItem.Year
    Wscript.Echo "Mês           : " & objItem.Month
    Wscript.Echo "Dia           : " & objItem.Day & vbCrLf
    Wscript.Echo "Trimestre     : " & objItem.Quarter 
    Wscript.Echo "Semana no mês : " & objItem.WeekInMonth
    Wscript.Echo "Dia da semana : " & objItem.DayOfWeek & vbCrLf
    Wscript.Echo "Hora          : " & objItem.Hour
    Wscript.Echo "Minuto        : " & objItem.Minute
    Wscript.Echo "Segundo       : " & objItem.Second & vbCrLf
Next
	
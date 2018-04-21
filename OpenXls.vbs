Set objExcel = CreateObject("Excel.Application") 

objExcel.Visible = TRUE 
objExcel.DisplayAlerts = TRUE

Set objWorkbook = objExcel.Workbooks.Open("C:\Users\Fernando\Documents\Geral\CredCarro.xls",,,,"6891856FjO")

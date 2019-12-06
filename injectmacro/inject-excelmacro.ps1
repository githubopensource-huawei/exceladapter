$excelapp=New-Object -comobject excel.application
$excelapp.DisplayAlerts = $false
$excelapp.EnableEvents = $false
$x = Split-Path -Parent $MyInvocation.MyCommand.Definition
$wb = $excelapp.workbooks.open($x + "\inject.xlsm")
$excelapp.run("doInjectMacros")
$wb.close()
$excelapp.quit()
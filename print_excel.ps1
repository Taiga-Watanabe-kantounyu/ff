$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Open($args[0])
$excel.Visible = $false
$ws = $workbook.Worksheets.Item(1)
$ws.PrintOut(1, 1, 1, $false, '東京事務所 bizhub C360i')
while ($excel.ActivePrinter -ne $null -and $excel.Ready -eq $false) { Start-Sleep -Seconds 1 }
Write-Output 'OK'
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

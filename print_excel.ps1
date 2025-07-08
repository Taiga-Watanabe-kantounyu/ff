param(
    [string]$excelFile,
    [string]$printerName
)

if (-not (Test-Path $excelFile)) {
    Write-Error "ファイルが存在しません: $excelFile"
    exit 1
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

try {
    $workbook = $excel.Workbooks.Open($excelFile)
    if ($printerName -and $printerName.Trim() -ne "") {
        $workbook.PrintOut(ActivePrinter=$printerName)
    } else {
        $workbook.PrintOut()
    }
    $workbook.Close($false)
} catch {
    Write-Error "印刷中にエラーが発生しました: $_"
    exit 2
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

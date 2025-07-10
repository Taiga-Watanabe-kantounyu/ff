param (
    [string]$excelFilePath,
    [string]$printerName
)

Write-Host "Excel印刷スクリプト開始"
Write-Host "ファイル: $excelFilePath"
Write-Host "プリンタ: $printerName"

function Get-MyPrinterOnPort {
    param($PrinterName)
    $portName = $null
    $keyPath = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts'
    if (Test-Path $keyPath) {
        $regKey = Get-Item $keyPath
        $printerPort = $regKey.GetValue($PrinterName)
        if ($printerPort -ne $null) {
            $portinfos = $printerPort -split ","
            if ($portinfos.length -gt 1) {
                $portName = $portinfos[1]
                return "$PrinterName on $portName"
            } else {
                Write-Host "【エラー】レジストリにポート名がありません: $PrinterName [$printerPort]"
            }
        } else {
            Write-Host "【エラー】レジストリにエントリーがありません: $keyPath $PrinterName"
        }
    } else {
        Write-Host "【エラー】レジストリキーがありません: $keyPath"
    }
    return $null
}

# Excel起動
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

# ファイルを開く
$workbook = $excel.Workbooks.Open($excelFilePath)
# プリンター設定（存在確認付き）
if ($printerName -and $printerName.Trim() -ne "") {
    try {
        $availablePrinters = Get-WmiObject -Class Win32_Printer | Select-Object -ExpandProperty Name
        if ($availablePrinters -match [regex]::Escape($printerName)) {
                $formattedPrinterName = Get-MyPrinterOnPort $printerName
        if ($formattedPrinterName) {
                    $excel.ActivePrinter = $formattedPrinterName
                    Write-Host "プリンターを設定しました: $formattedPrinterName"
        } else {
                    Write-Warning "プリンター名のフォーマットに失敗しました"
        }
            } else {
                Write-Warning "指定されたプリンターが見つかりません: $printerName"
                Write-Host "利用可能なプリンター:"
                $availablePrinters | ForEach-Object { Write-Host "  - $_" }
                throw "指定されたプリンターが見つかりません: $printerName"
            }
        } catch {
            Write-Warning "プリンター設定でエラーが発生しました: $($_.Exception.Message)"
        }
}

# 印刷実行（非同期で印刷命令だけ出して、すぐ次へ）
$faxProcess = Start-Process -FilePath "C:\Users\7342\Documents\ff_2\sendfax.ahk" -PassThru
$workbook.PrintOut()

Write-Host "印刷命令を送信しました"

# Workbookを閉じる（保存せず）
$workbook.Close($false)

# Excelを終了
$excel.Quit()

# COMオブジェクトの解放
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "Excelを閉じました。スクリプト終了。"

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8

$file = Join-Path $PSScriptRoot '基本補助費概算表格式.xlsm'

Write-Host ''
Write-Host '=== 基本補助費概算表 一鍵計算 ===' -ForegroundColor Cyan
Write-Host ''

if (-not (Test-Path $file)) {
    Write-Host "找不到檔案：$file" -ForegroundColor Red
    Read-Host '按 Enter 結束'
    exit 1
}

# 檔案若正被 Excel 開著，先請使用者關閉，避免存檔失敗
try {
    $fs = [System.IO.File]::Open($file, 'Open', 'ReadWrite', 'None')
    $fs.Close()
} catch {
    Write-Host '「基本補助費概算表格式.xlsm」目前正被開啟中。' -ForegroundColor Yellow
    Write-Host '請先在 Excel 關閉該檔案，再重新執行本程式。' -ForegroundColor Yellow
    Read-Host '按 Enter 結束'
    exit 1
}

Write-Host '正在讀取「競賽資訊」並重算，請稍候...' -ForegroundColor Gray

$xl = $null
$wb = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false

    $wb = $xl.Workbooks.Open($file)
    $xl.Run("'" + $wb.Name + "'!更新資訊_靜默")
    $wb.Save()

    Write-Host ''
    Write-Host '產製結果：' -ForegroundColor Cyan
    $n = 0
    foreach ($ws in $wb.Worksheets) {
        if ($ws.Name.StartsWith('基本補助費')) {
            $n++
            $total = $ws.Range('D7').Value2
            Write-Host ('  ‧ {0}' -f $ws.Name) -ForegroundColor Green
            Write-Host ('      合計 {0:N0} 元' -f $total) -ForegroundColor White
        }
    }
    if ($n -eq 0) {
        Write-Host '  （沒有產出任何概算表，請檢查「競賽資訊」的 A/B/C/E 欄是否填妥）' -ForegroundColor Yellow
    }
    Write-Host ''
    Write-Host "完成！共 $n 張概算表，已存回原檔。" -ForegroundColor Green
}
catch {
    Write-Host ''
    Write-Host "發生錯誤：$($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($wb) { try { $wb.Close($false) | Out-Null } catch {} }
    if ($xl) { try { $xl.Quit() | Out-Null } catch {} }
}

Write-Host ''
Read-Host '按 Enter 關閉視窗'

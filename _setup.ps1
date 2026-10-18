#requires -Version 3.0
<#
  體育組差旅 － 本機環境安裝／檢查程式

  用法：
    雙擊「安裝.bat」即可（互動式）
    或：powershell -ExecutionPolicy Bypass -File _setup.ps1 [-CheckOnly] [-NonInteractive]

  本程式會自動做的事（安全）：
    ‧檢查必要檔案是否齊全
    ‧解除 Google Drive／瀏覽器加上的「來自網際網路」封鎖（巨集被擋的主因）
    ‧檢查 Excel 是否安裝、回報目前的巨集安全性設定
    ‧讀取活頁簿內容，確認工作表與 VBA 專案存在

  下列動作「一定會先問過你」，預設都是不做：
    ‧把本資料夾加入 Excel 信任位置（會寫入 HKCU 登錄）
    ‧開啟「信任存取 VBA 專案物件模型」（會寫入 HKCU 登錄）
    ‧建立桌面捷徑
    ‧從 .bas 重新匯入 Module1／Module2
  這幾項都會先印出要變更的內容；不想讓程式改，可照畫面上的手動步驟自己設定。
#>

param(
    [switch]$CheckOnly,
    [switch]$NonInteractive
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$script:warnCount = 0
$script:failCount = 0

#----------------------------------------------------------------------------- 輸出小工具
function Title($t) {
    Write-Host ''
    Write-Host ('=== ' + $t + ' ===') -ForegroundColor Cyan
}
function Ok($m)   { Write-Host ('  [OK]   ' + $m) -ForegroundColor Green }
function Info($m) { Write-Host ('  [--]   ' + $m) -ForegroundColor Gray }
function Warn($m) { Write-Host ('  [注意] ' + $m) -ForegroundColor Yellow; $script:warnCount++ }
function Fail($m) { Write-Host ('  [缺少] ' + $m) -ForegroundColor Red;    $script:failCount++ }
function Step($m) { Write-Host ('  [做了] ' + $m) -ForegroundColor White }

function Ask($question, $defaultYes) {
    if ($NonInteractive -or $CheckOnly) { return $false }
    $hint = '[y/N]'
    if ($defaultYes) { $hint = '[Y/n]' }
    Write-Host ''
    $a = Read-Host ('  ' + $question + ' ' + $hint)
    if ($a -eq '') { return [bool]$defaultYes }
    return ($a -match '^[yY]')
}

function Get-OfficeVersionKey {
    foreach ($v in '16.0', '15.0', '14.0') {
        if (Test-Path ('HKCU:\Software\Microsoft\Office\' + $v + '\Excel')) { return $v }
    }
    return $null
}

# 把 .bas（UTF-8）轉成 VBE 看得懂的 Big5(cp950) 暫存檔後匯入，避免中文變亂碼
function Import-Modules {
    $wb2Path = Join-Path $root '基本補助費概算表格式.xlsm'
    $pairs = @(
        @{ Name = 'Module1'; File = 'Module1_更新資訊.bas' },
        @{ Name = 'Module2'; File = 'Module2_車資試算.bas' }
    )

    foreach ($p in $pairs) {
        if (-not (Test-Path -LiteralPath (Join-Path $root $p.File))) {
            Warn ($p.File + ' 不存在，略過 ' + $p.Name)
            return
        }
    }

    $stamp  = Get-Date -Format 'yyyyMMdd_HHmm'
    $backup = Join-Path $root ('基本補助費概算表格式_匯入前備份_' + $stamp + '.xlsm')
    Copy-Item -LiteralPath $wb2Path -Destination $backup -Force
    Step ('已備份：' + (Split-Path $backup -Leaf))

    $tmpDir = Join-Path $env:TEMP ('tiyu_bas_' + $stamp)
    New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null

    $xl = $null
    $wb = $null
    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $wb = $xl.Workbooks.Open($wb2Path)

        $vbp = $null
        try { $vbp = $wb.VBProject } catch { $vbp = $null }
        if ($null -eq $vbp) {
            Warn '仍無法存取 VBA 專案，請確認 (b) 已開啟並重開 Excel 後再試'
            return
        }

        foreach ($p in $pairs) {
            $srcTxt = Get-Content -LiteralPath (Join-Path $root $p.File) -Raw -Encoding UTF8
            $tmp = Join-Path $tmpDir ($p.Name + '.bas')
            [System.IO.File]::WriteAllText($tmp, $srcTxt, [System.Text.Encoding]::GetEncoding(950))

            foreach ($c in @($vbp.VBComponents)) {
                if ($c.Name -eq $p.Name) { $vbp.VBComponents.Remove($c) }
            }
            $vbp.VBComponents.Import($tmp) | Out-Null

            $cm = $vbp.VBComponents.Item($p.Name).CodeModule
            $txt = $cm.Lines(1, $cm.CountOfLines)
            if ($txt -like '*費率對照*') {
                Step ($p.Name + ' 已匯入，' + $cm.CountOfLines + ' 行，中文正常')
            } else {
                Warn ($p.Name + ' 匯入後中文疑似亂碼，請改用 Alt+F11 手動匯入 Module2_車資試算_Big5.bas')
            }
        }

        # 車資試算不存在時順便建起來
        $hasCalc = $false
        foreach ($s in $wb.Worksheets) { if ($s.Name -eq '車資試算') { $hasCalc = $true } }
        if (-not $hasCalc) {
            $xl.Run(("'" + $wb.Name + "'!建立車資試算表_靜默")) | Out-Null
            Step '已建立「車資試算」工作表'
        }

        $wb.Save()
        Step '活頁簿已存檔'
    }
    catch {
        Warn ('匯入失敗：' + $_.Exception.Message)
        Info ('原檔備份在：' + (Split-Path $backup -Leaf))
    }
    finally {
        if ($wb) { try { $wb.Close($false) | Out-Null } catch { } }
        if ($xl) { try { $xl.Quit() | Out-Null } catch { } }
        if (Test-Path $tmpDir) { Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

Write-Host ''
Write-Host '############################################################' -ForegroundColor Cyan
Write-Host '#  體育組差旅 － 本機環境安裝／檢查                         #' -ForegroundColor Cyan
Write-Host '############################################################' -ForegroundColor Cyan
Info ('專案資料夾：' + $root)
if ($CheckOnly) { Warn '－CheckOnly 模式：只檢查，不做任何變更' }

#----------------------------------------------------------------------------- 1. 檔案
Title '1. 必要檔案'

$wbName   = '基本補助費概算表格式.xlsm'
$wbPath   = Join-Path $root $wbName
$required = @(
    $wbName,
    'Module1_更新資訊.bas',
    'Module2_車資試算.bas',
    '_calc.ps1',
    '一鍵計算差旅費.bat'
)
$optional = [ordered]@{
    '交通費明細.pdf'             = '【參考B】【參考D】的票價來源依據'
    '基本補助費基準表.pdf'       = '【一】地區費率的法源依據'
    '運動賽事地點.docx'          = '【參考D】場館清單的來源'
    'Module2_車資試算_Big5.bas'  = '手動匯入 Module2 用的 Big5 版本'
    'CLAUDE.md'                  = '專案說明（給 Claude Code 與接手的人看）'
}

foreach ($f in $required) {
    if (Test-Path -LiteralPath (Join-Path $root $f)) { Ok $f } else { Fail ($f + '　←　必要檔案，請從 Google Drive 重新同步') }
}
foreach ($f in $optional.Keys) {
    if (Test-Path -LiteralPath (Join-Path $root $f)) { Info ($f + '　－　' + $optional[$f]) }
    else { Warn ($f + ' 不存在（非必要）－ 原用途：' + $optional[$f]) }
}

#----------------------------------------------------------------------------- 2. 解除封鎖
Title '2. 解除「來自網際網路」封鎖'
Info 'Google Drive 同步下來的檔案常被 Windows 標記，標記後 Excel 會直接封鎖巨集。'

$blocked = @()
foreach ($item in (Get-ChildItem -LiteralPath $root -File)) {
    $z = $null
    try { $z = Get-Item -LiteralPath $item.FullName -Stream 'Zone.Identifier' -ErrorAction Stop } catch { $z = $null }
    if ($null -ne $z) { $blocked += $item }
}

if ($blocked.Count -eq 0) {
    Ok '沒有被封鎖的檔案'
} else {
    foreach ($b in $blocked) { Warn ('被封鎖：' + $b.Name) }
    if ($CheckOnly) {
        Info '（CheckOnly 模式，未解除）'
    } else {
        foreach ($b in $blocked) {
            Unblock-File -LiteralPath $b.FullName
            Step ('已解除封鎖：' + $b.Name)
        }
    }
}

#----------------------------------------------------------------------------- 3. Excel
Title '3. Excel 與巨集安全性'

$excelExe = $null
$appPaths = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe'
if (Test-Path $appPaths) {
    try { $excelExe = (Get-ItemProperty -Path $appPaths).'(default)' } catch { $excelExe = $null }
}
if ($excelExe) { Ok ('Excel：' + $excelExe) }
else { Fail ' 找不到 Excel。本專案的計算完全依賴 Excel 巨集，請先安裝 Microsoft Excel（桌面版，非網頁版）。' }

$ov = Get-OfficeVersionKey
if ($null -eq $ov) {
    Warn '找不到 Office 使用者設定（HKCU\Software\Microsoft\Office\<版本>\Excel）'
} else {
    Ok ('Office 設定版本：' + $ov)
    $secKey = 'HKCU:\Software\Microsoft\Office\' + $ov + '\Excel\Security'

    $vbaWarn = $null
    if (Test-Path $secKey) { $vbaWarn = (Get-ItemProperty -Path $secKey).VBAWarnings }
    switch ($vbaWarn) {
        1       { Warn '巨集設定＝啟用所有巨集（安全性偏低，但本專案可直接執行）' }
        2       { Info '巨集設定＝停用並通知（Windows 預設）。開檔後要按上方黃色列的「啟用內容」。' }
        3       { Info '巨集設定＝僅啟用數位簽署的巨集。本專案未簽章，需用「信任位置」才能跑。' }
        4       { Warn '巨集設定＝停用所有巨集且不通知。必須設定「信任位置」，否則巨集完全不會執行。' }
        default { Info '巨集設定＝未設定，採 Office 預設（停用並通知）' }
    }

    $vbom = $null
    if (Test-Path $secKey) { $vbom = (Get-ItemProperty -Path $secKey).AccessVBOM }
    if ($vbom -eq 1) { Ok '信任存取 VBA 專案物件模型＝已開啟（可用本程式自動匯入 .bas）' }
    else { Info '信任存取 VBA 專案物件模型＝未開啟（只影響「自動匯入模組」，日常計算不需要）' }

    # 信任位置
    $tlRoot = $secKey + '\Trusted Locations'
    $alreadyTrusted = $false
    if (Test-Path $tlRoot) {
        foreach ($loc in (Get-ChildItem -Path $tlRoot -ErrorAction SilentlyContinue)) {
            $p = (Get-ItemProperty -Path $loc.PSPath).Path
            if ($p) {
                $p = $p.TrimEnd('\')
                if ($p -eq $root.TrimEnd('\')) { $alreadyTrusted = $true }
            }
        }
    }
    if ($alreadyTrusted) { Ok '本資料夾已在 Excel 信任位置清單中' }
    else { Info '本資料夾尚未加入 Excel 信任位置' }
}

#----------------------------------------------------------------------------- 4. 活頁簿內容
Title '4. 活頁簿內容（直接讀檔，不需開啟 Excel）'

if (-not (Test-Path -LiteralPath $wbPath)) {
    Fail ($wbName + ' 不存在，略過內容檢查')
} else {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($wbPath)

        $e = $zip.GetEntry('xl/workbook.xml')
        $sr = New-Object System.IO.StreamReader($e.Open())
        $xml = [xml]$sr.ReadToEnd()
        $sr.Close()

        $sheets = @()
        foreach ($s in $xml.workbook.sheets.sheet) { $sheets += $s.name }
        Info ('工作表：' + ($sheets -join '、'))

        foreach ($need in '競賽資訊', '費率對照', '車資試算') {
            if ($sheets -contains $need) { Ok ('工作表「' + $need + '」存在') }
            elseif ($need -eq '車資試算') { Warn '工作表「車資試算」不存在 － 匯入模組後執行巨集【建立車資試算表】即可產生' }
            else { Fail ('工作表「' + $need + '」不存在，活頁簿可能不完整') }
        }

        $vba = $zip.GetEntry('xl/vbaProject.bin')
        if ($null -eq $vba) { Fail 'VBA 專案不存在（vbaProject.bin）－ 巨集完全不會動，請重新同步活頁簿' }
        else { Ok ('VBA 專案存在，' + [math]::Round($vba.Length / 1KB) + ' KB') }
    }
    catch {
        Fail ('讀取活頁簿失敗：' + $_.Exception.Message)
    }
    finally {
        if ($zip) { $zip.Dispose() }
    }
}

#----------------------------------------------------------------------------- 5. 選用設定
Title '5. 選用設定（都會先問過你）'

if ($CheckOnly -or $NonInteractive) {
    Info '非互動模式，略過選用設定。'
} else {

    # 5a. 信任位置
    if ($null -ne $ov -and -not $alreadyTrusted) {
        Write-Host ''
        Write-Host '  (a) 把本資料夾加入 Excel 信任位置' -ForegroundColor White
        Write-Host '      好處：開檔不再出現「啟用內容」黃色列，巨集直接可用。' -ForegroundColor Gray
        Write-Host '      代價：這是安全性設定。此資料夾內的 Office 檔案之後都會被直接信任，' -ForegroundColor Gray
        Write-Host '            所以只有在你確定這個 Google Drive 資料夾只有自己人放東西時才適合。' -ForegroundColor Gray
        Write-Host ('      會寫入：' + $tlRoot + '\LocationN') -ForegroundColor DarkGray
        Write-Host ('              Path = ' + $root) -ForegroundColor DarkGray
        Write-Host '              AllowSubfolders = 1' -ForegroundColor DarkGray
        Write-Host '      手動做法：Excel →檔案→選項→信任中心→信任中心設定→信任位置→新增位置' -ForegroundColor DarkGray
        if (Ask '要現在加入信任位置嗎？' $false) {
            if (-not (Test-Path $tlRoot)) { New-Item -Path $tlRoot -Force | Out-Null }
            $n = 1
            while (Test-Path ($tlRoot + '\Location' + $n)) { $n++ }
            $new = New-Item -Path ($tlRoot + '\Location' + $n) -Force
            New-ItemProperty -Path $new.PSPath -Name 'Path' -Value $root -PropertyType String -Force | Out-Null
            New-ItemProperty -Path $new.PSPath -Name 'AllowSubfolders' -Value 1 -PropertyType DWord -Force | Out-Null
            New-ItemProperty -Path $new.PSPath -Name 'Description' -Value '體育組差旅' -PropertyType String -Force | Out-Null
            Step ('已加入信任位置 Location' + $n)
            if ($root -notmatch '^[A-Za-z]:\\') {
                Warn '本資料夾看起來是網路路徑；若仍被擋，需另外勾選「允許信任我的網路上的位置」。'
            }
        } else {
            Info '略過（可照上面的手動做法自己設定）'
        }
    }

    # 5b. VBOM
    if ($null -ne $ov -and $vbom -ne 1) {
        Write-Host ''
        Write-Host '  (b) 開啟「信任存取 VBA 專案物件模型」' -ForegroundColor White
        Write-Host '      好處：本程式與 Claude Code 才能自動把 .bas 匯入活頁簿。' -ForegroundColor Gray
        Write-Host '      代價：這是安全性設定。開啟後，任何能跑巨集的程式都能改寫 VBA 程式碼。' -ForegroundColor Gray
        Write-Host '            日常計算不需要它 － 只有要更新模組時才用得到。' -ForegroundColor Gray
        Write-Host ('      會寫入：' + $secKey + '\AccessVBOM = 1 (DWORD)') -ForegroundColor DarkGray
        Write-Host '      手動做法：Excel →檔案→選項→信任中心→信任中心設定→巨集設定→' -ForegroundColor DarkGray
        Write-Host '                勾「信任存取 VBA 專案物件模型」' -ForegroundColor DarkGray
        if (Ask '要現在開啟嗎？' $false) {
            if (-not (Test-Path $secKey)) { New-Item -Path $secKey -Force | Out-Null }
            New-ItemProperty -Path $secKey -Name 'AccessVBOM' -Value 1 -PropertyType DWord -Force | Out-Null
            Step '已開啟 AccessVBOM'
            $vbom = 1
        } else {
            Info '略過（之後要匯入模組時，用 Alt+F11 →檔案→匯入檔案 手動做也可以）'
        }
    }

    # 5c. 桌面捷徑
    Write-Host ''
    Write-Host '  (c) 在桌面建立「一鍵計算差旅費」捷徑' -ForegroundColor White
    if (Ask '要建立桌面捷徑嗎？' $false) {
        $bat = Join-Path $root '一鍵計算差旅費.bat'
        if (Test-Path -LiteralPath $bat) {
            $lnkPath = Join-Path ([Environment]::GetFolderPath('Desktop')) '體育組差旅－一鍵計算.lnk'
            $wsh = New-Object -ComObject WScript.Shell
            $lnk = $wsh.CreateShortcut($lnkPath)
            $lnk.TargetPath = $bat
            $lnk.WorkingDirectory = $root
            $lnk.Description = '體育組差旅：重算所有基本補助費概算表'
            $lnk.Save()
            Step ('已建立捷徑：' + $lnkPath)
        } else {
            Warn '找不到「一鍵計算差旅費.bat」，略過'
        }
    } else {
        Info '略過'
    }

    # 5d. 重新匯入模組
    Write-Host ''
    Write-Host '  (d) 從 .bas 重新匯入 Module1／Module2 到活頁簿' -ForegroundColor White
    Write-Host '      什麼時候要做：.bas 比活頁簿內的程式碼新，或活頁簿的巨集壞掉／不見了。' -ForegroundColor Gray
    Write-Host '      會先自動備份活頁簿再動作。' -ForegroundColor Gray
    if ($vbom -ne 1) {
        Info '需先完成 (b) 才能自動匯入；目前略過。'
    } elseif (-not $excelExe) {
        Info '需安裝 Excel 才能匯入；目前略過。'
    } elseif (Ask '要現在重新匯入模組嗎？' $false) {
        Import-Modules
    } else {
        Info '略過'
    }
}

#----------------------------------------------------------------------------- 總結
Title '結果'
if ($script:failCount -gt 0) {
    Write-Host ('  有 ' + $script:failCount + ' 項缺少、' + $script:warnCount + ' 項注意事項，請看上面紅字。') -ForegroundColor Red
} elseif ($script:warnCount -gt 0) {
    Write-Host ('  必要項目都齊全，有 ' + $script:warnCount + ' 項注意事項（黃字）。') -ForegroundColor Yellow
} else {
    Write-Host '  全部正常，可以開始使用。' -ForegroundColor Green
}

Write-Host ''
Write-Host '  日常使用：' -ForegroundColor Cyan
Write-Host '    1. 開啟「基本補助費概算表格式.xlsm」（若出現黃色列請按「啟用內容」）'
Write-Host '    2. 在「競賽資訊」填比賽資料 → 按【計算差旅費】'
Write-Host '    3. 要算火車票價 → 到「車資試算」選場館 → 按【寫入費率對照【三】】'
Write-Host '    ‧Excel 沒開著時，也可直接雙擊「一鍵計算差旅費.bat」'
Write-Host ''

if (-not $NonInteractive) { Read-Host '  按 Enter 結束' | Out-Null }

<#
    MailSorter 環境檢查：確認這台電腦（GCB 設定下）能否執行 MailSorter。
    不會修改任何設定，只做讀取與測試。
#>
$ok = $true
function Show-Result([string]$Item, [bool]$Pass, [string]$Detail) {
    if ($Pass) { Write-Host ('[通過] ' + $Item) -ForegroundColor Green }
    else       { Write-Host ('[失敗] ' + $Item) -ForegroundColor Red; $script:ok = $false }
    if ($Detail) { Write-Host ('       ' + $Detail) }
}

Write-Host '==== MailSorter 環境檢查 ====' -ForegroundColor Cyan
Write-Host ''

$mode = $ExecutionContext.SessionState.LanguageMode
Show-Result 'PowerShell 語言模式' ($mode -eq 'FullLanguage') "目前：$mode（需要 FullLanguage）"

Write-Host ''
Write-Host '執行原則（Get-ExecutionPolicy -List）：'
Get-ExecutionPolicy -List | Format-Table -AutoSize | Out-String | Write-Host
$mp = Get-ExecutionPolicy -Scope MachinePolicy
$up = Get-ExecutionPolicy -Scope UserPolicy
$gpo = @($mp, $up) | Where-Object { $_ -ne 'Undefined' } | Select-Object -First 1
if ($gpo) {
    Show-Result '群組原則設定的執行原則' ($gpo -in @('RemoteSigned', 'Unrestricted', 'Bypass')) "群組原則：$gpo（AllSigned/Restricted 時需請資訊單位簽章或核准）"
} else {
    Show-Result '群組原則設定的執行原則' $true '未由群組原則強制，可由 MailSorter.bat 以 RemoteSigned 執行'
}

try {
    $ol = New-Object -ComObject Outlook.Application
    $name = $ol.GetNamespace('MAPI').CurrentUser.Name
    Show-Result '連線 Outlook' $true "目前使用者：$name"
} catch { Show-Result '連線 Outlook' $false $_.Exception.Message }

# Outlook 程式設計存取安全性：實際測試讀取寄件者地址與另存信件（另存到暫存檔後立即刪除）
if ($ol) {
    $mail = $null
    try {
        $inbox = $ol.GetNamespace('MAPI').GetDefaultFolder(6)
        $items = $inbox.Items
        $items.Sort('[ReceivedTime]', $true)
        $n = 0
        foreach ($it in $items) { $n++; if ($it.Class -eq 43) { $mail = $it; break }; if ($n -ge 30) { break } }
    } catch { }
    if (-not $mail) {
        Write-Host '[略過] Outlook 程式存取測試：收件匣中找不到可測試的信件'
    } else {
        try {
            $addr = [string]$mail.SenderEmailAddress
            Show-Result 'Outlook 程式存取：讀取寄件者地址' $true ''
        } catch { Show-Result 'Outlook 程式存取：讀取寄件者地址' $false ('被拒絕：' + $_.Exception.Message) }
        $tmp = Join-Path $env:TEMP ('MailSorter_check_{0}.msg' -f [guid]::NewGuid().ToString('N'))
        try {
            $mail.SaveAs($tmp, 9)
            Show-Result 'Outlook 程式存取：另存信件（轉 PDF 必需）' $true ''
        } catch {
            Show-Result 'Outlook 程式存取：另存信件（轉 PDF 必需）' $false ('被拒絕：' + $_.Exception.Message)
            Write-Host '       Outlook「程式設計存取安全性」設定為自動拒絕，MailSorter 無法存檔。' -ForegroundColor Yellow
            Write-Host '       這是機關的安全性管控，請將本畫面截圖提供資訊單位評估。' -ForegroundColor Yellow
        } finally { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
    }

    # 列出相關的 Outlook 安全性原則（原始值，供資訊單位判讀）
    $names = 'AdminSecurityMode', 'ObjectModelGuard', 'PromptOOMSaveAs', 'PromptOOMAddressInformationAccess', 'PromptOOMAddressBookAccess', 'PromptOOMFormulaAccess'
    $found = @()
    foreach ($hive in 'HKCU', 'HKLM') {
        foreach ($ver in '16.0', '15.0', '14.0') {
            $key = "${hive}:\Software\Policies\Microsoft\Office\$ver\Outlook\Security"
            if (Test-Path -LiteralPath $key) {
                $props = Get-ItemProperty -LiteralPath $key -ErrorAction SilentlyContinue
                foreach ($nm in $names) {
                    $v = $props.$nm
                    if ($null -ne $v) { $found += ('       {0}\...\{1}\Outlook\Security  {2} = {3}' -f $hive, $ver, $nm, $v) }
                }
            }
        }
    }
    Write-Host ''
    Write-Host 'Outlook 安全性原則（原始值，供資訊單位判讀）：'
    if ($found.Count) { $found | ForEach-Object { Write-Host $_ } } else { Write-Host '       （未找到相關原則設定）' }
    Write-Host ''
}

try {
    $w = New-Object -ComObject Word.Application
    $ver = $w.Version
    $w.Quit()
    Show-Result '啟動 Word（轉 PDF 用）' $true "Word 版本：$ver"
} catch { Show-Result '啟動 Word（轉 PDF 用）' $false $_.Exception.Message }

$files = Get-ChildItem -LiteralPath $PSScriptRoot -File -ErrorAction SilentlyContinue |
    Where-Object { Get-Item -LiteralPath $_.FullName -Stream Zone.Identifier -ErrorAction SilentlyContinue }
if ($files) {
    Show-Result '檔案未被標記為「來自網際網路」' $false ('請在檔案「內容」勾選「解除封鎖」：' + (($files | ForEach-Object { $_.Name }) -join '、'))
} else {
    Show-Result '檔案未被標記為「來自網際網路」' $true ''
}

Write-Host ''
if ($ok) { Write-Host '全部通過，可以執行 MailSorter.bat。' -ForegroundColor Green }
else     { Write-Host '有項目未通過，請參考 README.md「無法執行時」一節，或洽資訊單位。' -ForegroundColor Yellow }
Write-Host ''
Read-Host '按 Enter 結束'

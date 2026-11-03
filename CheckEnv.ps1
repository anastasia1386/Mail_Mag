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

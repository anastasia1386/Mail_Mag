#Requires -Version 5.1
<#
    MailSorter - Outlook 信件自動分類歸檔工具（免安裝）

    功能：
      * 以下拉式選單選擇 Outlook 資料夾作為資料來源（可多個、可含子資料夾）
      * 依寄件者／主旨／本文規則，將信件存到對應的本機目錄（可自動建立目錄）
      * 信件本文轉存 PDF（透過 Word），並可保存原始 .msg 與附件
      * 規則每次儲存都會留下修改紀錄（誰、何時、改了什麼），並可還原至任一版本
      * 預覽模式：只列出會存到哪裡，不實際存檔

    只使用 Windows 內建的 PowerShell 5.1、Outlook 與 Word，不需安裝任何軟體。
    執行方式：雙擊 MailSorter.bat
              ArchiveSelected.bat：不開管理介面，直接歸檔 Outlook 中目前選取的信件
#>
param(
    [switch]$ProcessSelected,   # 直接處理 Outlook 中選取的信件（不開管理介面）
    [switch]$Preview,           # 搭配 -ProcessSelected：只預覽、不存檔
    [switch]$Force              # 搭配 -ProcessSelected：已歸檔（有類別）的信件也重新處理
)

# ---------------------------------------------------------------------------
# 環境檢查（受限語言模式下無法呼叫 Outlook / Word，也無法建立視窗）
# ---------------------------------------------------------------------------
if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
    Write-Host '此電腦的 PowerShell 為「受限語言模式」(ConstrainedLanguage)，本程式無法執行。' -ForegroundColor Red
    Write-Host '請洽資訊單位協助將本程式加入允許清單（或改以數位簽章方式核准）。'
    Read-Host '按 Enter 結束'
    exit 1
}

# 啟動階段若發生未處理的錯誤：顯示訊息、寫入 logs\startup_error.log，並以代碼 1 結束（啟動檔會暫停讓使用者看到訊息）
trap {
    $err = $_
    $text = "MailSorter 發生錯誤，無法繼續執行：`r`n`r`n" + $err.Exception.Message + "`r`n`r`n" + $err.InvocationInfo.PositionMessage
    Write-Host ''
    Write-Host $text -ForegroundColor Red
    try {
        $logDirTrap = Join-Path $PSScriptRoot 'logs'
        if (-not (Test-Path -LiteralPath $logDirTrap)) { [void][System.IO.Directory]::CreateDirectory($logDirTrap) }
        $line = "[{0}] PowerShell {1}`r`n{2}`r`n{3}`r`n`r`n" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $PSVersionTable.PSVersion, $err.Exception.ToString(), $err.InvocationInfo.PositionMessage
        [System.IO.File]::AppendAllText((Join-Path $logDirTrap 'startup_error.log'), $line, (New-Object System.Text.UTF8Encoding($true)))
    } catch { }
    try { [void][System.Windows.Forms.MessageBox]::Show($text, 'MailSorter', 'OK', 'Error') } catch { }
    exit 1
}

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'
try { $Host.UI.RawUI.WindowTitle = 'MailSorter（請勿關閉此視窗）' } catch { }
Write-Host 'MailSorter 啟動中…（此視窗會在程式開啟後自動縮到工作列，使用期間請勿關閉）'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$AppName    = 'MailSorter 信件歸檔工具'
$AppVersion = '1.1.1'
$BaseDir    = $PSScriptRoot
if (-not $BaseDir) { $BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$ConfigPath = Join-Path $BaseDir 'config.json'
$HistoryDir = Join-Path $BaseDir 'history'
$LogDir     = Join-Path $BaseDir 'logs'
$TempDir    = Join-Path $env:TEMP 'MailSorter'
$Utf8NoBom  = New-Object System.Text.UTF8Encoding($false)

# Outlook / Word 常數
$olFolderMail   = 0     # Folder.DefaultItemType：郵件資料夾
$olMailClass    = 43    # MailItem.Class
$olMSGUnicode   = 9
$olMHTML        = 10
$wdFormatPDF    = 17
$PR_ATTACH_HIDDEN     = 'http://schemas.microsoft.com/mapi/proptag/0x7FFE000B'
$PR_ATTACH_CONTENT_ID = 'http://schemas.microsoft.com/mapi/proptag/0x3712001F'

# ---------------------------------------------------------------------------
# 預設值
# ---------------------------------------------------------------------------
$RuleDefaults = [ordered]@{
    Id              = ''
    Name            = '新規則'
    Enabled         = $true
    SourceId        = ''         # 空白 = 所有來源
    MatchMode       = 'All'      # All = 全部條件成立；Any = 任一條件成立
    SenderContains  = ''
    SubjectContains = ''
    BodyContains    = ''
    UseRegex        = $false
    TargetPath      = ''
    CreateFolder    = $true
    SavePdf         = $true
    SaveMsg         = $true
    SaveAttachments = $false
    StopProcessing  = $true
}

$SourceDefaults = [ordered]@{
    Id                = ''
    FolderPath        = ''
    EntryID           = ''
    StoreID           = ''
    IncludeSubfolders = $false
    Enabled           = $true
}

$OptionDefaults = [ordered]@{
    FileNameTemplate  = '{yyyy}{MM}{dd}_{HH}{mm}_{主旨}'
    DaysBack          = 7
    UnreadOnly        = $false
    SkipCategorized   = $true
    CategoryName      = '已歸檔'
    AddCategory       = $true
    MarkAsRead        = $false
    SkipInlineImages  = $true
    DefaultTargetPath = ''
}

# 規則欄位的中文名稱（修改紀錄使用，順序即顯示順序）
$RuleFieldLabels = [ordered]@{
    Name            = '規則名稱'
    Enabled         = '啟用'
    SourceId        = '限定來源'
    MatchMode       = '條件組合'
    SenderContains  = '寄件者包含'
    SubjectContains = '主旨包含'
    BodyContains    = '本文包含'
    UseRegex        = '使用規則運算式'
    TargetPath      = '存放路徑'
    CreateFolder    = '自動建立目錄'
    SavePdf         = '輸出 PDF'
    SaveMsg         = '保存原始信件'
    SaveAttachments = '保存附件'
    StopProcessing  = '符合後停止比對'
}

$PlaceholderHelp = '可用代碼：{寄件者} {寄件者信箱} {寄件網域} {主旨} {規則} {yyyy} {MM} {dd} {HH} {mm}'

# ---------------------------------------------------------------------------
# 共用函式
# ---------------------------------------------------------------------------
function Show-Msg([string]$Text, [string]$Icon = 'Information') {
    if ($script:headless) {
        # 沒有主視窗時，讓訊息顯示在最上層（避免被 Outlook 視窗蓋住）
        [void][System.Windows.Forms.MessageBox]::Show($Text, $AppName, 'OK', $Icon, 'Button1', 'DefaultDesktopOnly')
    } else {
        [void][System.Windows.Forms.MessageBox]::Show($Text, $AppName, 'OK', $Icon)
    }
}

function Confirm-Msg([string]$Text) {
    return ([System.Windows.Forms.MessageBox]::Show($Text, $AppName, 'YesNo', 'Question') -eq 'Yes')
}

function Read-JsonFile([string]$Path) {
    $text = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
    return (ConvertFrom-Json $text)
}

function Write-JsonFile([string]$Path, $Object) {
    $json = ConvertTo-Json -InputObject $Object -Depth 10
    [System.IO.File]::WriteAllText($Path, $json, $Utf8NoBom)
}

# 將任何集合轉成 ArrayList（避免 PowerShell 5.1 陣列展開問題）
function ConvertTo-List($Items) {
    $list = New-Object System.Collections.ArrayList
    if ($null -ne $Items) {
        foreach ($i in $Items) { if ($null -ne $i) { [void]$list.Add($i) } }
    }
    return , $list
}

# 補上缺少的欄位（讀取舊版設定檔時使用）
function Complete-Object($Object, $Defaults) {
    $names = @($Object.PSObject.Properties | ForEach-Object { $_.Name })
    foreach ($k in $Defaults.Keys) {
        if ($names -notcontains $k) {
            $Object | Add-Member -NotePropertyName $k -NotePropertyValue $Defaults[$k]
        }
    }
    return $Object
}

# 深層複製規則清單
function Copy-Rules($Rules) {
    $arr = @()
    foreach ($r in $Rules) { $arr += $r }
    $json = ConvertTo-Json -InputObject $arr -Depth 10
    $list = ConvertTo-List (ConvertFrom-Json $json)
    foreach ($r in $list) {
        [void](Complete-Object $r $RuleDefaults)
        if (-not $r.Id) { $r.Id = [guid]::NewGuid().ToString() }
    }
    return , $list
}

function New-Rule {
    $r = [pscustomobject]$RuleDefaults
    $r.Id = [guid]::NewGuid().ToString()
    $r.TargetPath = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) '信件歸檔\{寄件者}\{yyyy}{MM}')
    return $r
}

function Split-Keywords([string]$Text) {
    if (-not $Text) { return @() }
    return @($Text -split '[;；]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Test-Keywords([string]$Text, [string]$Keywords, [bool]$UseRegex) {
    if ($null -eq $Text) { $Text = '' }
    foreach ($k in (Split-Keywords $Keywords)) {
        if ($UseRegex) {
            if ([regex]::IsMatch($Text, $k, 'IgnoreCase')) { return $true }
        } else {
            if ($Text.IndexOf($k, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
        }
    }
    return $false
}

function Get-SafeName([string]$Text, [int]$MaxLength = 80) {
    if ($null -eq $Text) { $Text = '' }
    $s = $Text -replace '[\\/:*?"<>|\r\n\t]', '_'
    $s = $s.Trim().TrimEnd('.').Trim()
    if ($s.Length -gt $MaxLength) { $s = $s.Substring(0, $MaxLength).Trim() }
    if (-not $s) { $s = '_' }
    return $s
}

# 展開路徑/檔名範本
function Expand-Template([string]$Template, $Info, [string]$RuleName, [switch]$IsPath) {
    $t = [datetime]$Info.Received
    $map = [ordered]@{
        '{寄件者信箱}' = $Info.SenderEmail
        '{寄件網域}'   = $Info.Domain
        '{寄件者}'     = $Info.SenderName
        '{主旨}'       = $Info.Subject
        '{規則}'       = $RuleName
        '{yyyy}'       = $t.ToString('yyyy')
        '{MM}'         = $t.ToString('MM')
        '{dd}'         = $t.ToString('dd')
        '{HH}'         = $t.ToString('HH')
        '{mm}'         = $t.ToString('mm')
    }
    $result = [string]$Template
    foreach ($k in $map.Keys) {
        # String.Replace 區分大小寫，{MM}(月) 與 {mm}(分) 不會混淆
        $result = $result.Replace($k, (Get-SafeName ([string]$map[$k])))
    }
    if ($IsPath) {
        $parts = $result -split '\\'
        for ($i = 1; $i -lt $parts.Count; $i++) {
            if ($parts[$i]) { $parts[$i] = $parts[$i].TrimEnd('. ') ; if (-not $parts[$i]) { $parts[$i] = '_' } }
        }
        $result = $parts -join '\'
        if (-not [System.IO.Path]::IsPathRooted($result)) { $result = Join-Path $BaseDir $result }
    } else {
        $result = Get-SafeName $result 120
    }
    return $result
}

function Get-UniqueBase([string]$Dir, [string]$Name) {
    $maxLen = 230 - $Dir.Length
    if ($maxLen -lt 20) { $maxLen = 20 }
    if ($Name.Length -gt $maxLen) { $Name = $Name.Substring(0, $maxLen).Trim() }
    $base = Join-Path $Dir $Name
    $n = 1
    while ([System.IO.File]::Exists("$base.pdf") -or [System.IO.File]::Exists("$base.msg") -or [System.IO.Directory]::Exists("$base`_附件")) {
        $n++
        $base = Join-Path $Dir ("{0}_{1}" -f $Name, $n)
    }
    return $base
}

function Get-UniqueFile([string]$Dir, [string]$FileName) {
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $ext  = [System.IO.Path]::GetExtension($FileName)
    $p = Join-Path $Dir $FileName
    $n = 1
    while ([System.IO.File]::Exists($p)) {
        $n++
        $p = Join-Path $Dir ("{0}_{1}{2}" -f $stem, $n, $ext)
    }
    return $p
}

# ---------------------------------------------------------------------------
# 設定檔
# ---------------------------------------------------------------------------
function Import-Config {
    $c = $null
    if (Test-Path -LiteralPath $ConfigPath) {
        try { $c = Read-JsonFile $ConfigPath }
        catch {
            Show-Msg ("設定檔 config.json 讀取失敗，將使用預設設定。`n`n" + $_.Exception.Message) 'Warning'
            $c = $null
        }
    }
    $sources = ConvertTo-List $(if ($c) { $c.Sources })
    foreach ($s in $sources) { [void](Complete-Object $s $SourceDefaults) }
    $rules = Copy-Rules $(if ($c) { $c.Rules })
    if ($c -and $c.Options) { $opt = Complete-Object $c.Options $OptionDefaults }
    else { $opt = [pscustomobject]$OptionDefaults }

    $script:config = [pscustomobject]@{
        Version = 1
        Sources = $sources
        Options = $opt
    }
    $script:savedRules = $rules
    $script:isNewConfig = ($null -eq $c)
}

function Export-Config {
    $obj = [pscustomobject]@{
        Version   = 1
        SavedAt   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Sources   = @($script:config.Sources)
        Rules     = @($script:savedRules)
        Options   = $script:config.Options
    }
    Write-JsonFile $ConfigPath $obj
}

# ---------------------------------------------------------------------------
# 規則修改紀錄
# ---------------------------------------------------------------------------
function Format-FieldValue([string]$Field, $Value) {
    if ($Field -eq 'SourceId') {
        if (-not $Value) { return '(所有來源)' }
        $s = $script:config.Sources | Where-Object { $_.Id -eq $Value } | Select-Object -First 1
        if ($s) { return $s.FolderPath } else { return '(已移除的來源)' }
    }
    if ($Field -eq 'MatchMode') {
        if ($Value -eq 'Any') { return '任一條件成立' } else { return '全部條件成立' }
    }
    if ($Value -is [bool]) { if ($Value) { return '是' } else { return '否' } }
    return [string]$Value
}

function Format-RuleSummary($Rule) {
    $parts = @()
    foreach ($f in $RuleFieldLabels.Keys) {
        if ($f -eq 'Name') { continue }
        $parts += ('{0}={1}' -f $RuleFieldLabels[$f], (Format-FieldValue $f $Rule.$f))
    }
    return ($parts -join '；')
}

function Get-RuleChanges($Old, $New) {
    $changes = New-Object System.Collections.ArrayList
    $Old = @($Old | Where-Object { $null -ne $_ -and $_.Id })
    $New = @($New | Where-Object { $null -ne $_ -and $_.Id })
    $oldMap = @{}; foreach ($r in $Old) { $oldMap[$r.Id] = $r }
    $newMap = @{}; foreach ($r in $New) { $newMap[$r.Id] = $r }

    foreach ($r in $New) {
        if (-not $oldMap.ContainsKey($r.Id)) {
            [void]$changes.Add([pscustomobject]@{ Action = '新增'; RuleId = $r.Id; RuleName = $r.Name; Field = ''; Before = ''; After = (Format-RuleSummary $r) })
            continue
        }
        $o = $oldMap[$r.Id]
        foreach ($f in $RuleFieldLabels.Keys) {
            if ([string]$o.$f -cne [string]$r.$f) {
                [void]$changes.Add([pscustomobject]@{
                    Action = '修改'; RuleId = $r.Id; RuleName = $r.Name; Field = $RuleFieldLabels[$f]
                    Before = (Format-FieldValue $f $o.$f); After = (Format-FieldValue $f $r.$f)
                })
            }
        }
    }
    foreach ($r in $Old) {
        if (-not $newMap.ContainsKey($r.Id)) {
            [void]$changes.Add([pscustomobject]@{ Action = '刪除'; RuleId = $r.Id; RuleName = $r.Name; Field = ''; Before = (Format-RuleSummary $r); After = '' })
        }
    }
    # 順序變更（只比較兩版共有的規則）
    $oldOrder = @($Old | Where-Object { $newMap.ContainsKey($_.Id) } | ForEach-Object { $_.Id })
    $newOrder = @($New | Where-Object { $oldMap.ContainsKey($_.Id) } | ForEach-Object { $_.Id })
    if (($oldOrder -join '|') -ne ($newOrder -join '|')) {
        [void]$changes.Add([pscustomobject]@{
            Action = '順序'; RuleId = ''; RuleName = ''; Field = '規則順序'
            Before = (@($Old | Where-Object { $newMap.ContainsKey($_.Id) } | ForEach-Object { $_.Name }) -join ' > ')
            After  = (@($New | Where-Object { $oldMap.ContainsKey($_.Id) } | ForEach-Object { $_.Name }) -join ' > ')
        })
    }
    return , $changes
}

function Get-ChangeSummary($Changes) {
    $add = @($Changes | Where-Object { $_.Action -eq '新增' }).Count
    $mod = @($Changes | Where-Object { $_.Action -eq '修改' } | ForEach-Object { $_.RuleId } | Select-Object -Unique).Count
    $del = @($Changes | Where-Object { $_.Action -eq '刪除' }).Count
    $ord = @($Changes | Where-Object { $_.Action -eq '順序' }).Count
    $s = "新增 $add 條、修改 $mod 條、刪除 $del 條"
    if ($ord) { $s += '、調整順序' }
    return $s
}

function Add-HistoryEntry([string]$Action, [string]$Note, $Changes, $Rules) {
    if (-not (Test-Path -LiteralPath $HistoryDir)) { [void][System.IO.Directory]::CreateDirectory($HistoryDir) }
    $now = Get-Date
    $entry = [pscustomobject]@{
        Time     = $now.ToString('yyyy-MM-dd HH:mm:ss')
        User     = "$env:USERDOMAIN\$env:USERNAME"
        Computer = $env:COMPUTERNAME
        Action   = $Action
        Note     = $Note
        Summary  = (Get-ChangeSummary $Changes)
        Changes  = @($Changes)
        Rules    = @($Rules)
    }
    $file = Join-Path $HistoryDir ('rules_{0}.json' -f $now.ToString('yyyyMMdd_HHmmss_fff'))
    Write-JsonFile $file $entry
}

function Get-HistoryEntries {
    if (-not (Test-Path -LiteralPath $HistoryDir)) { return }
    Get-ChildItem -LiteralPath $HistoryDir -Filter 'rules_*.json' | Sort-Object Name -Descending | ForEach-Object {
        try {
            $e = Read-JsonFile $_.FullName
            $e | Add-Member -NotePropertyName FilePath -NotePropertyValue $_.FullName -Force
            $e
        } catch { }
    }
}

function Format-HistoryDetail($Entry) {
    $nl = "`r`n"
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("時間：$($Entry.Time)$nl")
    [void]$sb.Append("使用者：$($Entry.User)（電腦：$($Entry.Computer)）$nl")
    [void]$sb.Append("動作：$($Entry.Action)$nl")
    if ($Entry.Note) { [void]$sb.Append("修改說明：$($Entry.Note)$nl") }
    [void]$sb.Append("摘要：$($Entry.Summary)$nl$nl")
    [void]$sb.Append("── 變更內容 ──$nl")
    foreach ($c in $Entry.Changes) {
        switch ($c.Action) {
            '新增' { [void]$sb.Append("[新增] 「$($c.RuleName)」$nl      $($c.After)$nl") }
            '刪除' { [void]$sb.Append("[刪除] 「$($c.RuleName)」$nl      原內容：$($c.Before)$nl") }
            '修改' { [void]$sb.Append("[修改] 「$($c.RuleName)」 $($c.Field)：「$($c.Before)」 → 「$($c.After)」$nl") }
            '順序' { [void]$sb.Append("[順序] 調整規則順序$nl      原：$($c.Before)$nl      新：$($c.After)$nl") }
        }
    }
    $rules = @($Entry.Rules)
    [void]$sb.Append("$nl── 此版本共 $($rules.Count) 條規則 ──$nl")
    $i = 0
    foreach ($r in $rules) {
        $i++
        $flag = ''; if (-not $r.Enabled) { $flag = '（停用）' }
        [void]$sb.Append("$i. $($r.Name)$flag$nl")
    }
    return $sb.ToString()
}

# ---------------------------------------------------------------------------
# Outlook
# ---------------------------------------------------------------------------
$script:outlook = $null
$script:ns = $null
$script:word = $null

function Connect-Outlook {
    if (-not $script:ns) {
        $script:outlook = New-Object -ComObject Outlook.Application
        $script:ns = $script:outlook.GetNamespace('MAPI')
    }
    return $script:ns
}

function Add-FolderRecursive($Folder, $List, [int]$Depth) {
    try {
        if ($Folder.DefaultItemType -eq $olFolderMail) {
            [void]$List.Add([pscustomobject]@{ FolderPath = [string]$Folder.FolderPath; EntryID = [string]$Folder.EntryID; StoreID = [string]$Folder.StoreID })
        }
        if ($Depth -lt 10) {
            foreach ($sub in $Folder.Folders) { Add-FolderRecursive $sub $List ($Depth + 1) }
        }
    } catch { }   # 無權限或離線的資料夾略過
}

function Get-OutlookFolders {
    $ns = Connect-Outlook
    $list = New-Object System.Collections.ArrayList
    foreach ($root in $ns.Folders) { Add-FolderRecursive $root $list 0 }
    return , $list
}

function Format-FolderPath([string]$Path) {
    return (($Path.TrimStart('\') -split '\\') -join '  >  ')
}

function Get-MailFolder($Source) {
    $ns = Connect-Outlook
    if ($Source.EntryID) {
        try { return $ns.GetFolderFromID($Source.EntryID, $Source.StoreID) } catch { }
    }
    # 依路徑尋找（例如 Outlook 設定檔重建後 EntryID 改變）
    $parts = @($Source.FolderPath.TrimStart('\') -split '\\')
    $f = $ns.Folders.Item($parts[0])
    for ($i = 1; $i -lt $parts.Count; $i++) { $f = $f.Folders.Item($parts[$i]) }
    return $f
}

function Get-FolderTree($Folder, [bool]$IncludeSub) {
    $list = New-Object System.Collections.ArrayList
    [void]$list.Add($Folder)
    if ($IncludeSub) {
        $tmp = New-Object System.Collections.ArrayList
        foreach ($sub in $Folder.Folders) { Add-FolderRecursive $sub $tmp 1 }
        $ns = Connect-Outlook
        foreach ($t in $tmp) { [void]$list.Add($ns.GetFolderFromID($t.EntryID, $t.StoreID)) }
    }
    return , $list
}

function Get-MailInfo($Mail) {
    $email = ''
    try {
        if ($Mail.SenderEmailType -eq 'EX') {
            $u = $Mail.Sender.GetExchangeUser()
            if ($u) { $email = [string]$u.PrimarySmtpAddress }
        } else {
            $email = [string]$Mail.SenderEmailAddress
        }
    } catch { }
    if (-not $email) { try { $email = [string]$Mail.SenderEmailAddress } catch { } }
    $domain = ''
    if ($email -match '@(.+)$') { $domain = $Matches[1] }
    $subj = [string]$Mail.Subject
    if (-not $subj.Trim()) { $subj = '(無主旨)' }
    return [pscustomobject]@{
        Item        = $Mail
        SenderName  = [string]$Mail.SenderName
        SenderEmail = $email
        Domain      = $domain
        Subject     = $subj
        Received    = [datetime]$Mail.ReceivedTime
        Categories  = [string]$Mail.Categories
        Body        = $null
    }
}

function Test-Rule($Rule, $Info, [string]$SourceId) {
    if (-not $Rule.Enabled) { return $false }
    if ($Rule.SourceId -and $Rule.SourceId -ne $SourceId) { return $false }
    $results = @()
    $useRe = [bool]$Rule.UseRegex
    if (([string]$Rule.SenderContains).Trim()) {
        $results += (Test-Keywords ("{0}`n{1}" -f $Info.SenderName, $Info.SenderEmail) $Rule.SenderContains $useRe)
    }
    if (([string]$Rule.SubjectContains).Trim()) {
        $results += (Test-Keywords $Info.Subject $Rule.SubjectContains $useRe)
    }
    if (([string]$Rule.BodyContains).Trim()) {
        if ($null -eq $Info.Body) { $Info.Body = [string]$Info.Item.Body }
        $results += (Test-Keywords $Info.Body $Rule.BodyContains $useRe)
    }
    if ($results.Count -eq 0) { return $true }      # 未設定任何條件 = 全部符合
    if ($Rule.MatchMode -eq 'Any') { return ($results -contains $true) }
    return (-not ($results -contains $false))
}

function Export-MailPdf($Mail, [string]$PdfPath) {
    if (-not $script:word) {
        $script:word = New-Object -ComObject Word.Application
        $script:word.Visible = $false
        $script:word.DisplayAlerts = 0
    }
    if (-not [System.IO.Directory]::Exists($TempDir)) { [void][System.IO.Directory]::CreateDirectory($TempDir) }
    $tmp = Join-Path $TempDir ([guid]::NewGuid().ToString() + '.mht')
    $Mail.SaveAs($tmp, $olMHTML)
    $doc = $null
    try {
        # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
        $doc = $script:word.Documents.Open($tmp, $false, $true, $false)
        $doc.ExportAsFixedFormat($PdfPath, $wdFormatPDF)
    } finally {
        if ($doc) { try { $doc.Close(0) } catch { } }
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    }
}

function Save-Attachments($Mail, [string]$Dir, [bool]$SkipInline) {
    $saved = 0; $failed = 0
    $html = $null
    $count = $Mail.Attachments.Count
    for ($i = 1; $i -le $count; $i++) {
        $a = $Mail.Attachments.Item($i)
        try {
            $hidden = $false
            try { $hidden = [bool]$a.PropertyAccessor.GetProperty($PR_ATTACH_HIDDEN) } catch { }
            if ($hidden) { continue }
            if ($SkipInline) {
                $cid = ''
                try { $cid = [string]$a.PropertyAccessor.GetProperty($PR_ATTACH_CONTENT_ID) } catch { }
                if ($cid) {
                    if ($null -eq $html) { $html = [string]$Mail.HTMLBody }
                    if ($html.IndexOf("cid:$cid", [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { continue }
                }
            }
            if (-not [System.IO.Directory]::Exists($Dir)) { [void][System.IO.Directory]::CreateDirectory($Dir) }
            $p = Get-UniqueFile $Dir (Get-SafeName ([string]$a.FileName) 100)
            $a.SaveAsFile($p)
            $saved++
        } catch { $failed++ }
    }
    return [pscustomobject]@{ Saved = $saved; Failed = $failed }
}

function Save-Mail($Info, $Rule, [string]$Dir) {
    if (-not [System.IO.Directory]::Exists($Dir)) {
        if ($Rule.CreateFolder) { [void][System.IO.Directory]::CreateDirectory($Dir) }
        else { throw "目錄不存在（此規則未允許自動建立）：$Dir" }
    }
    $opt = $script:config.Options
    $name = Expand-Template $opt.FileNameTemplate $Info $Rule.Name
    $base = Get-UniqueBase $Dir $name
    $out = @()
    if ($Rule.SavePdf) { Export-MailPdf $Info.Item "$base.pdf"; $out += 'PDF' }
    if ($Rule.SaveMsg) { $Info.Item.SaveAs("$base.msg", $olMSGUnicode); $out += 'MSG' }
    if ($Rule.SaveAttachments -and $Info.Item.Attachments.Count -gt 0) {
        $r = Save-Attachments $Info.Item "$base`_附件" ([bool]$opt.SkipInlineImages)
        if ($r.Saved) { $out += "附件 $($r.Saved) 個" }
        if ($r.Failed) { $out += "附件失敗 $($r.Failed) 個" }
    }
    if ($out.Count -eq 0) { return '（規則未勾選任何輸出）' }
    return ('已存：' + ($out -join '、') + '  →  ' + [System.IO.Path]::GetFileName($base))
}

function Complete-Mail($Info) {
    $opt = $script:config.Options
    $m = $Info.Item
    $changed = $false
    if ($opt.AddCategory -and $opt.CategoryName) {
        $cats = [string]$m.Categories
        $list = @($cats -split '\s*[,;，]\s*' | Where-Object { $_ })
        if ($list -notcontains $opt.CategoryName) {
            if ($cats) { $m.Categories = "$cats, $($opt.CategoryName)" } else { $m.Categories = $opt.CategoryName }
            $changed = $true
        }
    }
    if ($opt.MarkAsRead -and $m.UnRead) { $m.UnRead = $false; $changed = $true }
    if ($changed) { $m.Save() }
}

function Close-Office {
    if ($script:word) {
        try { $script:word.Quit() } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:word) } catch { }
        $script:word = $null
    }
}

# ---------------------------------------------------------------------------
# 介面元件輔助
# ---------------------------------------------------------------------------
$UiFont   = New-Object System.Drawing.Font('Microsoft JhengHei UI', 9)
$HintFont = New-Object System.Drawing.Font('Microsoft JhengHei UI', 8.5)
$HintColor = [System.Drawing.Color]::DimGray

function Add-Ctl($Parent, $Ctl, [int]$X, [int]$Y, [int]$W = 0, [int]$H = 0, [string]$Anchor = 'Top,Left') {
    $Ctl.Location = New-Object System.Drawing.Point($X, $Y)
    if ($W -gt 0) { $Ctl.Width = $W }
    if ($H -gt 0) { $Ctl.Height = $H }
    $Ctl.Anchor = $Anchor
    $Parent.Controls.Add($Ctl)
    return $Ctl
}

function New-Lbl([string]$Text, [switch]$Hint) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text
    $l.AutoSize = $true
    if ($Hint) { $l.Font = $HintFont; $l.ForeColor = $HintColor }
    return $l
}

function New-Btn([string]$Text) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.Height = 28
    return $b
}

function New-Txt {
    $t = New-Object System.Windows.Forms.TextBox
    return $t
}

function New-Chk([string]$Text) {
    $c = New-Object System.Windows.Forms.CheckBox
    $c.Text = $Text
    $c.AutoSize = $true
    return $c
}

function New-Combo {
    $c = New-Object System.Windows.Forms.ComboBox
    $c.DropDownStyle = 'DropDownList'
    return $c
}

function New-ListView([string[]]$Columns, [int[]]$Widths) {
    $lv = New-Object System.Windows.Forms.ListView
    $lv.View = 'Details'
    $lv.FullRowSelect = $true
    $lv.GridLines = $true
    $lv.HideSelection = $false
    for ($i = 0; $i -lt $Columns.Count; $i++) { [void]$lv.Columns.Add($Columns[$i], $Widths[$i]) }
    return $lv
}

function Invoke-Safe([scriptblock]$Block) {
    try { & $Block }
    catch { Show-Msg ('發生錯誤：' + $_.Exception.Message) 'Error' }
}

function Select-Folder([string]$Current) {
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = '請選擇存放的根目錄（之後可在路徑後面加上 {寄件者} 等代碼）'
    $root = $Current
    if ($root -and $root.IndexOf('{') -ge 0) { $root = $root.Substring(0, $root.IndexOf('{')).TrimEnd('\') }
    if ($root -and [System.IO.Directory]::Exists($root)) { $dlg.SelectedPath = $root }
    if ($dlg.ShowDialog() -eq 'OK') { return $dlg.SelectedPath }
    return $null
}

# ---------------------------------------------------------------------------
# 載入設定
# ---------------------------------------------------------------------------
Import-Config
if ($script:isNewConfig) {
    $sample = New-Rule
    $sample.Name = '範例：依寄件者分類'
    $sample.Enabled = $false
    [void]$script:savedRules.Add($sample)
    Export-Config
    $chg = Get-RuleChanges (New-Object System.Collections.ArrayList) $script:savedRules
    Add-HistoryEntry '建立初始設定' '' $chg $script:savedRules
} elseif (-not @(Get-HistoryEntries).Count) {
    $chg = Get-RuleChanges (New-Object System.Collections.ArrayList) $script:savedRules
    Add-HistoryEntry '初始版本（既有設定）' '' $chg $script:savedRules
}
$script:workRules   = Copy-Rules $script:savedRules
$script:dirty       = $false
$script:loadingRule = $false
$script:updatingList = $false
$script:loadingSrc  = $false
$script:folderList  = New-Object System.Collections.ArrayList
$script:ruleSrcIds  = @('')
$script:cancel      = $false
$script:headless    = [bool]$ProcessSelected
$script:headlessResults = New-Object System.Collections.ArrayList
$script:singleRuleIds = @()
$script:running     = $false

# ---------------------------------------------------------------------------
# 主視窗
# ---------------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "$AppName  v$AppVersion"
$form.Font = $UiFont
$form.ClientSize = New-Object System.Drawing.Size(1080, 700)
$form.MinimumSize = New-Object System.Drawing.Size(1000, 680)
$form.StartPosition = 'CenterScreen'
$form.SuspendLayout()
$form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
$form.AutoScaleMode = 'Dpi'

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'
$form.Controls.Add($tabs)

$status = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = '就緒'
[void]$status.Items.Add($statusLabel)
$form.Controls.Add($status)

function Set-Status([string]$Text) {
    $statusLabel.Text = $Text
    [System.Windows.Forms.Application]::DoEvents()
}

$tabSrc  = New-Object System.Windows.Forms.TabPage '1. 資料來源'
$tabRule = New-Object System.Windows.Forms.TabPage '2. 規則'
$tabOpt  = New-Object System.Windows.Forms.TabPage '3. 選項'
$tabRun  = New-Object System.Windows.Forms.TabPage '4. 執行'
$tabHist = New-Object System.Windows.Forms.TabPage '5. 規則修改紀錄'
foreach ($t in @($tabSrc, $tabRule, $tabOpt, $tabRun, $tabHist)) {
    # 先給定設計尺寸，讓 Anchor 以此為基準計算（之後視窗縮放時會正確跟著調整）
    $t.Size = New-Object System.Drawing.Size(1072, 648)
    $tabs.TabPages.Add($t)
}

# ===========================================================================
# 1. 資料來源
# ===========================================================================
[void](Add-Ctl $tabSrc (New-Lbl 'Outlook 資料夾：') 15 22)
$cboFolder = Add-Ctl $tabSrc (New-Combo) 120 18 720 0 'Top,Left,Right'
$cboFolder.MaxDropDownItems = 25
$btnReloadFolders = Add-Ctl $tabSrc (New-Btn '重新讀取資料夾') 850 16 160 0 'Top,Right'
$chkSrcSub = Add-Ctl $tabSrc (New-Chk '包含子資料夾') 120 54
$btnAddSrc = Add-Ctl $tabSrc (New-Btn '＋ 加入為資料來源') 260 50 170
[void](Add-Ctl $tabSrc (New-Lbl '從下拉選單選擇信箱中的資料夾（含共用信箱、封存資料檔），按「加入為資料來源」。清單中打勾＝啟用；規則可「限定來源」只套用到特定資料夾。' -Hint) 15 90)

$lvSrc = Add-Ctl $tabSrc (New-ListView @('Outlook 資料夾', '包含子資料夾', '狀態') @(700, 120, 160)) 15 115 1040 460 'Top,Bottom,Left,Right'
$lvSrc.CheckBoxes = $true
$btnRemoveSrc = Add-Ctl $tabSrc (New-Btn '移除選取的來源') 15 590 150 0 'Bottom,Left'
$btnToggleSub = Add-Ctl $tabSrc (New-Btn '切換「包含子資料夾」') 175 590 180 0 'Bottom,Left'

function Update-FolderCombo {
    Set-Status '正在讀取 Outlook 資料夾…'
    $script:folderList = Get-OutlookFolders
    $cboFolder.BeginUpdate()
    $cboFolder.Items.Clear()
    foreach ($f in $script:folderList) { [void]$cboFolder.Items.Add((Format-FolderPath $f.FolderPath)) }
    $cboFolder.EndUpdate()
    if ($cboFolder.Items.Count -gt 0) { $cboFolder.SelectedIndex = 0 }
    Update-SourceList
    Set-Status ("已讀取 {0} 個 Outlook 郵件資料夾" -f $script:folderList.Count)
}

function Update-SourceList {
    $script:loadingSrc = $true
    try {
        $lvSrc.BeginUpdate()
        $lvSrc.Items.Clear()
        foreach ($s in $script:config.Sources) {
            $it = New-Object System.Windows.Forms.ListViewItem((Format-FolderPath $s.FolderPath))
            $sub = '否'; if ($s.IncludeSubfolders) { $sub = '是' }
            [void]$it.SubItems.Add($sub)
            $state = ''
            if ($script:folderList.Count -gt 0) {
                $found = $script:folderList | Where-Object { $_.EntryID -eq $s.EntryID -or $_.FolderPath -eq $s.FolderPath } | Select-Object -First 1
                if ($found) { $state = '正常' } else { $state = '找不到資料夾' }
            }
            [void]$it.SubItems.Add($state)
            $it.Checked = [bool]$s.Enabled
            $it.Tag = $s.Id
            [void]$lvSrc.Items.Add($it)
        }
        $lvSrc.EndUpdate()
    } finally { $script:loadingSrc = $false }
    Update-RuleSourceCombo
}

$btnReloadFolders.Add_Click({ Invoke-Safe { Update-FolderCombo } })

$btnAddSrc.Add_Click({ Invoke-Safe {
    $i = $cboFolder.SelectedIndex
    if ($i -lt 0) { Show-Msg '請先從下拉選單選擇 Outlook 資料夾。'; return }
    $f = $script:folderList[$i]
    $dup = $script:config.Sources | Where-Object { $_.EntryID -eq $f.EntryID -or $_.FolderPath -eq $f.FolderPath }
    if ($dup) { Show-Msg '此資料夾已經是資料來源。'; return }
    $s = [pscustomobject]$SourceDefaults
    $s.Id = [guid]::NewGuid().ToString()
    $s.FolderPath = $f.FolderPath
    $s.EntryID = $f.EntryID
    $s.StoreID = $f.StoreID
    $s.IncludeSubfolders = $chkSrcSub.Checked
    [void]$script:config.Sources.Add($s)
    Export-Config
    Update-SourceList
    Set-Status ('已加入資料來源：' + (Format-FolderPath $f.FolderPath))
} })

$lvSrc.Add_ItemChecked({ param($sender, $e)
    if ($script:loadingSrc) { return }
    Invoke-Safe {
        $s = $script:config.Sources | Where-Object { $_.Id -eq $e.Item.Tag } | Select-Object -First 1
        if ($s) { $s.Enabled = $e.Item.Checked; Export-Config }
    }
})

$btnRemoveSrc.Add_Click({ Invoke-Safe {
    if ($lvSrc.SelectedItems.Count -eq 0) { Show-Msg '請先在清單中選取要移除的來源。'; return }
    $id = $lvSrc.SelectedItems[0].Tag
    $used = @(@($script:workRules) + @($script:savedRules) | Where-Object { $_.SourceId -eq $id } | ForEach-Object { $_.Name } | Select-Object -Unique)
    $msg = '確定移除此資料來源？'
    if ($used.Count) { $msg += "`n`n下列規則限定此來源，移除後這些規則將不會再被套用：`n" + ($used -join "`n") }
    if (-not (Confirm-Msg $msg)) { return }
    $s = $script:config.Sources | Where-Object { $_.Id -eq $id } | Select-Object -First 1
    [void]$script:config.Sources.Remove($s)
    Export-Config
    Update-SourceList
} })

$btnToggleSub.Add_Click({ Invoke-Safe {
    if ($lvSrc.SelectedItems.Count -eq 0) { Show-Msg '請先在清單中選取來源。'; return }
    $id = $lvSrc.SelectedItems[0].Tag
    $s = $script:config.Sources | Where-Object { $_.Id -eq $id } | Select-Object -First 1
    $s.IncludeSubfolders = -not [bool]$s.IncludeSubfolders
    Export-Config
    Update-SourceList
} })

# ===========================================================================
# 2. 規則
# ===========================================================================
$lstRules = Add-Ctl $tabRule (New-Object System.Windows.Forms.ListBox) 15 15 300 510 'Top,Bottom,Left'
$lstRules.IntegralHeight = $false
$btnRuleAdd  = Add-Ctl $tabRule (New-Btn '新增') 15 535 95 0 'Bottom,Left'
$btnRuleCopy = Add-Ctl $tabRule (New-Btn '複製') 117 535 95 0 'Bottom,Left'
$btnRuleDel  = Add-Ctl $tabRule (New-Btn '刪除') 219 535 96 0 'Bottom,Left'
$btnRuleUp   = Add-Ctl $tabRule (New-Btn '▲ 上移') 15 568 148 0 'Bottom,Left'
$btnRuleDown = Add-Ctl $tabRule (New-Btn '▼ 下移') 167 568 148 0 'Bottom,Left'

$grp = Add-Ctl $tabRule (New-Object System.Windows.Forms.GroupBox) 330 8 730 590 'Top,Bottom,Left,Right'
$grp.Text = '規則內容（由上而下依序比對）'

[void](Add-Ctl $grp (New-Lbl '規則名稱：') 15 30)
$txtName = Add-Ctl $grp (New-Txt) 130 26 350
$chkEnabled = Add-Ctl $grp (New-Chk '啟用此規則') 500 28

[void](Add-Ctl $grp (New-Lbl '限定來源：') 15 65)
$cboRuleSrc = Add-Ctl $grp (New-Combo) 130 61 580 0 'Top,Left,Right'

[void](Add-Ctl $grp (New-Lbl '條件組合：') 15 100)
$cboMode = Add-Ctl $grp (New-Combo) 130 96 250
[void]$cboMode.Items.Add('符合下列「全部」條件 (AND)')
[void]$cboMode.Items.Add('符合下列「任一」條件 (OR)')
$chkRegex = Add-Ctl $grp (New-Chk '使用規則運算式 (Regex)') 400 98

[void](Add-Ctl $grp (New-Lbl '寄件者包含：') 15 135)
$txtSender = Add-Ctl $grp (New-Txt) 130 131 580 0 'Top,Left,Right'
[void](Add-Ctl $grp (New-Lbl '主旨包含：') 15 170)
$txtSubject = Add-Ctl $grp (New-Txt) 130 166 580 0 'Top,Left,Right'
[void](Add-Ctl $grp (New-Lbl '本文包含：') 15 205)
$txtBody = Add-Ctl $grp (New-Txt) 130 201 580 0 'Top,Left,Right'
[void](Add-Ctl $grp (New-Lbl '多個關鍵字以「;」分隔，任一個出現即成立；寄件者會比對名稱與電子郵件地址；空白＝不檢查該條件。' -Hint) 130 228)

[void](Add-Ctl $grp (New-Lbl '存放路徑：') 15 263)
$txtPath = Add-Ctl $grp (New-Txt) 130 259 490 0 'Top,Left,Right'
$btnBrowse = Add-Ctl $grp (New-Btn '瀏覽…') 630 257 80 0 'Top,Right'
[void](Add-Ctl $grp (New-Lbl $PlaceholderHelp -Hint) 130 286)

$chkCreate = Add-Ctl $grp (New-Chk '目錄不存在時自動建立（未勾選時，目錄不存在的信件會略過並記錄）') 130 315
[void](Add-Ctl $grp (New-Lbl '輸出內容：') 15 350)
$chkPdf = Add-Ctl $grp (New-Chk '信件本文 PDF') 130 348
$chkMsg = Add-Ctl $grp (New-Chk '原始信件 .msg') 270 348
$chkAtt = Add-Ctl $grp (New-Chk '附件') 410 348
$chkStop = Add-Ctl $grp (New-Chk '符合此規則後，不再比對後面的規則') 130 380

[void](Add-Ctl $grp (New-Lbl '路徑預覽：') 15 420)
$lblPreview = Add-Ctl $grp (New-Lbl '') 130 420
$lblPreview.AutoSize = $false
$lblPreview.Size = New-Object System.Drawing.Size(580, 60)
$lblPreview.Anchor = 'Top,Left,Right'
$lblPreview.ForeColor = [System.Drawing.Color]::DarkBlue
[void](Add-Ctl $grp (New-Lbl '（以「寄件者＝範例寄件者、主旨＝範例主旨、今天日期」示範）' -Hint) 130 482)

[void](Add-Ctl $tabRule (New-Lbl '修改說明：') 15 612 0 0 'Bottom,Left')
$txtNote = Add-Ctl $tabRule (New-Txt) 90 608 380 0 'Bottom,Left'
$btnSaveRules = Add-Ctl $tabRule (New-Btn '儲存規則（寫入修改紀錄）') 480 606 230 0 'Bottom,Left'
$btnDiscard = Add-Ctl $tabRule (New-Btn '放棄未儲存的變更') 720 606 160 0 'Bottom,Left'
$lblDirty = Add-Ctl $tabRule (New-Lbl '') 890 612 0 0 'Bottom,Left'
$lblDirty.ForeColor = [System.Drawing.Color]::Firebrick

$SampleInfo = [pscustomobject]@{ SenderName = '範例寄件者'; SenderEmail = 'sample@example.gov.tw'; Domain = 'example.gov.tw'; Subject = '範例主旨'; Received = (Get-Date) }

function Set-Dirty([bool]$Value) {
    $script:dirty = $Value
    if ($Value) { $lblDirty.Text = '● 有未儲存的變更' } else { $lblDirty.Text = '' }
}

function Get-RuleListText([int]$Index, $Rule) {
    $t = '{0}. {1}' -f ($Index + 1), $Rule.Name
    if (-not $Rule.Enabled) { $t += '（停用）' }
    return $t
}

function Update-RuleList([int]$SelectIndex = -1) {
    $script:updatingList = $true
    try {
        $lstRules.BeginUpdate()
        $lstRules.Items.Clear()
        for ($i = 0; $i -lt $script:workRules.Count; $i++) { [void]$lstRules.Items.Add((Get-RuleListText $i $script:workRules[$i])) }
        $lstRules.EndUpdate()
        if ($SelectIndex -ge $lstRules.Items.Count) { $SelectIndex = $lstRules.Items.Count - 1 }
        if ($SelectIndex -lt 0 -and $lstRules.Items.Count -gt 0) { $SelectIndex = 0 }
        $lstRules.SelectedIndex = $SelectIndex
    } finally { $script:updatingList = $false }
    Show-RuleEditor
}

function Update-RuleSourceCombo([string]$ExtraId = '') {
    $prev = $script:loadingRule
    $script:loadingRule = $true
    try {
        $cboRuleSrc.Items.Clear()
        [void]$cboRuleSrc.Items.Add('(所有來源)')
        $ids = @('')
        foreach ($s in $script:config.Sources) {
            [void]$cboRuleSrc.Items.Add((Format-FolderPath $s.FolderPath))
            $ids += $s.Id
        }
        if ($ExtraId -and ($ids -notcontains $ExtraId)) {
            [void]$cboRuleSrc.Items.Add('(已移除的來源 — 此規則目前不會被套用)')
            $ids += $ExtraId
        }
        $script:ruleSrcIds = $ids
    } finally { $script:loadingRule = $prev }
    if (-not $prev) { Show-RuleEditor }
}

function Update-PathPreview {
    $i = $lstRules.SelectedIndex
    if ($i -lt 0) { $lblPreview.Text = ''; return }
    try {
        $r = $script:workRules[$i]
        $dir = Expand-Template $r.TargetPath $SampleInfo $r.Name -IsPath
        $file = Expand-Template $script:config.Options.FileNameTemplate $SampleInfo $r.Name
        $lblPreview.Text = (Join-Path $dir $file) + '.pdf'
    } catch { $lblPreview.Text = '（路徑格式有誤）' }
}

function Show-RuleEditor {
    $script:loadingRule = $true
    try {
        $i = $lstRules.SelectedIndex
        if ($i -lt 0 -or $i -ge $script:workRules.Count) {
            $grp.Enabled = $false
            foreach ($t in @($txtName, $txtSender, $txtSubject, $txtBody, $txtPath)) { $t.Text = '' }
            return
        }
        $grp.Enabled = $true
        $r = $script:workRules[$i]
        Update-RuleSourceCombo ([string]$r.SourceId)
        $txtName.Text = $r.Name
        $chkEnabled.Checked = [bool]$r.Enabled
        $idx = [array]::IndexOf($script:ruleSrcIds, [string]$r.SourceId)
        if ($idx -lt 0) { $idx = 0 }
        $cboRuleSrc.SelectedIndex = $idx
        if ($r.MatchMode -eq 'Any') { $cboMode.SelectedIndex = 1 } else { $cboMode.SelectedIndex = 0 }
        $chkRegex.Checked = [bool]$r.UseRegex
        $txtSender.Text = $r.SenderContains
        $txtSubject.Text = $r.SubjectContains
        $txtBody.Text = $r.BodyContains
        $txtPath.Text = $r.TargetPath
        $chkCreate.Checked = [bool]$r.CreateFolder
        $chkPdf.Checked = [bool]$r.SavePdf
        $chkMsg.Checked = [bool]$r.SaveMsg
        $chkAtt.Checked = [bool]$r.SaveAttachments
        $chkStop.Checked = [bool]$r.StopProcessing
    } finally { $script:loadingRule = $false }
    Update-PathPreview
}

function Update-RuleFromEditor {
    if ($script:loadingRule) { return }
    $i = $lstRules.SelectedIndex
    if ($i -lt 0) { return }
    $r = $script:workRules[$i]
    $r.Name = $txtName.Text
    $r.Enabled = $chkEnabled.Checked
    if ($cboRuleSrc.SelectedIndex -ge 0) { $r.SourceId = $script:ruleSrcIds[$cboRuleSrc.SelectedIndex] }
    if ($cboMode.SelectedIndex -eq 1) { $r.MatchMode = 'Any' } else { $r.MatchMode = 'All' }
    $r.UseRegex = $chkRegex.Checked
    $r.SenderContains = $txtSender.Text
    $r.SubjectContains = $txtSubject.Text
    $r.BodyContains = $txtBody.Text
    $r.TargetPath = $txtPath.Text
    $r.CreateFolder = $chkCreate.Checked
    $r.SavePdf = $chkPdf.Checked
    $r.SaveMsg = $chkMsg.Checked
    $r.SaveAttachments = $chkAtt.Checked
    $r.StopProcessing = $chkStop.Checked
    Set-Dirty $true
    $script:updatingList = $true
    try {
        $text = Get-RuleListText $i $r
        if ($lstRules.Items[$i] -ne $text) { $lstRules.Items[$i] = $text; $lstRules.SelectedIndex = $i }
    } finally { $script:updatingList = $false }
    Update-PathPreview
}

$lstRules.Add_SelectedIndexChanged({ if (-not $script:updatingList) { Show-RuleEditor } })
foreach ($t in @($txtName, $txtSender, $txtSubject, $txtBody, $txtPath)) { $t.Add_TextChanged({ Update-RuleFromEditor }) }
foreach ($c in @($chkEnabled, $chkRegex, $chkCreate, $chkPdf, $chkMsg, $chkAtt, $chkStop)) { $c.Add_CheckedChanged({ Update-RuleFromEditor }) }
foreach ($c in @($cboRuleSrc, $cboMode)) { $c.Add_SelectedIndexChanged({ Update-RuleFromEditor }) }

$btnBrowse.Add_Click({ Invoke-Safe {
    $p = Select-Folder $txtPath.Text
    if ($p) { $txtPath.Text = Join-Path $p '{寄件者}' }
} })

$btnRuleAdd.Add_Click({ Invoke-Safe {
    $r = New-Rule
    [void]$script:workRules.Add($r)
    Set-Dirty $true
    Update-RuleList ($script:workRules.Count - 1)
    $txtName.Focus(); $txtName.SelectAll()
} })

$btnRuleCopy.Add_Click({ Invoke-Safe {
    $i = $lstRules.SelectedIndex
    if ($i -lt 0) { return }
    $copy = (Copy-Rules @($script:workRules[$i]))[0]
    $copy.Id = [guid]::NewGuid().ToString()
    $copy.Name = $copy.Name + '（複製）'
    $script:workRules.Insert($i + 1, $copy)
    Set-Dirty $true
    Update-RuleList ($i + 1)
} })

$btnRuleDel.Add_Click({ Invoke-Safe {
    $i = $lstRules.SelectedIndex
    if ($i -lt 0) { return }
    if (-not (Confirm-Msg ("確定刪除規則「{0}」？（按「儲存規則」後才會生效，並記錄於修改紀錄）" -f $script:workRules[$i].Name))) { return }
    $script:workRules.RemoveAt($i)
    Set-Dirty $true
    Update-RuleList $i
} })

function Move-Rule([int]$Offset) {
    $i = $lstRules.SelectedIndex
    $j = $i + $Offset
    if ($i -lt 0 -or $j -lt 0 -or $j -ge $script:workRules.Count) { return }
    $tmp = $script:workRules[$i]
    $script:workRules[$i] = $script:workRules[$j]
    $script:workRules[$j] = $tmp
    Set-Dirty $true
    Update-RuleList $j
}
$btnRuleUp.Add_Click({ Invoke-Safe { Move-Rule -1 } })
$btnRuleDown.Add_Click({ Invoke-Safe { Move-Rule 1 } })

function Test-RulesValid($Rules) {
    $errors = @()
    $n = 0
    foreach ($r in $Rules) {
        $n++
        if (-not ([string]$r.Name).Trim()) { $errors += "第 $n 條規則沒有名稱。" }
        if (-not ([string]$r.TargetPath).Trim()) { $errors += "規則「$($r.Name)」沒有設定存放路徑。" }
        if ($r.UseRegex) {
            foreach ($f in @('SenderContains', 'SubjectContains', 'BodyContains')) {
                foreach ($k in (Split-Keywords $r.$f)) {
                    try { [void][regex]::new($k) } catch { $errors += "規則「$($r.Name)」的規則運算式有誤：$k" }
                }
            }
        }
    }
    return $errors
}

function Save-Rules([string]$Action, [string]$Note) {
    $errors = @(Test-RulesValid $script:workRules)
    if ($errors.Count) { Show-Msg ("無法儲存：`n`n" + ($errors -join "`n")) 'Warning'; return $false }
    $changes = Get-RuleChanges $script:savedRules $script:workRules
    if ($changes.Count -eq 0) {
        Set-Dirty $false
        Show-Msg '規則沒有任何變更，不需儲存。'
        return $false
    }
    $script:savedRules = Copy-Rules $script:workRules
    Export-Config
    Add-HistoryEntry $Action $Note $changes $script:savedRules
    Set-Dirty $false
    Update-HistoryList
    Set-Status ("規則已儲存：" + (Get-ChangeSummary $changes))
    return $true
}

$btnSaveRules.Add_Click({ Invoke-Safe {
    if (Save-Rules '儲存規則' $txtNote.Text.Trim()) { $txtNote.Text = ''; Show-Msg '規則已儲存，並已寫入修改紀錄。' }
} })

$btnDiscard.Add_Click({ Invoke-Safe {
    if (-not $script:dirty) { return }
    if (-not (Confirm-Msg '確定放棄所有未儲存的規則變更？')) { return }
    $script:workRules = Copy-Rules $script:savedRules
    Set-Dirty $false
    Update-RuleList 0
} })

# ===========================================================================
# 3. 選項
# ===========================================================================
[void](Add-Ctl $tabOpt (New-Lbl '檔名格式：') 20 25)
$txtFileTpl = Add-Ctl $tabOpt (New-Txt) 200 21 500
[void](Add-Ctl $tabOpt (New-Lbl ($PlaceholderHelp + '（副檔名自動加上；同名檔案自動加 _2、_3）') -Hint) 200 48)

[void](Add-Ctl $tabOpt (New-Lbl '只處理最近幾天的信件：') 20 85)
$numDays = Add-Ctl $tabOpt (New-Object System.Windows.Forms.NumericUpDown) 200 82 80
$numDays.Minimum = 0; $numDays.Maximum = 3650
[void](Add-Ctl $tabOpt (New-Lbl '天（0 ＝ 不限，資料夾很大時處理會較久）' -Hint) 290 86)

$chkUnread = Add-Ctl $tabOpt (New-Chk '只處理「未讀」信件') 200 120
$chkSkipCat = Add-Ctl $tabOpt (New-Chk '略過已標上此類別的信件（避免重複歸檔）：') 200 152
$txtCat = Add-Ctl $tabOpt (New-Txt) 500 150 150
$chkAddCat = Add-Ctl $tabOpt (New-Chk '歸檔完成後，在 Outlook 信件加上上述類別') 200 184
$chkMarkRead = Add-Ctl $tabOpt (New-Chk '歸檔完成後，將信件標示為已讀') 200 216
$chkSkipInline = Add-Ctl $tabOpt (New-Chk '保存附件時，略過內嵌在本文中的圖片（如簽名檔圖示）') 200 248

[void](Add-Ctl $tabOpt (New-Lbl '未符合任何規則的信件：') 20 292)
$txtDefault = Add-Ctl $tabOpt (New-Txt) 200 288 500
$btnDefBrowse = Add-Ctl $tabOpt (New-Btn '瀏覽…') 710 286 80
[void](Add-Ctl $tabOpt (New-Lbl '存到這個目錄（可用代碼，會自動建立）；空白＝不處理，只在結果中列出。' -Hint) 200 315)

$btnSaveOpt = Add-Ctl $tabOpt (New-Btn '儲存選項') 200 360 140

function Show-Options {
    $o = $script:config.Options
    $txtFileTpl.Text = $o.FileNameTemplate
    $numDays.Value = [decimal][int]$o.DaysBack
    $chkUnread.Checked = [bool]$o.UnreadOnly
    $chkSkipCat.Checked = [bool]$o.SkipCategorized
    $txtCat.Text = $o.CategoryName
    $chkAddCat.Checked = [bool]$o.AddCategory
    $chkMarkRead.Checked = [bool]$o.MarkAsRead
    $chkSkipInline.Checked = [bool]$o.SkipInlineImages
    $txtDefault.Text = $o.DefaultTargetPath
}

$btnDefBrowse.Add_Click({ Invoke-Safe { $p = Select-Folder $txtDefault.Text; if ($p) { $txtDefault.Text = Join-Path $p '未分類' } } })

$btnSaveOpt.Add_Click({ Invoke-Safe {
    if (-not $txtFileTpl.Text.Trim()) { Show-Msg '檔名格式不可空白。' 'Warning'; return }
    $o = $script:config.Options
    $o.FileNameTemplate = $txtFileTpl.Text.Trim()
    $o.DaysBack = [int]$numDays.Value
    $o.UnreadOnly = $chkUnread.Checked
    $o.SkipCategorized = $chkSkipCat.Checked
    $o.CategoryName = $txtCat.Text.Trim()
    $o.AddCategory = $chkAddCat.Checked
    $o.MarkAsRead = $chkMarkRead.Checked
    $o.SkipInlineImages = $chkSkipInline.Checked
    $o.DefaultTargetPath = $txtDefault.Text.Trim()
    Export-Config
    Update-PathPreview
    Show-Msg '選項已儲存。'
} })

# ===========================================================================
# 4. 執行
# ===========================================================================
$btnPreview = Add-Ctl $tabRun (New-Btn '預覽（不存檔）') 15 15 160
$btnRun     = Add-Ctl $tabRun (New-Btn '▶ 開始歸檔') 185 15 140
$btnStop    = Add-Ctl $tabRun (New-Btn '■ 停止') 335 15 90
$btnStop.Enabled = $false
$btnOpenLog = Add-Ctl $tabRun (New-Btn '開啟處理紀錄資料夾') 880 15 170 0 'Top,Right'
[void](Add-Ctl $tabRun (New-Lbl '↑ 處理「1. 資料來源」中所有啟用的資料夾。' -Hint) 435 22)

$grpSingle = Add-Ctl $tabRun (New-Object System.Windows.Forms.GroupBox) 15 52 1040 92 'Top,Left,Right'
$grpSingle.Text = '單一信件（只處理指定的信件）'
[void](Add-Ctl $grpSingle (New-Lbl '套用方式：') 12 27)
$cboSingleMode = Add-Ctl $grpSingle (New-Combo) 85 23 330
$chkSinglePreview = Add-Ctl $grpSingle (New-Chk '只預覽') 430 26
$chkSingleRedo = Add-Ctl $grpSingle (New-Chk '已歸檔過的信件也重新處理') 510 26
$btnSelMail = Add-Ctl $grpSingle (New-Btn '處理 Outlook 選取的信件') 720 21 170
$btnMsgFile = Add-Ctl $grpSingle (New-Btn '開啟 .msg 檔處理…') 896 21 132
[void](Add-Ctl $grpSingle (New-Lbl '先在 Outlook 選取一封或多封信件（或開啟信件視窗）再按按鈕。「套用規則：…」直接使用該規則的路徑與輸出設定，不檢查條件。' -Hint) 12 60)

[void](Add-Ctl $tabRun (New-Lbl '預覽使用「目前編輯中（含未儲存）」的規則；正式歸檔使用「已儲存」的規則。每次正式歸檔的結果會存在 logs 資料夾（CSV）。' -Hint) 15 152)
$progress = Add-Ctl $tabRun (New-Object System.Windows.Forms.ProgressBar) 15 172 1040 18 'Top,Left,Right'
$lvResult = Add-Ctl $tabRun (New-ListView @('結果', '收件時間', '寄件者', '主旨', '規則', '目的地 / 說明') @(70, 120, 160, 260, 140, 400)) 15 196 1040 424 'Top,Bottom,Left,Right'

function Update-SingleModeCombo {
    $prev = $cboSingleMode.SelectedIndex
    $prevId = ''
    if ($prev -ge 2 -and ($prev - 2) -lt $script:singleRuleIds.Count) { $prevId = $script:singleRuleIds[$prev - 2] }
    $cboSingleMode.Items.Clear()
    [void]$cboSingleMode.Items.Add('依規則自動比對')
    [void]$cboSingleMode.Items.Add('手動選擇存放目錄…（存 PDF、MSG、附件）')
    $ids = @()
    $src = $script:savedRules
    if ($chkSinglePreview.Checked) { $src = $script:workRules }
    foreach ($r in $src) {
        $t = '套用規則：' + $r.Name
        if (-not $r.Enabled) { $t += '（停用中）' }
        [void]$cboSingleMode.Items.Add($t)
        $ids += $r.Id
    }
    $script:singleRuleIds = $ids
    $idx = 0
    if ($prev -eq 1) { $idx = 1 }
    elseif ($prevId) { $k = [array]::IndexOf($ids, $prevId); if ($k -ge 0) { $idx = $k + 2 } }
    $cboSingleMode.SelectedIndex = $idx
}

function Add-Result([string]$State, $Info, [string]$RuleName, [string]$Detail) {
    if ($script:headless) {
        $subj = ''; if ($Info) { $subj = $Info.Subject }
        $line = "[$State] $subj"
        if ($RuleName) { $line += "（$RuleName）" }
        [void]$script:headlessResults.Add("$line`n      $Detail")
        return
    }
    $it = New-Object System.Windows.Forms.ListViewItem($State)
    $time = ''; $sender = ''; $subject = ''
    if ($Info) { $time = $Info.Received.ToString('yyyy-MM-dd HH:mm'); $sender = $Info.SenderName; $subject = $Info.Subject }
    foreach ($v in @($time, $sender, $subject, $RuleName, $Detail)) { [void]$it.SubItems.Add([string]$v) }
    switch ($State) {
        '完成'   { $it.ForeColor = [System.Drawing.Color]::DarkGreen }
        '錯誤'   { $it.ForeColor = [System.Drawing.Color]::Firebrick }
        '略過'   { $it.ForeColor = [System.Drawing.Color]::DarkOrange }
        '未符合' { $it.ForeColor = [System.Drawing.Color]::Gray }
    }
    [void]$lvResult.Items.Add($it)
    $it.EnsureVisible()
}

function Set-Running([bool]$Running) {
    $script:running = $Running
    $btnPreview.Enabled = -not $Running
    $btnRun.Enabled = -not $Running
    $btnStop.Enabled = $Running
    $grpSingle.Enabled = -not $Running
    foreach ($p in @($tabSrc, $tabRule, $tabOpt, $tabHist)) { $p.Enabled = -not $Running }
}

# 取得本次要使用的規則：預覽用「編輯中」的規則，正式歸檔用「已儲存」的規則。回傳 $null 表示取消。
function Get-RulesForRun([bool]$DryRun) {
    if ($DryRun) { $rules = $script:workRules }
    else {
        if ($script:dirty -and -not $script:headless) {
            if (Confirm-Msg '規則有未儲存的變更。是否先儲存規則再開始歸檔？（選「否」則以上次儲存的規則執行）') {
                if (-not (Save-Rules '儲存規則（歸檔前）' '')) { if ($script:dirty) { return $null } }
            }
        }
        $rules = $script:savedRules
    }
    $errors = @(Test-RulesValid $rules)
    if ($errors.Count) { Show-Msg ("規則有誤：`n`n" + ($errors -join "`n")) 'Warning'; return $null }
    return , $rules
}

function New-RunContext([bool]$DryRun) {
    $mode = '預覽'; if (-not $DryRun) { $mode = '歸檔' }
    $rows = New-Object System.Collections.ArrayList
    return [pscustomobject]@{
        DryRun          = $DryRun
        Mode            = $mode
        RunTime         = (Get-Date)
        LogRows         = $rows
        Stat            = @{ Done = 0; Skip = 0; NoMatch = 0; Error = 0; Total = 0 }
        SkipCategorized = [bool]$script:config.Options.SkipCategorized
        ShowSkipped     = $false     # 是否在結果清單列出「已歸檔而略過」的信件
        ForceRule       = $null      # 指定套用的規則（不比對條件）
    }
}

# 處理一封信件（資料夾批次與單一信件共用）
function Invoke-OneMail($Mail, [string]$SourceId, [string]$FolderLabel, $Rules, $Ctx, [bool]$AllowUpdate = $true) {
    $opt = $script:config.Options
    $Ctx.Stat.Total++
    $info = $null
    try {
        $info = Get-MailInfo $Mail
        if ($Ctx.SkipCategorized -and $opt.CategoryName) {
            $cats = @($info.Categories -split '\s*[,;，]\s*')
            if ($cats -contains $opt.CategoryName) {
                $Ctx.Stat.Skip++
                if ($Ctx.ShowSkipped) { Add-Result '略過' $info '' ("已有類別「{0}」（已歸檔過）" -f $opt.CategoryName) }
                return
            }
        }

        if ($Ctx.ForceRule) {
            $matched = @($Ctx.ForceRule)
        } else {
            $matched = @()
            foreach ($r in $Rules) {
                if (Test-Rule $r $info $SourceId) { $matched += $r; if ($r.StopProcessing) { break } }
            }
            if ($matched.Count -eq 0) {
                if ($opt.DefaultTargetPath) {
                    $d = [pscustomobject]$RuleDefaults
                    $d.Name = '(未符合規則－預設目錄)'; $d.TargetPath = $opt.DefaultTargetPath
                    $matched = @($d)
                } else {
                    $why = '沒有符合的規則'
                    if (-not $SourceId) { $why += '（此信件不在任何資料來源內，「限定來源」的規則不會套用）' }
                    Add-Result '未符合' $info '' $why
                    $Ctx.Stat.NoMatch++
                    return
                }
            }
        }

        $anyOk = $false
        foreach ($r in $matched) {
            $dir = Expand-Template $r.TargetPath $info $r.Name -IsPath
            if ($Ctx.DryRun) {
                $file = Expand-Template $opt.FileNameTemplate $info $r.Name
                $note = ''
                if (-not [System.IO.Directory]::Exists($dir)) {
                    if ($r.CreateFolder) { $note = '（將建立新目錄）' } else { $note = '（目錄不存在，將略過）' }
                }
                Add-Result '預覽' $info $r.Name ("$dir\$file $note")
                $Ctx.Stat.Done++
                continue
            }
            try {
                $detail = Save-Mail $info $r $dir
                Add-Result '完成' $info $r.Name ("$dir  ｜  $detail")
                [void]$Ctx.LogRows.Add((New-LogRow $Ctx.RunTime $Ctx.Mode $FolderLabel $info $r.Name $dir '完成' $detail))
                $Ctx.Stat.Done++; $anyOk = $true
            } catch {
                $state = '錯誤'
                if ($_.Exception.Message -like '目錄不存在*') { $state = '略過'; $Ctx.Stat.Skip++ } else { $Ctx.Stat.Error++ }
                Add-Result $state $info $r.Name $_.Exception.Message
                [void]$Ctx.LogRows.Add((New-LogRow $Ctx.RunTime $Ctx.Mode $FolderLabel $info $r.Name $dir $state $_.Exception.Message))
            }
        }
        if ($anyOk -and $AllowUpdate) { try { Complete-Mail $info } catch { } }
    } catch {
        $Ctx.Stat.Error++
        Add-Result '錯誤' $info '' $_.Exception.Message
        if ($info) { [void]$Ctx.LogRows.Add((New-LogRow $Ctx.RunTime $Ctx.Mode $FolderLabel $info '' '' '錯誤' $_.Exception.Message)) }
    }
}

# 結束一次執行：關閉 Word、寫入處理紀錄，回傳摘要文字
function Complete-Run($Ctx) {
    Close-Office
    if ($Ctx.LogRows.Count -gt 0) {
        try {
            if (-not (Test-Path -LiteralPath $LogDir)) { [void][System.IO.Directory]::CreateDirectory($LogDir) }
            $logFile = Join-Path $LogDir ('處理紀錄_{0}.csv' -f $Ctx.RunTime.ToString('yyyyMMdd'))
            $Ctx.LogRows | Export-Csv -LiteralPath $logFile -Append -NoTypeInformation -Encoding UTF8
        } catch { Show-Msg ('處理紀錄寫入失敗：' + $_.Exception.Message) 'Warning' }
    }
    $st = $Ctx.Stat
    $verb = '已歸檔'; if ($Ctx.DryRun) { $verb = '預計存放' }
    $msg = "{0}完成：共檢查 {1} 封；{2} {3} 筆、略過 {4} 封、未符合 {5} 封、錯誤 {6} 筆" -f $Ctx.Mode, $st.Total, $verb, $st.Done, $st.Skip, $st.NoMatch, $st.Error
    if ($script:cancel) { $msg = '（已手動停止）' + $msg }
    Set-Status $msg
    return $msg
}

function Invoke-Processing([bool]$DryRun) {
    $sources = @($script:config.Sources | Where-Object { $_.Enabled })
    if ($sources.Count -eq 0) { Show-Msg '尚未設定（或未啟用）任何資料來源，請先到「1. 資料來源」加入 Outlook 資料夾。' 'Warning'; return }
    $rules = Get-RulesForRun $DryRun
    if ($null -eq $rules) { return }

    $opt = $script:config.Options
    $cutoff = [datetime]::MinValue
    if ([int]$opt.DaysBack -gt 0) { $cutoff = (Get-Date).Date.AddDays(-[int]$opt.DaysBack) }
    $ctx = New-RunContext $DryRun

    $lvResult.Items.Clear()
    $script:cancel = $false
    Set-Running $true
    try {
        foreach ($src in $sources) {
            if ($script:cancel) { break }
            try { $root = Get-MailFolder $src }
            catch { Add-Result '錯誤' $null '' ('找不到資料夾：' + (Format-FolderPath $src.FolderPath)); $ctx.Stat.Error++; continue }

            foreach ($folder in (Get-FolderTree $root ([bool]$src.IncludeSubfolders))) {
                if ($script:cancel) { break }
                $fpath = Format-FolderPath ([string]$folder.FolderPath)
                Set-Status "讀取：$fpath"
                $items = $folder.Items
                if ($opt.UnreadOnly) { $items = $items.Restrict('[UnRead] = True') }
                $items.Sort('[ReceivedTime]', $true)
                $mails = New-Object System.Collections.ArrayList
                foreach ($it in $items) {
                    if ($script:cancel) { break }
                    try {
                        if ($it.Class -ne $olMailClass) { continue }
                        if ([datetime]$it.ReceivedTime -lt $cutoff) { break }
                        [void]$mails.Add($it)
                    } catch { }
                    if (($mails.Count % 50) -eq 0) { [System.Windows.Forms.Application]::DoEvents() }
                }

                $progress.Value = 0
                $progress.Maximum = [Math]::Max(1, $mails.Count)
                $n = 0
                foreach ($m in $mails) {
                    [System.Windows.Forms.Application]::DoEvents()
                    if ($script:cancel) { break }
                    $n++
                    $progress.Value = [Math]::Min($n, $progress.Maximum)
                    Set-Status ("{0}：{1} / {2}" -f $fpath, $n, $mails.Count)
                    Invoke-OneMail $m $src.Id $fpath $rules $ctx $true
                }
            }
        }
    } finally {
        Set-Running $false
        $msg = Complete-Run $ctx
    }
    Show-Msg $msg
}

# ---------------------------------------------------------------------------
# 單一信件
# ---------------------------------------------------------------------------

# 找出信件所在資料夾屬於哪個資料來源（含「包含子資料夾」的上層來源），找不到回傳空字串
function Get-SourceIdForFolder($Folder) {
    $f = $Folder
    for ($depth = 0; $depth -lt 20 -and $f; $depth++) {
        try { $eid = [string]$f.EntryID; $path = [string]$f.FolderPath } catch { break }
        foreach ($s in $script:config.Sources) {
            if (($s.EntryID -eq $eid -or $s.FolderPath -eq $path) -and ($depth -eq 0 -or $s.IncludeSubfolders)) { return [string]$s.Id }
        }
        try { $f = $f.Parent; if ($f.Class -ne 2) { break } } catch { break }   # 2 = olFolder
    }
    return ''
}

function New-SingleItem($Mail, [bool]$FromFile, [string]$FilePath = '') {
    $srcId = ''; $label = $FilePath
    if (-not $FromFile) {
        try { $srcId = Get-SourceIdForFolder $Mail.Parent; $label = Format-FolderPath ([string]$Mail.Parent.FolderPath) } catch { }
    }
    return [pscustomobject]@{ Mail = $Mail; SourceId = $srcId; Label = $label; FromFile = $FromFile }
}

# Outlook 目前選取的信件；若最上層是已開啟的信件視窗，則取該封信件
function Get-OutlookSelection {
    [void](Connect-Outlook)
    $list = New-Object System.Collections.ArrayList
    $win = $null
    try { $win = $script:outlook.ActiveWindow() } catch { }
    if ($win -and $win.Class -eq 35) {                 # 35 = olInspector（開啟中的信件視窗）
        $it = $win.CurrentItem
        if ($it.Class -eq $olMailClass) { [void]$list.Add((New-SingleItem $it $false)) }
        return , $list
    }
    $exp = $null
    try { $exp = $script:outlook.ActiveExplorer() } catch { }
    if (-not $exp) { throw '找不到開啟中的 Outlook 視窗，請先開啟 Outlook 並選取信件。' }
    $sel = $exp.Selection
    for ($i = 1; $i -le $sel.Count; $i++) {
        $it = $sel.Item($i)
        if ($it.Class -eq $olMailClass) { [void]$list.Add((New-SingleItem $it $false)) }
    }
    return , $list
}

function Invoke-SingleMail([string]$Kind) {
    $dry = $chkSinglePreview.Checked
    $rules = Get-RulesForRun $dry
    if ($null -eq $rules) { return }
    $ctx = New-RunContext $dry
    $ctx.SkipCategorized = -not $chkSingleRedo.Checked
    $ctx.ShowSkipped = $true

    # 套用方式
    $sel = $cboSingleMode.SelectedIndex
    if ($sel -eq 1) {
        $p = Select-Folder ''
        if (-not $p) { return }
        $fr = [pscustomobject]$RuleDefaults
        $fr.Name = '(手動指定目錄)'; $fr.TargetPath = $p; $fr.SaveAttachments = $true
        $ctx.ForceRule = $fr
    } elseif ($sel -ge 2) {
        $id = $script:singleRuleIds[$sel - 2]
        $fr = $rules | Where-Object { $_.Id -eq $id } | Select-Object -First 1
        if (-not $fr) { Show-Msg '找不到這條規則（可能尚未儲存），請先儲存規則或改用「只預覽」。' 'Warning'; return }
        $ctx.ForceRule = $fr
    }

    # 取得信件
    $items = New-Object System.Collections.ArrayList
    $opened = New-Object System.Collections.ArrayList
    if ($Kind -eq 'Selection') {
        $items = Get-OutlookSelection
        if ($items.Count -eq 0) { Show-Msg '請先在 Outlook 選取（或開啟）要處理的信件。'; return }
    } else {
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = 'Outlook 信件 (*.msg)|*.msg'
        $dlg.Multiselect = $true
        $dlg.Title = '選擇要歸檔的信件檔'
        if ($dlg.ShowDialog() -ne 'OK') { return }
        $ns = Connect-Outlook
        foreach ($f in $dlg.FileNames) {
            try {
                $m = $ns.OpenSharedItem($f)
                [void]$opened.Add($m)
                if ($m.Class -eq $olMailClass) { [void]$items.Add((New-SingleItem $m $true $f)) }
                else { Show-Msg ("不是一般郵件，已略過：$f") 'Warning' }
            } catch { Show-Msg ("無法開啟 $f`n`n" + $_.Exception.Message) 'Warning' }
        }
        if ($items.Count -eq 0) { return }
    }

    if (-not $dry) {
        $how = $cboSingleMode.Text
        if ($sel -eq 1) { $how = '存到 ' + $ctx.ForceRule.TargetPath }
        if (-not (Confirm-Msg ("確定歸檔 {0} 封信件？`n套用方式：{1}" -f $items.Count, $how))) { foreach ($m in $opened) { try { $m.Close(1) } catch { } }; return }
    }

    $lvResult.Items.Clear()
    $script:cancel = $false
    Set-Running $true
    try {
        $progress.Value = 0
        $progress.Maximum = [Math]::Max(1, $items.Count)
        $n = 0
        foreach ($x in $items) {
            [System.Windows.Forms.Application]::DoEvents()
            if ($script:cancel) { break }
            $n++
            $progress.Value = [Math]::Min($n, $progress.Maximum)
            Set-Status ("單一信件：{0} / {1}" -f $n, $items.Count)
            # .msg 檔不在 Outlook 信箱中，不加類別、不標已讀
            Invoke-OneMail $x.Mail $x.SourceId $x.Label $rules $ctx (-not $x.FromFile)
        }
    } finally {
        foreach ($m in $opened) { try { $m.Close(1) } catch { } }   # 1 = olDiscard
        Set-Running $false
        $msg = Complete-Run $ctx
    }
    Show-Msg $msg
}

function New-LogRow($RunTime, $Mode, $Folder, $Info, $RuleName, $Dir, $State, $Detail) {
    return [pscustomobject][ordered]@{
        '執行時間'   = $RunTime.ToString('yyyy-MM-dd HH:mm:ss')
        '執行者'     = "$env:USERDOMAIN\$env:USERNAME"
        '模式'       = $Mode
        '來源資料夾' = $Folder
        '收件時間'   = $Info.Received.ToString('yyyy-MM-dd HH:mm:ss')
        '寄件者'     = $Info.SenderName
        '寄件者信箱' = $Info.SenderEmail
        '主旨'       = $Info.Subject
        '規則'       = $RuleName
        '目的地'     = $Dir
        '結果'       = $State
        '說明'       = $Detail
    }
}

$btnPreview.Add_Click({ Invoke-Safe { Invoke-Processing $true } })
$btnRun.Add_Click({ Invoke-Safe {
    if (-not (Confirm-Msg '確定開始歸檔？信件將依「已儲存」的規則存到本機目錄。建議先執行「預覽」確認結果。')) { return }
    Invoke-Processing $false
} })
$btnSelMail.Add_Click({ Invoke-Safe { Invoke-SingleMail 'Selection' } })
$btnMsgFile.Add_Click({ Invoke-Safe { Invoke-SingleMail 'File' } })
$cboSingleMode.Add_DropDown({ Invoke-Safe { Update-SingleModeCombo } })
$chkSinglePreview.Add_CheckedChanged({ Invoke-Safe { Update-SingleModeCombo } })
$tabs.Add_SelectedIndexChanged({ if ($tabs.SelectedTab -eq $tabRun) { Invoke-Safe { Update-SingleModeCombo } } })
$btnStop.Add_Click({ $script:cancel = $true; Set-Status '正在停止…' })
$btnOpenLog.Add_Click({ Invoke-Safe {
    if (-not (Test-Path -LiteralPath $LogDir)) { [void][System.IO.Directory]::CreateDirectory($LogDir) }
    Start-Process explorer.exe $LogDir
} })

# ===========================================================================
# 5. 規則修改紀錄
# ===========================================================================
[void](Add-Ctl $tabHist (New-Lbl '篩選（規則名稱／使用者／說明）：') 15 20)
$txtHistFilter = Add-Ctl $tabHist (New-Txt) 220 16 250
$btnHistRefresh = Add-Ctl $tabHist (New-Btn '重新整理') 480 14 100
$btnHistExport = Add-Ctl $tabHist (New-Btn '匯出全部紀錄 (CSV)') 590 14 170
$btnHistOpen = Add-Ctl $tabHist (New-Btn '開啟紀錄資料夾') 770 14 140

$lvHist = Add-Ctl $tabHist (New-ListView @('時間', '使用者', '動作', '摘要', '修改說明') @(150, 170, 190, 260, 260)) 15 50 1040 250 'Top,Left,Right'
$txtHistDetail = New-Txt
$txtHistDetail.Multiline = $true
$txtHistDetail = Add-Ctl $tabHist $txtHistDetail 15 310 1040 270 'Top,Bottom,Left,Right'
$txtHistDetail.ReadOnly = $true
$txtHistDetail.ScrollBars = 'Both'
$txtHistDetail.WordWrap = $false
$txtHistDetail.Font = New-Object System.Drawing.Font('Microsoft JhengHei UI', 9.5)
$btnRestore = Add-Ctl $tabHist (New-Btn '↺ 將規則還原為選取的版本') 15 592 230 0 'Bottom,Left'
[void](Add-Ctl $tabHist (New-Lbl '紀錄只會新增、不會被覆寫；還原動作本身也會留下一筆紀錄。' -Hint) 260 598 0 0 'Bottom,Left')

function Update-HistoryList {
    $filter = $txtHistFilter.Text.Trim()
    $lvHist.BeginUpdate()
    $lvHist.Items.Clear()
    foreach ($e in (Get-HistoryEntries)) {
        if ($filter) {
            $hay = ($e.User + ' ' + $e.Note + ' ' + $e.Action + ' ' + ((@($e.Changes) | ForEach-Object { $_.RuleName }) -join ' ') + ' ' + ((@($e.Rules) | ForEach-Object { $_.Name }) -join ' '))
            if ($hay.IndexOf($filter, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
        }
        $it = New-Object System.Windows.Forms.ListViewItem([string]$e.Time)
        foreach ($v in @($e.User, $e.Action, $e.Summary, $e.Note)) { [void]$it.SubItems.Add([string]$v) }
        $it.Tag = $e
        [void]$lvHist.Items.Add($it)
    }
    $lvHist.EndUpdate()
    $txtHistDetail.Text = ''
}

$lvHist.Add_SelectedIndexChanged({
    if ($lvHist.SelectedItems.Count -gt 0) { $txtHistDetail.Text = Format-HistoryDetail $lvHist.SelectedItems[0].Tag }
})
$btnHistRefresh.Add_Click({ Invoke-Safe { Update-HistoryList } })
$txtHistFilter.Add_KeyDown({ param($s, $e) if ($e.KeyCode -eq 'Enter') { Invoke-Safe { Update-HistoryList } } })
$btnHistOpen.Add_Click({ Invoke-Safe {
    if (-not (Test-Path -LiteralPath $HistoryDir)) { [void][System.IO.Directory]::CreateDirectory($HistoryDir) }
    Start-Process explorer.exe $HistoryDir
} })

$btnHistExport.Add_Click({ Invoke-Safe {
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'CSV 檔案 (*.csv)|*.csv'
    $dlg.FileName = '規則修改紀錄_{0}.csv' -f (Get-Date).ToString('yyyyMMdd')
    if ($dlg.ShowDialog() -ne 'OK') { return }
    $rows = foreach ($e in (Get-HistoryEntries)) {
        $chg = @($e.Changes)
        if ($chg.Count -eq 0) {
            [pscustomobject][ordered]@{ '時間' = $e.Time; '使用者' = $e.User; '電腦' = $e.Computer; '動作' = $e.Action; '修改說明' = $e.Note; '變更類型' = ''; '規則' = ''; '欄位' = ''; '修改前' = ''; '修改後' = '' }
        }
        foreach ($c in $chg) {
            [pscustomobject][ordered]@{ '時間' = $e.Time; '使用者' = $e.User; '電腦' = $e.Computer; '動作' = $e.Action; '修改說明' = $e.Note; '變更類型' = $c.Action; '規則' = $c.RuleName; '欄位' = $c.Field; '修改前' = $c.Before; '修改後' = $c.After }
        }
    }
    @($rows) | Export-Csv -LiteralPath $dlg.FileName -NoTypeInformation -Encoding UTF8
    Show-Msg ('已匯出：' + $dlg.FileName)
} })

$btnRestore.Add_Click({ Invoke-Safe {
    if ($lvHist.SelectedItems.Count -eq 0) { Show-Msg '請先選取要還原的版本。'; return }
    $e = $lvHist.SelectedItems[0].Tag
    $msg = "確定將規則還原為「$($e.Time)」的版本？"
    if ($script:dirty) { $msg += "`n`n注意：目前尚未儲存的規則變更將被捨棄。" }
    if (-not (Confirm-Msg $msg)) { return }
    $backup = $script:workRules
    $script:workRules = Copy-Rules $e.Rules
    if (Save-Rules ("還原版本（{0}）" -f $e.Time) ("由 $($e.Time) 版本還原")) {
        Update-RuleList 0
        Show-Msg '規則已還原，並已寫入修改紀錄。'
    } else {
        $script:workRules = $backup
    }
} })

# ===========================================================================
# 不開管理介面：直接處理 Outlook 中選取的信件（ArchiveSelected.bat）
# ===========================================================================
if ($script:headless) {
    try {
        $rules = Get-RulesForRun ([bool]$Preview)
        if ($null -ne $rules) {
            $items = Get-OutlookSelection
            if ($items.Count -eq 0) {
                Show-Msg '請先在 Outlook 選取（或開啟）要歸檔的信件，再執行一次。'
            } else {
                $ctx = New-RunContext ([bool]$Preview)
                $ctx.ShowSkipped = $true
                if ($Force) { $ctx.SkipCategorized = $false }
                foreach ($x in $items) { Invoke-OneMail $x.Mail $x.SourceId $x.Label $rules $ctx $true }
                $msg = Complete-Run $ctx
                $lines = @($script:headlessResults)
                $detail = ($lines | Select-Object -First 15) -join "`n"
                if ($lines.Count -gt 15) { $detail += "`n…（其餘 $($lines.Count - 15) 筆請見 logs 資料夾）" }
                $icon = 'Information'
                if ($ctx.Stat.Error -gt 0) { $icon = 'Warning' }
                Show-Msg ($msg + "`n`n" + $detail) $icon
            }
        }
    } catch {
        Show-Msg ('發生錯誤：' + $_.Exception.Message) 'Error'
    } finally {
        Close-Office
    }
    $form.Dispose()
    exit
}

# ===========================================================================
# 啟動與關閉
# ===========================================================================
# 程式視窗開啟後，把黑色主控台視窗縮到工作列（失敗時不影響使用）
function Hide-ConsoleWindow {
    try {
        if (-not ('MailSorterNative.ConsoleWin' -as [type])) {
            Add-Type -Namespace MailSorterNative -Name ConsoleWin -MemberDefinition @'
[DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@
        }
        [void][MailSorterNative.ConsoleWin]::ShowWindow([MailSorterNative.ConsoleWin]::GetConsoleWindow(), 6)   # 6 = 最小化
    } catch { }
}

$form.Add_Shown({
    Hide-ConsoleWindow
    $form.Activate()
    Show-Options
    Update-SourceList
    Update-RuleList 0
    Update-HistoryList
    Update-SingleModeCombo
    Set-Dirty $false
    try { Update-FolderCombo }
    catch { Set-Status '無法連線 Outlook'; Show-Msg ("無法讀取 Outlook 資料夾，請確認 Outlook 已安裝並設定好信箱。`n`n" + $_.Exception.Message) 'Warning' }
})

$form.Add_FormClosing({ param($s, $e)
    if ($script:running) {
        if (Confirm-Msg '正在處理信件，是否停止處理？（停止後請再按一次關閉）') { $script:cancel = $true }
        $e.Cancel = $true
        return
    }
    if ($script:dirty) {
        $ans = [System.Windows.Forms.MessageBox]::Show('規則有未儲存的變更，是否儲存？', $AppName, 'YesNoCancel', 'Question')
        if ($ans -eq 'Cancel') { $e.Cancel = $true; return }
        if ($ans -eq 'Yes') {
            try { [void](Save-Rules '儲存規則（關閉程式時）' '') } catch { Show-Msg $_.Exception.Message 'Error'; $e.Cancel = $true; return }
        }
    }
    Close-Office
})

$form.ResumeLayout($false)
[void]$form.ShowDialog()
$form.Dispose()

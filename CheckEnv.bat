@echo off
rem MailSorter environment check
title MailSorter CheckEnv
cd /d "%~dp0"
echo ==== Quick check ====
powershell.exe -NoProfile -Command "Write-Host ('LanguageMode : ' + $ExecutionContext.SessionState.LanguageMode); Get-ExecutionPolicy -List | Out-String | Write-Host; Get-ChildItem -LiteralPath '%~dp0' -Filter *.ps1 | ForEach-Object { $z = Get-Item -LiteralPath $_.FullName -Stream Zone.Identifier -ErrorAction SilentlyContinue; Write-Host ($_.Name + ' : blocked-from-internet = ' + [bool]$z) }"
echo.
echo ==== Full check ====
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File "%~dp0CheckEnv.ps1"
if errorlevel 1 pause

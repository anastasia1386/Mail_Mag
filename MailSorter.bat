@echo off
rem MailSorter launcher (no installation required)
cd /d "%~dp0"
start "" powershell.exe -NoProfile -STA -ExecutionPolicy RemoteSigned -WindowStyle Minimized -File "%~dp0MailSorter.ps1"

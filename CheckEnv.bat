@echo off
rem MailSorter environment check
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File "%~dp0CheckEnv.ps1"

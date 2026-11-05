@echo off
rem Archive the mail(s) currently selected in Outlook, using saved rules (no GUI)
cd /d "%~dp0"
start "" powershell.exe -NoProfile -STA -ExecutionPolicy RemoteSigned -WindowStyle Minimized -File "%~dp0MailSorter.ps1" -ProcessSelected

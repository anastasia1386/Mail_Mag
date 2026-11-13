@echo off
rem Archive the mail(s) currently selected in Outlook, using saved rules (no GUI)
title MailSorter
cd /d "%~dp0"
powershell.exe -NoProfile -STA -ExecutionPolicy RemoteSigned -WindowStyle Minimized -File "%~dp0MailSorter.ps1" -ProcessSelected
if errorlevel 1 (
  echo.
  echo  MailSorter could not run. Please take a screenshot of this window.
  echo  Details may also be in logs\startup_error.log
  pause
)

@echo off
rem MailSorter launcher (no installation required)
title MailSorter
cd /d "%~dp0"
powershell.exe -NoProfile -STA -ExecutionPolicy RemoteSigned -File "%~dp0MailSorter.ps1"
if errorlevel 1 (
  echo.
  echo ============================================================
  echo  MailSorter could not start. Please take a screenshot of this
  echo  window. Details may also be in logs\startup_error.log
  echo  Run CheckEnv.bat to check this computer.
  echo ============================================================
  pause
)

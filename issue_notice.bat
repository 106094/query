@echo off & setlocal
set batchPath=%~dp0
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%issue_notice.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%weblinks_download.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%ITU-T_Report_mail.ps1
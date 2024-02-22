@echo off & setlocal
set batchPath=%~dp0
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%filename_len_check.ps1"
rem PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%leave_notice.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%MultiModulesWarning.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%return_notice.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%return_notice_w.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%ticket_notice.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%50diskcheck.ps1"
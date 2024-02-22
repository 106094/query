@echo off & setlocal
set batchPath=%~dp0
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%MultiModulesWarning.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%return_notice.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%ticket_notice.ps1"
@echo off & setlocal
set batchPath=%~dp0
rem PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%job_notice.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%drv_eva_sheets.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%backup.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%goldenreport_getinfo.ps1"
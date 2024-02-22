@echo off
set batchPath=%~dp0
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%auto_check_verify_drv.ps1"

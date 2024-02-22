@echo off
set batchPath=%~dp0
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%auto_RDVD_download.ps1"
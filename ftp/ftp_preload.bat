@echo off & setlocal
set batchPath=%~dp0
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%ftp_preload.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%dmicode_collect.ps1"
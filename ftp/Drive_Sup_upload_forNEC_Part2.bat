@echo off & setlocal
set batchPath=%~dp0
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%Drive_Sup_upload_forNEC_Part2.ps1"
@echo off & setlocal
set batchPath=%~dp0
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%rls_mails2.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%rls_mails.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%goemon_moving.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%release_note_query.ps1"

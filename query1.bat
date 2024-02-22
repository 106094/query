@echo off & setlocal
set batchPath=%~dp0
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%WB_moving.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%rls_mails.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%rls_mails2.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%goemon_moving.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%release_note_query.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%BIOSquery.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%zinfo_query.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%type2_query.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%get_eventID.ps1"
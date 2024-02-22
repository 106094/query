@echo off & setlocal
set batchPath=%~dp0
net time \\192.168.20.20 /set /yes
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%rls_mails.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%rls_mails2.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%borrow_mail_list.ps1
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%goemon_moving.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%release_note_query.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%FWquery.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%BIOSquery.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%zinfo_query.ps1"
rem PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%module_query.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%module_query_APP.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%module_query_UET.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%tool_req.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%RDVD_query.ps1"
PowerShell.exe -ExecutionPolicy Bypass -File "%batchPath%drv_tkts_query.ps1"
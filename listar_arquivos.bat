@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0listar_arquivos.ps1" %*
pause

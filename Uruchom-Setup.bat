@echo off
:: Uruchom skrypt PowerShell jako Administrator z pominięciem ExecutionPolicy
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Setup-Firmowy.ps1"
pause
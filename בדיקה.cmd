@echo off
rem Sadar - check
rem Hebrew output is produced by the PowerShell script, not here:
rem cmd.exe cannot be relied on to render it correctly.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build\launcher.ps1" -Action check

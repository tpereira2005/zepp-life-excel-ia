@echo off
rem Regenera o Excel e o JSON a partir da exportacao local, sem contactar a Xiaomi/Zepp.
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\atualizar-zepp.ps1" -SoExcel %*

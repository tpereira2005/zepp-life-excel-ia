@echo off
rem Descarrega as novas pesagens da Zepp Life e atualiza o Excel e o JSON para IA.
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\atualizar-zepp.ps1" %*

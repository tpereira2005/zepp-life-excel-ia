@echo off
rem Apaga a conta e a sessao Xiaomi guardadas neste computador (os dados ficam).
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\atualizar-zepp.ps1" -EsquecerConta -SemPausa
pause

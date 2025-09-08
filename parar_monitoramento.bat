@echo off
taskkill /f /im cmd.exe /fi "windowtitle eq INICIANDO MONITORAMENTO AUTOMÁTICO" >nul 2>&1
taskkill /f /im cmd.exe /fi "windowtitle eq ALFA SHARKS FLOW - AUTOMATIZADOR" >nul 2>&1
echo.
echo 🛑 Monitoramento automático parado
echo.
pause
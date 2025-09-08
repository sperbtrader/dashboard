@echo off
chcp 65001 >nul
echo.
echo 🚀 INICIANDO MONITORAMENTO AUTOMÁTICO
echo ====================================
echo.
echo 🤖 Monitorando planilha para atualizações automáticas
echo 📊 Pasta: C:\Users\Sperb Trader\Desktop\Alfa Sharks Flow
echo 🌐 Dashboard: https://sperbtrader-dashboard.streamlit.app/
echo ⏰ Verificando a cada 30 segundos
echo.
echo 💡 Mantenha esta janela aberta para monitoramento contínuo
echo 🛑 Feche para parar o monitoramento
echo.

:: Iniciar monitoramento em loop
:loop
call automatizador.bat
goto loop
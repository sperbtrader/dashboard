@echo off
chcp 65001 >nul
echo.
echo 🤖 ALFA SHARKS FLOW - AUTOMATIZADOR
echo ==================================
echo.

:: Navegar para a pasta correta
cd /d "C:\Users\Sperb Trader\Desktop\Alfa Sharks Flow"

:: Verificar se a planilha foi modificada
for %%F in ("planilhafluxo_PROCESSADO_COMPLETO.xlsx") do (
    set "filetime=%%~tF"
)

:: Esperar 5 segundos e verificar novamente
timeout /t 5 /nobreak >nul

for %%F in ("planilhafluxo_PROCESSADO_COMPLETO.xlsx") do (
    if not "%%~tF"=="%filetime%" (
        echo 📊 Planilha modificada! Atualizando dashboard...
        
        :: Fechar Excel se estiver aberto
        taskkill /f /im excel.exe >nul 2>&1
        
        :: Atualizar no GitHub
        git add planilhafluxo_PROCESSADO_COMPLETO.xlsx
        git commit -m "Atualização automática: %date% %time%"
        git push origin main
        
        echo ✅ Dashboard atualizado no Streamlit Cloud!
        echo 🌐 Acesse: https://sperbtrader-dashboard.streamlit.app/
        echo.
    ) else (
        echo ⏳ Planilha não modificada. Verificando novamente em 30 segundos...
    )
)

:: Verificar a cada 30 segundos
timeout /t 30 /nobreak >nul

:: Reiniciar o monitoramento
start /min "" "%~f0"
exit
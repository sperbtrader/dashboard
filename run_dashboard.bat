@echo off
chcp 65001 >nul
echo.
echo 📊 ALFA SHARKS FLOW DASHBOARD
echo =============================
echo.

:: Fechar Excel se estiver aberto
echo 🚀 Fechando Excel...
taskkill /f /im excel.exe >nul 2>&1
timeout /t 2 /nobreak >nul

:: Navegar para a pasta correta
echo 📂 Navegando para a pasta...
cd /d "C:\Users\Sperb Trader\Desktop\Alfa Sharks Flow"

:: Verificar se o arquivo existe
if not exist "dashboard.py" (
    echo ❌ ERRO: Arquivo dashboard.py não encontrado!
    echo 📍 Verifique se está na pasta correta
    pause
    exit /b 1
)

:: Verificar se Python está instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ ERRO: Python não encontrado!
    echo 📦 Instale o Python em: https://python.org
    pause
    exit /b 1
)

:: Verificar se Streamlit está instalado
python -c "import streamlit" 2>nul
if errorlevel 1 (
    echo 📦 Instalando Streamlit...
    pip install streamlit
)

:: Verificar outras dependências
echo 🔄 Verificando dependências...
pip install pandas plotly openpyxl

:: Executar o dashboard
echo.
echo 🚀 Iniciando Alfa Sharks Flow Dashboard...
echo 📋 URL: http://localhost:8501
echo ⏳ Aguarde o carregamento...
echo.
echo 💡 DICA: Mantenha esta janela aberta enquanto usar o dashboard
echo.

timeout /t 3 /nobreak >nul

:: Executar o Streamlit
streamlit run dashboard.py

echo.
echo ✅ Dashboard finalizado!
echo.
pause
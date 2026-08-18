@echo off
cd /d "%~dp0"
echo ========================================
echo   DataHub v2 - Gerando Executavel
echo ========================================
echo.

echo Instalando PyInstaller...
python -m pip install pyinstaller

echo.
echo Limpando build anterior (dist)...
if exist dist rmdir /s /q dist
if exist DataHub.spec del /q DataHub.spec

echo.
echo Descobrindo diretorio do Selenium...
for /f "tokens=*" %%i in ('python -c "import selenium; import os; print(os.path.dirname(selenium.__file__))"') do set SELENIUM_DIR=%%i
echo Selenium em: %SELENIUM_DIR%

echo.
echo Gerando executavel...
python -m PyInstaller --noconfirm ^
    --name DataHub ^
    --onedir ^
    --windowed ^
    --noupx ^
    --icon "appicon.ico" ^
    --add-data "config.py;." ^
    --add-data "%SELENIUM_DIR%;selenium" ^
    --add-data "core/processors/dashboard_template.html;core/processors" ^
    --add-data "version.txt;." ^
    --add-data "appicon.ico;." ^
    --hidden-import=pandas ^
    --hidden-import=openpyxl ^
    --hidden-import=selenium ^
    --hidden-import=webdriver_manager ^
    --hidden-import=customtkinter ^
    --hidden-import=numpy ^
    --hidden-import=core.processors.incidencias ^
    --hidden-import=core.processors.utilidades ^
    --hidden-import=core.base_mensal ^
    --hidden-import=core.downloader.geonline ^
    --hidden-import=core.downloader.operview ^
    --hidden-import=core.scheduler ^
    --hidden-import=integrations.sharepoint ^
    --hidden-import=integrations.n8n ^
    --hidden-import=ui.main_window ^
    --hidden-import=ui.config_window ^
    --hidden-import=selenium.webdriver.chrome.options ^
    --hidden-import=selenium.webdriver.common.by ^
    --hidden-import=selenium.webdriver.support.ui ^
    --hidden-import=selenium.webdriver.support.expected_conditions ^
    --hidden-import=selenium.webdriver.common.action_chains ^
    --hidden-import=selenium.webdriver.common.keys ^
    --hidden-import=selenium.webdriver.chrome.service ^
    --hidden-import=selenium.webdriver.chrome.webdriver ^
    --hidden-import=selenium.webdriver.remote.webdriver ^
    --hidden-import=selenium.webdriver.remote.webelement ^
    --hidden-import=selenium.common.exceptions ^
    --hidden-import=core.updater ^
    app.py

echo.
echo ========================================
echo   Build concluido!
echo   Pasta: dist\DataHub\
echo   Executavel: dist\DataHub\DataHub.exe
echo ========================================
echo.
echo Para distribuir / publicar atualizacao:
echo 1. Compacte a PASTA dist\DataHub (nao o conteudo de dentro) em DataHub.zip
echo 2. O usuario precisa ter o Google Chrome instalado
echo 3. Extraia e rode DataHub.exe - ou publique o DataHub.zip num Release do GitHub
echo ========================================

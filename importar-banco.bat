@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   Quote Follow Up - planilhas para o MariaDB
echo ============================================
where node >nul 2>nul
if errorlevel 1 ( echo. & echo [ERRO] Node.js nao encontrado. Instale em https://nodejs.org ^(LTS^). & pause & exit /b )
if not exist node_modules ( echo. & echo [ERRO] Rode o build.bat uma vez antes ^(ele instala as bibliotecas^). & pause & exit /b )
if not exist "Quote Follow Up data.xlsx" ( echo. & echo [ERRO] Copie "Quote Follow Up data.xlsx" para esta pasta. & pause & exit /b )
if not exist "Quote Follow Up reports.xlsx" ( echo. & echo [ERRO] Copie "Quote Follow Up reports.xlsx" para esta pasta. & pause & exit /b )
echo.
node import-xlsx.mjs "Quote Follow Up data.xlsx" "Quote Follow Up reports.xlsx"
if errorlevel 1 ( echo. & echo [ERRO] Falha na leitura ^(veja a mensagem acima^). & pause & exit /b )
echo.
echo Agora carregue no banco ^(pede a senha do MariaDB^):
echo   mariadb -u root -p ^< import.sql
echo.
echo O import.sql tem os dados da empresa e os codigos das senhas: nao mande para o GitHub.
echo.
pause

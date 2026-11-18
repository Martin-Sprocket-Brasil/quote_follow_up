@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   Quote Follow Up - compilador
echo ============================================
where node >nul 2>nul
if errorlevel 1 (
  echo.
  echo [ERRO] Node.js nao encontrado.
  echo Instale o Node.js em https://nodejs.org  ^(botao LTS^) e rode este arquivo de novo.
  echo.
  pause
  exit /b
)
set FALTA=
if not exist node_modules\express set FALTA=1
if not exist node_modules\pdf-lib set FALTA=1
if defined FALTA (
  echo.
  echo Instalando as bibliotecas ^(pode demorar 1-2 min^)...
  call npm install
  if errorlevel 1 ( echo. & echo [ERRO] Falha ao instalar. Verifique sua internet. & pause & exit /b )
)
echo.
echo Compilando...
node build.js
if errorlevel 1 ( echo. & echo [ERRO] Falha na compilacao ^(veja a mensagem acima^). & pause & exit /b )
echo.
echo Tudo certo! O arquivo QuoteFollowUp.html foi gerado nesta pasta.
echo Abra ele no Chrome/Edge.
echo.
pause

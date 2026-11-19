@echo off
chcp 65001 >nul
cd /d "%~dp0"
rem ============================================================
rem  Quote Follow Up - servidor (API + pagina)
rem  Dois cliques: liga o servidor nesta janela (fechar a janela = desligar).
rem  Como servico (tarefa agendada): iniciar-servidor.bat /servico
rem    -> sem janela, grava em servidor\servidor.log e religa sozinho se cair.
rem ============================================================
where node >nul 2>nul
if errorlevel 1 (
  echo [ERRO] Node.js nao encontrado. Instale o Node.js LTS em https://nodejs.org
  if not "%1"=="/servico" pause
  exit /b 1
)
set FALTA=
if not exist node_modules\express set FALTA=1
if not exist node_modules\pdf-lib set FALTA=1
if defined FALTA (
  echo Instalando as bibliotecas do servidor ^(so na primeira vez, 1-2 min^)...
  call npm install
  if errorlevel 1 (
    echo [ERRO] Falha ao instalar. Verifique a internet.
    if not "%1"=="/servico" pause
    exit /b 1
  )
)
if not exist pagina\QuoteFollowUp.html (
  echo Compilando a pagina...
  node codigo\build.js
)
rem versoes antigas guardavam a configuracao na pasta do programa: passa para a pasta servidor
if not exist servidor\server.config.json if exist server.config.json (
  move server.config.json servidor\server.config.json >nul
  echo Movi o server.config.json para a pasta servidor.
)
if not exist servidor\server.config.json (
  copy servidor\server.config.example.json servidor\server.config.json >nul
  echo.
  echo Criei o arquivo servidor\server.config.json. Coloque nele o usuario e a senha do banco
  echo ^(parte "db"^), salve e rode este arquivo de novo.
  echo.
  notepad servidor\server.config.json
  if not "%1"=="/servico" pause
  exit /b 1
)
if "%1"=="/servico" goto servico
node servidor\server.mjs
echo.
echo O servidor parou ^(veja a mensagem acima^).
pause
exit /b

:servico
echo %date% %time% - iniciando >> servidor\servidor.log
node servidor\server.mjs >> servidor\servidor.log 2>&1
echo %date% %time% - o servidor parou; religando em 10 s >> servidor\servidor.log
timeout /t 10 /nobreak >nul
goto servico

@echo off
chcp 65001 >nul
cd /d "%~dp0"
rem ============================================================
rem  Quote Follow Up - servidor (API + pagina)
rem  Dois cliques: liga o servidor nesta janela (fechar a janela = desligar).
rem  Como servico (tarefa agendada): iniciar-servidor.bat /servico
rem    -> sem janela, grava em servidor.log e religa sozinho se cair.
rem ============================================================
where node >nul 2>nul
if errorlevel 1 (
  echo [ERRO] Node.js nao encontrado. Instale o Node.js LTS em https://nodejs.org
  if not "%1"=="/servico" pause
  exit /b 1
)
if not exist node_modules\express (
  echo Instalando as bibliotecas do servidor ^(so na primeira vez, 1-2 min^)...
  call npm install
  if errorlevel 1 (
    echo [ERRO] Falha ao instalar. Verifique a internet.
    if not "%1"=="/servico" pause
    exit /b 1
  )
)
if not exist QuoteFollowUp.html (
  echo Compilando a pagina...
  node build.js
)
if not exist server.config.json (
  copy server.config.example.json server.config.json >nul
  echo.
  echo Criei o arquivo server.config.json. Coloque nele o usuario e a senha do banco
  echo ^(parte "db"^), salve e rode este arquivo de novo.
  echo.
  notepad server.config.json
  if not "%1"=="/servico" pause
  exit /b 1
)
if "%1"=="/servico" goto servico
node server.mjs
echo.
echo O servidor parou ^(veja a mensagem acima^).
pause
exit /b

:servico
echo %date% %time% - iniciando >> servidor.log
node server.mjs >> servidor.log 2>&1
echo %date% %time% - o servidor parou; religando em 10 s >> servidor.log
timeout /t 10 /nobreak >nul
goto servico

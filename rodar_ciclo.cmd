@echo off
setlocal
chcp 65001 >nul
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
if not exist "resultados" mkdir "resultados"

set TENTATIVA=0
:loop
set /a TENTATIVA+=1
echo ===== %date% %time%  (tentativa %TENTATIVA%) =====>> "resultados\ultimo_ciclo.log"
"C:\Users\guilherme-medina\AppData\Local\Programs\Python\Python313\python.exe" -u "orquestrador_doe.py" --vigia >> "resultados\ultimo_ciclo.log" 2>&1
if %ERRORLEVEL% EQU 0 exit /b 0
if %ERRORLEVEL% NEQ 2 exit /b %ERRORLEVEL%
if %TENTATIVA% GEQ 4 exit /b %ERRORLEVEL%
echo [%date% %time%] API falhou (codigo 2). Nova execucao em 180s...>> "resultados\ultimo_ciclo.log"
timeout /t 180 /nobreak >nul
goto loop

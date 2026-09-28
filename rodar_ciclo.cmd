@echo off
setlocal
chcp 65001 >nul
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
if not exist "resultados" mkdir "resultados"
echo ===== %date% %time% =====>> "resultados\ultimo_ciclo.log"
"C:\Users\guilherme-medina\AppData\Local\Programs\Python\Python313\python.exe" -u "orquestrador_doe.py" >> "resultados\ultimo_ciclo.log" 2>&1
exit /b %ERRORLEVEL%

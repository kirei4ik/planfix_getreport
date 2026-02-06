@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Запускаем выгрузку из Planfix...
powershell -ExecutionPolicy Bypass -File "%~dp0planfix_getreport.ps1"
echo.
echo Готово! Нажми любую клавишу для закрытия...
pause >nul
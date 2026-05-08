@echo off
chcp 65001 > nul
REM Непрерывный мониторинг папки с .msg-письмами в headless-режиме.
REM Двойной клик = запуск; Ctrl+C для остановки.

cd /d "%~dp0"
"%~dp0asud.exe" --headless --mode=email --watch

pause

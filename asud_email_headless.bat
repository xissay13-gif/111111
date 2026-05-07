@echo off
chcp 65001 > nul
REM Запуск asud.exe в headless-режиме сразу в email-режиме (создание из .msg).
REM Двойной клик = запуск; путь к папке с письмами спрашивается интерактивно.

cd /d "%~dp0"
"%~dp0asud.exe" --headless --mode=email

pause

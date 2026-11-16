@echo off
title AssignmentFixer - Khoi chay Ung dung
cd /d %~dp0
echo ========================================================
echo       ASSIGNMENTFIXER — NEU ACADEMIC EDITOR
echo ========================================================
echo Dang khoi chay giao dien Desktop...
python -m app.main
if %errorlevel% neq 0 (
    echo.
    echo [LOI] Khong the khoi chay ung dung!
    pause
)

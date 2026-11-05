@echo off
title AI Resume Analyzer

echo ==========================================
echo       AI RESUME ANALYZER
echo ==========================================
echo.
echo Starting Flask Backend...
echo.

cd /d "C:\AI Resume Analyzer"

start "Flask Backend" cmd /k ".venv\Scripts\python.exe backend\app.py"

timeout /t 3 /nobreak >nul

echo.
echo Flask Backend Started!
echo.
echo Open your frontend using Live Server.
echo ==========================================
pause
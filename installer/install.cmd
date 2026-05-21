@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_app.ps1"
exit /b %errorlevel%

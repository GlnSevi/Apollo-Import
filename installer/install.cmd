@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_app.ps1" -SourceExe "Apollo-Import-GUI-v0.1-onefile.exe"
exit /b %errorlevel%

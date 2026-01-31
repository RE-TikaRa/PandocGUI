@echo off
setlocal
taskkill /IM PandocGUI.exe /F >nul 2>&1
if exist bin rmdir /s /q bin
if exist obj rmdir /s /q obj
dotnet build -p:Platform=x64
endlocal

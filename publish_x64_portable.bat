@echo off
setlocal
taskkill /IM PandocGUI.exe /F >nul 2>&1
if exist bin rmdir /s /q bin
if exist obj rmdir /s /q obj
dotnet publish -c Release -p:Platform=x64 -p:PublishProfile=Properties\PublishProfiles\win-x64.pubxml
endlocal

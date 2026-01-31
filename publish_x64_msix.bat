@echo off
setlocal
taskkill /IM PandocGUI.exe /F >nul 2>&1
if exist bin rmdir /s /q bin
if exist obj rmdir /s /q obj
dotnet publish -c Release -p:Platform=x64 -p:RuntimeIdentifier=win-x64 -p:WindowsPackageType=MSIX -p:AppxPackageSigningEnabled=false -p:AppxBundle=Never -p:GenerateAppxPackageOnBuild=true
endlocal

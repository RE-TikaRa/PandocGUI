@echo off
setlocal
taskkill /IM PandocGUI.exe /F >nul 2>&1
if exist bin rmdir /s /q bin
if exist obj rmdir /s /q obj
set MSPDBCMF_PATH=C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Microsoft\VisualStudio\v17.0\AppxPackage\x86\MsPdbCmf.exe
dotnet publish -c Release -p:Platform=x64 -p:RuntimeIdentifier=win-x64 -p:WindowsPackageType=MSIX -p:AppxPackageSigningEnabled=false -p:AppxBundle=Never -p:GenerateAppxPackageOnBuild=true -p:MsPdbCmfExeFullpath="%MSPDBCMF_PATH%"
endlocal

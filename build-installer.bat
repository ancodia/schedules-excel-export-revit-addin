@echo off
setlocal

set PROJECT=src\SchedulesExcelExport.csproj
set VERSIONS=2020 2021 2022 2023 2024 2025

echo Building all Revit versions...
echo.

for %%V in (%VERSIONS%) do (
    echo Building Release-%%V...
    msbuild %PROJECT% /t:Restore /p:Configuration=Release-%%V /v:minimal /nologo
    if errorlevel 1 goto :error
    
    msbuild %PROJECT% /t:Build /p:Configuration=Release-%%V /v:minimal /nologo
    if errorlevel 1 goto :error
    echo.
)

echo.
echo Building installer...
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" "installer\script.iss"
if errorlevel 1 goto :error

echo.
echo SUCCESS! Installer at: src\output\SchedulesExcelExport.Revit.Addin.Installer.exe
goto :end

:error
echo.
echo BUILD FAILED!
exit /b 1

:end
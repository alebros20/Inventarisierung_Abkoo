@echo off
echo ================================
echo  Inventarisierung - Build + Publish
echo ================================
echo.

:: ==============================
::  VERSION HIER ANPASSEN
:: ==============================
set VERSION=1.0.0.0

:: ==============================
set MSBUILD="C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe"
set PROJECT="C:\Users\CC-Student\source\repos\WindowsFormsApp1\WindowsFormsApp1\Inventarisierung.csproj"
set INNO="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
set ISS="C:\Users\CC-Student\source\repos\WindowsFormsApp1\installer.iss"
set OUTPUT="C:\Users\CC-Student\Desktop\Installer"
set ASSEMBLYINFO="C:\Users\CC-Student\source\repos\WindowsFormsApp1\WindowsFormsApp1\Properties\AssemblyInfo.cs"

echo Version: %VERSION%
echo.

echo [1/3] Version in AssemblyInfo.cs setzen...
powershell -Command "(Get-Content %ASSEMBLYINFO%) -replace 'AssemblyVersion\(\".*\"\)', 'AssemblyVersion(\"%VERSION%\")' -replace 'AssemblyFileVersion\(\".*\"\)', 'AssemblyFileVersion(\"%VERSION%\")' | Set-Content %ASSEMBLYINFO%"
if %ERRORLEVEL% NEQ 0 (
    echo FEHLER beim Setzen der Version!
    pause
    exit /b 1
)
echo Version gesetzt!
echo.

echo [2/3] Kompilieren...
%MSBUILD% %PROJECT% /t:Build /p:Configuration=Release /v:minimal
if %ERRORLEVEL% NEQ 0 (
    echo FEHLER beim Kompilieren!
    pause
    exit /b 1
)
echo Kompilieren erfolgreich!
echo.

echo [3/3] Installer erstellen...
%INNO% %ISS% /DAppVersion=%VERSION%
if %ERRORLEVEL% NEQ 0 (
    echo FEHLER beim Erstellen des Installers!
    pause
    exit /b 1
)
echo.
echo ================================
echo  Fertig!
echo  Version:   %VERSION%
echo  Installer: %OUTPUT%\Inventarisierung_Setup.exe
echo ================================
pause

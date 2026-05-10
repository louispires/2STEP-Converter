@echo off
chcp 65001 > nul
setlocal EnableDelayedExpansion
title 2STEP-Converter
cd /d "%~dp0"


if exist "%~dp0lib" (
    set _MM_ROOT=%~dp0lib
) else if exist "%LOCALAPPDATA%\STLtoSTP" (
    set _MM_ROOT=%LOCALAPPDATA%\STLtoSTP
) else (
    echo No existing environment found. Where should the environment be installed?
    echo.
    echo  [1] Next to this script  ^(portable^)
    echo  [2] %LOCALAPPDATA%\STLtoSTP
    echo.
    choice /c 12 /n /m "Your choice: "
    if errorlevel 2 (
        set _MM_ROOT=%LOCALAPPDATA%\STLtoSTP
    ) else (
        set _MM_ROOT=%~dp0lib
    )
    echo.
)

set _LONGPATH=
for /f "tokens=3" %%V in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\FileSystem" /v LongPathsEnabled 2^>nul') do set _LONGPATH=%%V
if not "!_LONGPATH!"=="0x1" (
    echo [WARNING] Windows 260-character path limit is not enabled.
    echo.
    echo  [1] Enable long paths automatically and reboot  ^(recommended^)
    echo  [2] Install environment to %LOCALAPPDATA%\STLtoSTP  ^(no reboot needed^)
    echo.
    choice /c 12 /n /m "Your choice: "
    if errorlevel 2 (
        set _MM_ROOT=%LOCALAPPDATA%\STLtoSTP
        echo Using: %LOCALAPPDATA%\STLtoSTP
        echo.
    ) else (
        echo Requesting administrator access to enable long paths...
        powershell -Command "Start-Process cmd -ArgumentList '/c reg add HKLM\SYSTEM\CurrentControlSet\Control\FileSystem /v LongPathsEnabled /t REG_DWORD /d 1 /f' -Verb RunAs -Wait"
        echo.
        echo Long paths enabled. The PC will reboot in 10 seconds.
        echo Run this script again after reboot.
        shutdown /r /t 10 /c "Enabling Windows long path support for 2STEP"
        pause & exit /b 0
    )
)

set _MM=!_MM_ROOT!\micromamba.exe
set _ENV=!_MM_ROOT!\env
set _PY=!_ENV!\python.exe
set MAMBA_ROOT_PREFIX=!_MM_ROOT!
set CONDA_PKGS_DIRS=!_MM_ROOT!
set PYTHONNOUSERSITE=1

set PATH=!_ENV!\Library\bin;!_ENV!\Library\mingw-w64\bin;!_ENV!\Scripts;!_ENV!;%PATH%

if exist "!_MM!" goto :check_env

if not exist "!_MM_ROOT!" mkdir "!_MM_ROOT!"
echo Downloading portable Python manager (one-time, ~10 MB) ...
curl.exe --ssl-no-revoke -L --progress-bar -o "!_MM!" "https://github.com/mamba-org/micromamba-releases/releases/download/2.6.0-0/micromamba-win-64.exe"
if errorlevel 1 (
    echo [ERROR] Download failed. Check your internet connection.
    pause & exit /b 1
)
for %%F in ("!_MM!") do set _SZ=%%~zF
if !_SZ! LSS 5000000 (
    echo [ERROR] Download corrupt ^(!_SZ! bytes^). Delete micromamba.exe and retry.
    del /f /q "!_MM!" 2>nul
    pause & exit /b 1
)

:check_env
if exist "!_PY!" goto :check_occ

echo Setting up Python environment (one-time download, ~500 MB) ...
"!_MM!" create --prefix "!_ENV!" -c conda-forge python=3.12 pythonocc-core --yes
if errorlevel 1 (
    echo [ERROR] Failed to create Python environment.
    pause & exit /b 1
)
goto :run

:check_occ
"!_PY!" -c "from OCC.Core.StlAPI import StlAPI_Reader" >nul 2>&1
if errorlevel 1 (
    echo OpenCASCADE not found or broken -- reinstalling ...
    "!_MM!" install --prefix "!_ENV!" -c conda-forge pythonocc-core --yes
    if errorlevel 1 (
        echo [ERROR] Failed to install pythonocc-core.
        pause & exit /b 1
    )
)


:run
"!_PY!" "%~dp02STEP-Converter.py" %*

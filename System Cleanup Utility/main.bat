@echo off
setlocal EnableDelayedExpansion
title System Cleanup Utility

:: Set console colors
color 0A

:: Display header
echo.
echo ================================================================================
echo                           SYSTEM CLEANUP UTILITY                               
echo ================================================================================
echo.

:: Check for administrator privileges
net session >nul 2>&1
if %errorLevel% neq 0 (
    color 0C
    echo [ERROR] This script requires administrator privileges!
    echo.
    echo Attempting to restart with admin rights...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:: User choice menu
:menu
echo Please select an option:
echo.
echo   [1] Clean and Shutdown
echo   [2] Clean and Restart
echo.
set /p choice="Enter your choice (1 or 2): "

if "%choice%"=="1" (
    set "action=shutdown"
    goto :startCleanup
) else if "%choice%"=="2" (
    set "action=restart"
    goto :startCleanup
) else (
    echo.
    color 0C
    echo Invalid choice! Please enter 1 or 2.
    color 0A
    echo.
    goto :menu
)

:startCleanup
echo.
echo Starting cleanup process...
echo.

:: Initialize progress counter
set /a step=0
set /a totalSteps=7

:: Function to display progress
goto :main

:showProgress
set /a step+=1
echo.
echo [%step%/%totalSteps%] %~1
echo --------------------------------------------------------------------------------
goto :eof

:main
:: Step 1: Clear Windows Temp Files
call :showProgress "Cleaning temporary files..."
echo [*] Removing user temp files...
del /s /f /q "%TEMP%\*.*" >nul 2>&1
for /d %%x in ("%TEMP%\*") do @rd /s /q "%%x" >nul 2>&1

echo [*] Removing system temp files...
del /s /f /q "%WINDIR%\Temp\*.*" >nul 2>&1
for /d %%x in ("%WINDIR%\Temp\*") do @rd /s /q "%%x" >nul 2>&1

:: Step 2: Clear Prefetch Files
call :showProgress "Clearing Windows Prefetch data..."
del /s /f /q "%WINDIR%\Prefetch\*.*" >nul 2>&1

:: Step 3: Clear Thumbnail Cache
call :showProgress "Removing thumbnail cache..."
del /s /f /q "%LOCALAPPDATA%\Microsoft\Windows\Explorer\thumbcache_*.*" >nul 2>&1
del /s /f /q "%LOCALAPPDATA%\Microsoft\Windows\Explorer\iconcache_*.*" >nul 2>&1

:: Step 4: Clear Windows Update Cache
call :showProgress "Cleaning Windows Update cache..."
echo [*] Stopping Windows Update service...
net stop wuauserv >nul 2>&1
del /s /f /q "%WINDIR%\SoftwareDistribution\Download\*.*" >nul 2>&1
for /d %%x in ("%WINDIR%\SoftwareDistribution\Download\*") do @rd /s /q "%%x" >nul 2>&1
echo [*] Starting Windows Update service...
net start wuauserv >nul 2>&1

:: Step 5: Clear DNS Cache
call :showProgress "Flushing DNS cache..."
ipconfig /flushdns >nul 2>&1

:: Step 6: Empty Recycle Bin
call :showProgress "Emptying Recycle Bin..."
PowerShell.exe -NoProfile -ExecutionPolicy Bypass -Command "Clear-RecycleBin -Force -ErrorAction SilentlyContinue" >nul 2>&1

:: Step 7: Additional cleanup
call :showProgress "Performing final cleanup..."
:: Clear Windows error reporting files
del /s /f /q "%LOCALAPPDATA%\CrashDumps\*.*" >nul 2>&1
del /s /f /q "%PROGRAMDATA%\Microsoft\Windows\WER\ReportQueue\*.*" >nul 2>&1
del /s /f /q "%PROGRAMDATA%\Microsoft\Windows\WER\ReportArchive\*.*" >nul 2>&1

:: Clear Windows log files
del /s /f /q "%WINDIR%\Logs\CBS\*.log" >nul 2>&1
del /s /f /q "%WINDIR%\Logs\DISM\*.log" >nul 2>&1

:: Display summary
echo.
echo ================================================================================
color 0E
echo                              CLEANUP COMPLETE                                   
echo ================================================================================
echo.
echo [OK] Temporary files cleaned
echo [OK] Prefetch data cleared
echo [OK] Thumbnail cache removed
echo [OK] Windows Update cache cleaned
echo [OK] DNS cache flushed
echo [OK] Recycle Bin emptied
echo [OK] System optimized
echo.

:: Action based on user choice
if "%action%"=="shutdown" (
    echo ================================================================================
    echo                         AUTOMATIC SHUTDOWN IN PROGRESS                          
    echo ================================================================================
    echo.
    color 0C
    echo System will shutdown automatically in:
    for /l %%i in (10,-1,1) do (
        <nul set /p "=[%%i] "
        timeout /t 1 /nobreak >nul
    )
    echo.
    echo.
    color 0A
    echo [*] Initiating shutdown...
    shutdown /s /f /t 0
) else (
    echo ================================================================================
    echo                         AUTOMATIC RESTART IN PROGRESS                           
    echo ================================================================================
    echo.
    color 0C
    echo System will restart automatically in:
    for /l %%i in (10,-1,1) do (
        <nul set /p "=[%%i] "
        timeout /t 1 /nobreak >nul
    )
    echo.
    echo.
    color 0A
    echo [*] Initiating restart...
    shutdown /r /f /t 0
)
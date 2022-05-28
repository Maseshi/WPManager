@echo off
setlocal DisableDelayedExpansion

if not %verify% EQU 0 exit
title Install Office 365
cd ..

cls
call header.bat
echo This can take a long time as it downloads and installs all the applications available in Office 365 from Microsoft to your computer. which may incur internet charges The download and installation speed will depend on your hardware device.
echo.
echo  Required:
echo  - Network
echo.
choice /n /c YN /m "Are you sure you want to continue? [Y/n]"
if errorlevel 2 (
    cd .. & call "Windows Product Manager.bat"
    exit /b 0
)

cls
call header.bat
echo Checking your device information...
if "%PROCESSOR_ARCHITECTURE%" == "x86" (
    set bit=x86
) else if "%PROCESSOR_ARCHITECTURE%" == "AMD64" (
    set bit=x64
) else (
    echo %ESC%[93mYour device is not supported.%ESC%[0m
    echo.
    pause
    cd .. & call "Windows Product Manager.bat"
)
echo Calling the Office 365 installer...
cd office365
setup.exe /configure configuration-Office365-%bit%.xml && (
    echo Successfully installed
) || (
    echo Failed to install
)
echo.
pause
cd .. & call "Windows Product Manager.bat"

endlocal
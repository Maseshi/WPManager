@echo off
setlocal EnableDelayedExpansion

cd data & pushd %~dp0 & pushd data
call init.bat

:main
title %n%
cls
call header.bat
echo %bgcoffee%
echo.
echo %ESC%[93mIf you are worried about computer virus within this instruction set. You don't need to turn off any of your antivirus tools, we recommend leaving it as is. Although this instruction set does not contain any viruses. We will not modify system files or steal your personal information.%ESC%[0m
echo.
echo %ESC%[1m%n%%ESC%[0m
echo Within this tool you can manage the list of options below.
echo.
if not "%1"=="am_admin" (
    echo [1] Microsoft Office 365
    echo [2] Microsoft Windows
    echo [3] Run as administrator
    echo [Q] Exit the program
    echo.
    choice /n /c 123Q /m "Choose your topic to continue. [1, 2, 3, Q] "
    if errorlevel 4 exit
    if errorlevel 3 powershell start -verb runas '%0' am_admin & exit /b
    if errorlevel 2 goto :windows
    if errorlevel 1 goto :office
) else (
    echo [1] Microsoft Office 365
    echo [2] Microsoft Windows
    echo [Q] Exit the program
    echo.
    choice /n /c 12Q /m "Choose your topic to continue. [1, 2, Q] "
    if errorlevel 3 exit
    if errorlevel 2 goto :windows
    if errorlevel 1 goto :office
)
goto :main
pause & exit /b 0

:office
title Microsoft Office 365
cd data
cls
call header.bat
echo Microsoft Office 365
echo Install or activate Office 365 on your computer.
echo.
echo [1] Download and install Office 365.
echo [2] Activate Office 365
echo [B] Back to main page
echo [Q] Exit the program
echo.
choice /n /c 12BQ /m "Choose your topic to continue. [1, 2, B, Q] "
if errorlevel 4 exit
if errorlevel 3 goto :main
if errorlevel 2 cd prompt & call activate-office.bat
if errorlevel 1 cd prompt & call install-office.bat
goto :office
exit /b 0

@REM Comming soon...
@REM :windowsServer
@REM title Microsoft Windows Server
@REM cd data
@REM cls
@REM call header.bat
@REM echo %ESC%[1mMicrosoft Windows Server%ESC%[0m
@REM echo Manage your Windows Server or view your product information.
@REM echo.
@REM echo [1] Activate Windows Server
@REM echo [2] Explore the information of the product key.
@REM echo [3] Explore details about product key.
@REM echo [4] Explore the expiration date of the product key.
@REM echo [5] View Windows version
@REM echo [B] Back to main page
@REM echo [Q] Exit the program
@REM echo.
@REM choice /n /c 12345BQ /m "Choose your topic to continue. [1, 2, 3, 4, 5, B, Q] "
@REM if errorlevel 7 exit
@REM if errorlevel 6 goto :main
@REM if errorlevel 5 winver
@REM if errorlevel 4 wscript //nologo slmgr.vbs /xpr
@REM if errorlevel 3 wscript //nologo pdki.vbs
@REM if errorlevel 2 wscript //nologo slmgr.vbs /dli
@REM if errorlevel 1 cd prompt & call activate-win-server.bat
@REM goto :windowsServer
@REM exit /b 0

:windows
title Microsoft Windows
cd data
cls
call header.bat
echo %ESC%[1mMicrosoft Windows%ESC%[0m
echo Manage your Windows or view your product information.
echo.
echo [1] Activate Windows
echo [2] Explore the information of the product key.
echo [3] Explore details about product key.
echo [4] Explore the expiration date of the product key.
echo [5] View Windows version
echo [B] Back to main page
echo [Q] Exit the program
echo.
choice /n /c 12345BQ /m "Choose your topic to continue. [1, 2, 3, 4, 5, B, Q] "
if errorlevel 7 exit
if errorlevel 6 goto :main
if errorlevel 5 winver
if errorlevel 4 wscript //nologo slmgr.vbs /xpr
if errorlevel 3 wscript //nologo pdki.vbs
if errorlevel 2 wscript //nologo slmgr.vbs /dli
if errorlevel 1 cd prompt & call activate-windows.bat
goto :windows
exit /b 0

endlocal
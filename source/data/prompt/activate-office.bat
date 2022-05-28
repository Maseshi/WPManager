@echo off
setlocal DisableDelayedExpansion

if not %verify% EQU 0 exit
cd ..
goto :office

:office
title Activate Office 365

cls
call header.bat
echo Your Office 365 activation needs to be in admin mode. This will only change your Office 365 product key and will not modify any of your system files unless necessary.
echo.
echo  Required:
echo  - Administrator
echo  - Network
echo  Supported products:
echo  - Standard 2019
echo  - Professional Plus 2019
echo.
choice /n /c YN /m "Want to continue activating? [Y/n]"
if errorlevel 2 (
    cd .. & call "Windows Product Manager.bat"
    exit /b 0
)

cls
call header.bat
echo %ESC%[93m^<^?^> Running in the background:%ESC%[0m
echo Authenticating...
fsutil dirty query %systemdrive% >nul
if %errorlevel% EQU 0 (
    echo Testing internet connection...
    ping -n 1 -w 1000 www.google.com | findstr /c:"TTL=" >nul

    if %errorlevel% EQU 0 (
        echo Checking your Office 365 folders...
        if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" (
            cd /d "%ProgramFiles%\Microsoft Office\Office16"
        ) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" (
            cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
        ) else (
            set "err=echo You haven't installed Office 365. & echo.Please install Office 365 before doing this."
            goto error
        )

        echo Converting your retail license to volume one...
        for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do (
            cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
        )

        echo Uninstalling the current product key...
        cscript //nologo ospp.vbs /unpkey:6MWKP >nul

        echo Installing a new product key...
        cscript //nologo ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul
        goto :activate
    ) else (
        set "err=Please connect to the internet to fully activate your Office 365."
        goto error
    )
) else (
    set "err=Please open this file by %ESC%[1mRun as adminstrator."
    goto error
)
exit /b 0

:activate
echo Prepare for activation...
set cserver=16
set schoose=1

echo Checking %cserver% provider servers...
if %schoose% == 1 set server=kms.digiboy.ir
if %schoose% == 2 set server=hq1.chinancce.com
if %schoose% == 3 set server=54.223.212.31
if %schoose% == 4 set server=kms.cnlic.com
if %schoose% == 5 set server=kms.chinancce.com
if %schoose% == 6 set server=kms.ddns.net
if %schoose% == 7 set server=franklv.ddns.net
if %schoose% == 8 set server=k.zpale.com
if %schoose% == 9 set server=m.zpale.com
if %schoose% == 10 set server=mvg.zpale.com
if %schoose% == 11 set server=kms.shuax.com
if %schoose% == 12 set server=kensol263.imwork.net
if %schoose% == 13 set server=xykz.f3322.org
if %schoose% == 14 set server=kms789.com
if %schoose% == 15 set server=dimanyakms.sytes.net
if %schoose% == 16 set server=kms.03k.org
if %schoose% == 4 (
    set "err=Sorry, it seems unable to connect to any provider server at all. & echo.Please try again later."
    goto error
)

echo Setting the server port...
cscript //nologo ospp.vbs /setprt:1688 >nul

echo Connecting to the current code provider server...
cscript //nologo ospp.vbs /sethst:%server% >nul

echo Your Office 365 is being activated...
cscript //nologo ospp.vbs /act | find /i "successful" >nul && (
    echo %ESC%[92m^<^=^> Complete the process%ESC%[0m
    echo.
    pause

    cls
    call header.bat
    echo Now you are ready to go on.
    echo You have activated your Office 365. Please check if the Office 365 applications are activated. If it is still not activated try again or tell us about those errors. Have fun.
    echo How it works: bit.ly/understanding-kms
    echo.
    pause
    cd .. & call "Windows Product Manager.bat"
) || (
    echo .
    echo %ESC%[93mUnable to connect to the provider server. Trying to switch to another server...%ESC%[0m
    echo %ESC%[93mPlease wait...%ESC%[0m
    echo .
    set /a cserver-=1
    set /a schoose+=1
    goto :activate
)
exit /b 0



:error
echo %ESC%[91m^<^!^> Unable to continue:%ESC%[0m
echo %err%
echo.
pause
goto :office
exit /b 0

endlocal
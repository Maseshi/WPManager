@echo off
setlocal DisableDelayedExpansion

if not %verify% EQU 0 exit
cd ..

for /f "tokens=4-5 delims=. " %%i in ('ver') do set version=%%i.%%j
if "%version%" == "6.1" (
    set win=7
    goto :win7
) else if "%version%" == "6.2" (
    set win=8
    goto :win8
) else if "%version%" == "6.3" (
    set win=8.1
    goto :win8.1
) else if "%version%" == "10.0" (
    set win=10
    goto :win10
) else if "%version%" == "11.0" (
    set win=11
    goto :win11
) else (
    title Activate Windows

    cls
    call header.bat
    echo Unfortunately, this tool cannot be activated with your current Windows or your Windows may not be supported.
    echo.
    pause
    cd .. & call "Windows Product Manager.bat"
)

:win11
title Activate Windows 11

cls
call header.bat
echo The following actions must be in Administrator mode. It will only modify or change the product key within your Windows 11. This will not change any system files. It has nothing to do with your Windows 11 activation if you want to report any errors. You can contact the developer in a number of ways through the developer's website.
echo.
echo  Required:
echo  - Administrator
echo  - Network
echo  Supported products:
echo  - [Home, Home N, Home Single Language, Home Country Specific]
echo  - [Pro, Pro N, Pro for Workstations, Pro for Workstations N]
echo  - [Education, Education N]
echo  - [Enterprise, Enterprise N, Enterprise G, Enterprise G N,
echo     Enterprise LTSC 2019, Enterprise N LTSC 2019,
echo     Enterprise LTSB 2016, Enterprise N LTSB 2016,
echo     Enterprise 2015 LTSB, Enterprise 2015 LTSB N]
echo.
echo %ESC%[7mYou are running on Windows 11.%ESC%[0m
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
        echo Preparing to activate your Windows 11...
        for /f "delims=" %%a in ('cscript //nologo pdk.vbs') do (
            set oldKey=%%a
        )

        echo Configuring the original server...
        cscript //nologo slmgr.vbs /ckms >nul

        echo Uninstalling the current product key...
        cscript //nologo slmgr.vbs /upk >nul

        echo Clearing the old product code registration file...
        cscript //nologo slmgr.vbs /cpky >nul
            
        echo Searching for your edition of Windows 11...
        wmic os get name | findstr /c:"Home" >nul && (
            echo Installing Product Key for Home Edition...
            cscript //nologo slmgr.vbs /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99 >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Home N" >nul && (
            echo Installing Product Key for Home N Edition...
            cscript //nologo slmgr.vbs /ipk 3KHY7-WNT83-DGQKR-F7HPR-844BM >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Home Single Language" >nul && (
            echo Installing Product Key for Home Single Language Edition...
            cscript //nologo slmgr.vbs /ipk 7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Home Country Specific" >nul && (
            echo Installing Product Key for Home Country Specific Edition...
            cscript //nologo slmgr.vbs /ipk PVMJN-6DFY6-9CCP6-7BKTT-D3WVR >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Pro" >nul && (
            echo Installing Product Key for Pro Edition...
            cscript //nologo slmgr.vbs /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Pro N" >nul && (
            echo Installing Product Key for Pro N Edition...
            cscript //nologo slmgr.vbs /ipk MH37W-N47XK-V7XM9-C7227-GCQG9 >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Pro for Workstations" >nul && (
            echo Installing Product Key for Pro for Workstations Edition...
            cscript //nologo slmgr.vbs /ipk NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Pro for Workstations N" >nul && (
            echo Installing Product Key for Pro for Workstations N Edition...
            cscript //nologo slmgr.vbs /ipk 9FNHH-K3HBT-3W4TD-6383H-6XYWF>nul
            goto :activate
        ) || wmic os get name | findstr /c:"Pro Education" >nul && (
            echo Installing Product Key for Pro Education Edition...
            cscript //nologo slmgr.vbs /ipk 6TP4R-GNPTD-KYYHQ-7B7DP-J447Y >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Pro Education N" >nul && (
            echo Installing Product Key for Pro Education N Edition...
            cscript //nologo slmgr.vbs /ipk YVWGF-BXNMC-HTQYQ-CPQ99-66QFC >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Education" >nul && (
            echo Installing Product Key for Education Edition...
            cscript //nologo slmgr.vbs /ipk NW6C2-QMPVW-D7KKK-3GKT6-VCFB2 >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Education N" >nul && (
            echo Installing Product Key for Education N Edition...
            cscript //nologo slmgr.vbs /ipk 2WH4N-8QGBV-H22JP-CT43Q-MDWWJ >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise" >nul && (
            echo Installing Product Key for Enterprise Edition...
            cscript //nologo slmgr.vbs /ipk NPPR9-FWDCX-D2C8J-H872K-2YT43 >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise N" >nul && (
            echo Installing Product Key for Enterprise N Edition...
            cscript //nologo slmgr.vbs /ipk DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4 >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise G" >nul && (
            echo Installing Product Key for Enterprise G Edition...
            cscript //nologo slmgr.vbs /ipk YYVX9-NTFWV-6MDM3-9PT4T-4M68B >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise G N" >nul && (
            echo Installing Product Key for Enterprise G N Edition...
            cscript //nologo slmgr.vbs /ipk 44RPN-FTY23-9VTTB-MP9BX-T84FV >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise LTSC 2019" >nul && (
            echo Installing Product Key for Enterprise LTSC 2019 Edition...
            cscript //nologo slmgr.vbs /ipk M7XTQ-FN8P6-TTKYV-9D4CC-J462D >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise N LTSC 2019" >nul && (
            echo Installing Product Key for Enterprise N LTSC 2019 Edition...
            cscript //nologo slmgr.vbs /ipk 92NFX-8DJQP-P6BBQ-THF9C-7CG2H >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise LTSB 2016" >nul && (
            echo Installing Product Key for Enterprise LTSB 2016 Edition...
            cscript //nologo slmgr.vbs /ipk DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise N LTSB 2016" >nul && (
            echo Installing Product Key for Enterprise N LTSB 2016 Edition...
            cscript //nologo slmgr.vbs /ipk QFFDN-GRT3P-VKWWX-X7T3R-8B639 >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise 2015 LTSB" >nul && (
            echo Installing Product Key for Enterprise 2015 LTSB Edition...
            cscript //nologo slmgr.vbs /ipk WNMTR-4C88C-JK8YV-HQ7T2-76DF9 >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise 2015 LTSB N" >nul && (
            echo Installing Product Key for Enterprise 2015 LTSB N Edition...
            cscript //nologo slmgr.vbs /ipk 2F77B-TNFGY-69QQF-B8YKP-D69TJ >nul
            goto :activate
        ) || (
            echo Installing the previous product key...
            cscript //nologo slmgr.vbs /ipk %oldKey% >nul
            set "err=Sorry, your current version is not supported at this time."
            goto :error
        )
    ) else (
        set "err=Please connect to the internet to fully activate your Windows 11."
        goto :error
    )
) else (
    set "err=Please open this file by %ESC%[1mRun as adminstrator."
    goto :error
)
exit /b 0

:win10
title Activate Windows 10

cls
call header.bat
echo The following actions must be in Administrator mode. It will only modify or change the product key within your Windows 10. This will not change any system files. It has nothing to do with your Windows 10 activation if you want to report any errors. You can contact the developer in a number of ways through the developer's website.
echo.
echo  Required:
echo  - Administrator
echo  - Network
echo  Supported products:
echo  - [Home, Home N, Home Single Language, Home Country Specific]
echo  - [Pro, Pro N, Pro for Workstations, Pro for Workstations N]
echo  - [Education, Education N]
echo  - [Enterprise, Enterprise N, Enterprise G, Enterprise G N,
echo     Enterprise LTSC 2019, Enterprise N LTSC 2019,
echo     Enterprise LTSB 2016, Enterprise N LTSB 2016,
echo     Enterprise 2015 LTSB, Enterprise 2015 LTSB N]
echo.
echo %ESC%[7mYou are running on Windows 10.%ESC%[0m
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
        echo Preparing to activate your Windows 10...
        for /f "delims=" %%a in ('cscript //nologo pdk.vbs') do (
            set oldKey=%%a
        )

        echo Configuring the original server...
        cscript //nologo slmgr.vbs /ckms >nul

        echo Uninstalling the current product key...
        cscript //nologo slmgr.vbs /upk >nul

        echo Clearing the old product code registration file...
        cscript //nologo slmgr.vbs /cpky >nul
            
        echo Searching for your edition of Windows 10...
        wmic os get name | findstr /c:"Home" >nul && (
            echo Installing Product Key for Home Edition...
            cscript //nologo slmgr.vbs /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99 >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Home N" >nul && (
            echo Installing Product Key for Home N Edition...
            cscript //nologo slmgr.vbs /ipk 3KHY7-WNT83-DGQKR-F7HPR-844BM >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Home Single Language" >nul && (
            echo Installing Product Key for Home Single Language Edition...
            cscript //nologo slmgr.vbs /ipk 7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Home Country Specific" >nul && (
            echo Installing Product Key for Home Country Specific Edition...
            cscript //nologo slmgr.vbs /ipk PVMJN-6DFY6-9CCP6-7BKTT-D3WVR >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Pro" >nul && (
            echo Installing Product Key for Pro Edition...
            cscript //nologo slmgr.vbs /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Pro N" >nul && (
            echo Installing Product Key for Pro N Edition...
            cscript //nologo slmgr.vbs /ipk MH37W-N47XK-V7XM9-C7227-GCQG9 >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Pro for Workstations" >nul && (
            echo Installing Product Key for Pro for Workstations Edition...
            cscript //nologo slmgr.vbs /ipk NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Pro for Workstations N" >nul && (
            echo Installing Product Key for Pro for Workstations N Edition...
            cscript //nologo slmgr.vbs /ipk 9FNHH-K3HBT-3W4TD-6383H-6XYWF>nul
            goto :activate
        ) || wmic os get name | findstr /c:"Pro Education" >nul && (
            echo Installing Product Key for Pro Education Edition...
            cscript //nologo slmgr.vbs /ipk 6TP4R-GNPTD-KYYHQ-7B7DP-J447Y >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Pro Education N" >nul && (
            echo Installing Product Key for Pro Education N Edition...
            cscript //nologo slmgr.vbs /ipk YVWGF-BXNMC-HTQYQ-CPQ99-66QFC >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Education" >nul && (
            echo Installing Product Key for Education Edition...
            cscript //nologo slmgr.vbs /ipk NW6C2-QMPVW-D7KKK-3GKT6-VCFB2 >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Education N" >nul && (
            echo Installing Product Key for Education N Edition...
            cscript //nologo slmgr.vbs /ipk 2WH4N-8QGBV-H22JP-CT43Q-MDWWJ >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise" >nul && (
            echo Installing Product Key for Enterprise Edition...
            cscript //nologo slmgr.vbs /ipk NPPR9-FWDCX-D2C8J-H872K-2YT43 >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise N" >nul && (
            echo Installing Product Key for Enterprise N Edition...
            cscript //nologo slmgr.vbs /ipk DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4 >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise G" >nul && (
            echo Installing Product Key for Enterprise G Edition...
            cscript //nologo slmgr.vbs /ipk YYVX9-NTFWV-6MDM3-9PT4T-4M68B >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise G N" >nul && (
            echo Installing Product Key for Enterprise G N Edition...
            cscript //nologo slmgr.vbs /ipk 44RPN-FTY23-9VTTB-MP9BX-T84FV >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise LTSC 2019" >nul && (
            echo Installing Product Key for Enterprise LTSC 2019 Edition...
            cscript //nologo slmgr.vbs /ipk M7XTQ-FN8P6-TTKYV-9D4CC-J462D >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise N LTSC 2019" >nul && (
            echo Installing Product Key for Enterprise N LTSC 2019 Edition...
            cscript //nologo slmgr.vbs /ipk 92NFX-8DJQP-P6BBQ-THF9C-7CG2H >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise LTSB 2016" >nul && (
            echo Installing Product Key for Enterprise LTSB 2016 Edition...
            cscript //nologo slmgr.vbs /ipk DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise N LTSB 2016" >nul && (
            echo Installing Product Key for Enterprise N LTSB 2016 Edition...
            cscript //nologo slmgr.vbs /ipk QFFDN-GRT3P-VKWWX-X7T3R-8B639 >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise 2015 LTSB" >nul && (
            echo Installing Product Key for Enterprise 2015 LTSB Edition...
            cscript //nologo slmgr.vbs /ipk WNMTR-4C88C-JK8YV-HQ7T2-76DF9 >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise 2015 LTSB N" >nul && (
            echo Installing Product Key for Enterprise 2015 LTSB N Edition...
            cscript //nologo slmgr.vbs /ipk 2F77B-TNFGY-69QQF-B8YKP-D69TJ >nul
            goto :activate
        ) || (
            echo Installing the previous product key...
            cscript //nologo slmgr.vbs /ipk %oldKey% >nul
            set "err=Sorry, your current version is not supported at this time."
            goto :error
        )
    ) else (
        set "err=Please connect to the internet to fully activate your Windows 10."
        goto :error
    )
) else (
    set "err=Please open this file by %ESC%[1mRun as adminstrator."
    goto :error
)
exit /b 0

:win8.1
title Activate Windows 8.1

cls
call header.bat
echo The following actions must be in Administrator mode. It will only modify or change the product key within your Windows 8.1. This will not change any system files. It has nothing to do with your Windows 8.1 activation if you want to report any errors. You can contact the developer in a number of ways through the developer's website.
echo.
echo  Required:
echo  - Administrator
echo  - Network
echo  Supported products:
echo  - [Pro, Pro N]
echo  - [Enterprise, Enterprise N]
echo.
echo %ESC%[7mYou are running on Windows 8.1.%ESC%[0m
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
        echo Preparing to activate your Windows 8.1...
        for /f "delims=" %%a in ('cscript //nologo pdk.vbs') do (
            set oldKey=%%a
        )

        echo Configuring the original server...
        cscript //nologo slmgr.vbs /ckms >nul

        echo Uninstalling the current product key...
        cscript //nologo slmgr.vbs /upk >nul

        echo Clearing the old product code registration file...
        cscript //nologo slmgr.vbs /cpky >nul
            
        echo Searching for your edition of Windows 8.1...
        wmic os get name | findstr /c:"Pro" >nul && (
            echo Installing Product Key for Pro Edition...
            cscript //nologo slmgr.vbs /ipk GCRJD-8NW9H-F2CDX-CCM8D-9D6T9 >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Pro N" >nul && (
            echo Installing Product Key for Pro N Edition...
            cscript //nologo slmgr.vbs /ipk HMCNV-VVBFX-7HMBH-CTY9B-B4FXY >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise" >nul && (
            echo Installing Product Key for Enterprise Edition...
            cscript //nologo slmgr.vbs /ipk MHF9N-XY6XB-WVXMC-BTDCT-MKKG7 >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise N" >nul && (
            echo Installing Product Key for Enterprise N Edition...
            cscript //nologo slmgr.vbs /ipk TT4HM-HN7YT-62K67-RGRQJ-JFFXW >nul
            goto :activate
        ) || (
            echo Installing the previous product key...
            cscript //nologo slmgr.vbs /ipk %oldKey% >nul
            set "err=Sorry, your current version is not supported at this time."
            goto :error
        )
    ) else (
        set "err=Please connect to the internet to fully activate your Windows 8.1."
        goto :error
    )
) else (
    set "err=Please open this file by %ESC%[1mRun as adminstrator."
    goto :error
)
exit /b 0

:win8
title Activate Windows 8

cls
call header.bat
echo The following actions must be in Administrator mode. It will only modify or change the product key within your Windows 8. This will not change any system files. It has nothing to do with your Windows 8 activation if you want to report any errors. You can contact the developer in a number of ways through the developer's website.
echo.
echo  Required:
echo  - Administrator
echo  - Network
echo  Supported products:
echo  - [Pro, Pro N]
echo  - [Enterprise, Enterprise N]
echo.
echo %ESC%[7mYou are running on Windows 8.%ESC%[0m
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
        echo Preparing to activate your Windows 8...
        for /f "delims=" %%a in ('cscript //nologo pdk.vbs') do (
            set oldKey=%%a
        )

        echo Configuring the original server...
        cscript //nologo slmgr.vbs /ckms >nul

        echo Uninstalling the current product key...
        cscript //nologo slmgr.vbs /upk >nul

        echo Clearing the old product code registration file...
        cscript //nologo slmgr.vbs /cpky >nul
            
        echo Searching for your edition of Windows 8...
        wmic os get name | findstr /c:"Pro" >nul && (
            echo Installing Product Key for Pro Edition...
            cscript //nologo slmgr.vbs /ipk NG4HW-VH26C-733KW-K6F98-J8CK4 >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Pro N" >nul && (
            echo Installing Product Key for Pro N Edition...
            cscript //nologo slmgr.vbs /ipk XCVCF-2NXM9-723PB-MHCB7-2RYQQY >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise" >nul && (
            echo Installing Product Key for Enterprise Edition...
            cscript //nologo slmgr.vbs /ipk 32JNW-9KQ84-P47T8-D8GGY-CWCK7 >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise N" >nul && (
            echo Installing Product Key for Enterprise N Edition...
            cscript //nologo slmgr.vbs /ipk JMNMF-RHW7P-DMY6X-RF3DR-X2BQT >nul
            goto :activate
        ) || (
            echo Installing the previous product key...
            cscript //nologo slmgr.vbs /ipk %oldKey% >nul
            set "err=Sorry, your current version is not supported at this time."
            goto :error
        )
    ) else (
        set "err=Please connect to the internet to fully activate your Windows 8."
        goto :error
    )
) else (
    set "err=Please open this file by %ESC%[1mRun as adminstrator."
    goto :error
)
exit /b 0

:win7
title Activate Windows 7

cls
call header.bat
echo The following actions must be in Administrator mode. It will only modify or change the product key within your Windows 7. This will not change any system files. It has nothing to do with your Windows 7 activation if you want to report any errors. You can contact the developer in a number of ways through the developer's website.
echo.
echo  Required:
echo  - Administrator
echo  - Network
echo  Supported products:
echo  - [Professional, Professional N, Professional E]
echo  - [Enterprise, Enterprise N, Enterprise E]
echo.
echo %ESC%[7mYou are running on Windows 7.%ESC%[0m
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
        echo Preparing to activate your Windows 7...
        for /f "delims=" %%a in ('cscript //nologo pdk.vbs') do (
            set oldKey=%%a
        )

        echo Configuring the original server...
        cscript //nologo slmgr.vbs /ckms >nul

        echo Uninstalling the current product key...
        cscript //nologo slmgr.vbs /upk >nul

        echo Clearing the old product code registration file...
        cscript //nologo slmgr.vbs /cpky >nul
            
        echo Searching for your edition of Windows 7...
        wmic os get name | findstr /c:"Professional" >nul && (
            echo Installing Product Key for Professional Edition...
            cscript //nologo slmgr.vbs /ipk FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4 >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Professional N" >nul && (
            echo Installing Product Key for Professional N Edition...
            cscript //nologo slmgr.vbs /ipk MRPKT-YTG23-K7D7T-X2JMM-QY7MG >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Professional E" >nul && (
            echo Installing Product Key for Professional E Edition...
            cscript //nologo slmgr.vbs /ipk W82YF-2Q76Y-63HXB-FGJG9-GF7QX >nul
            goto :activate
        )|| wmic os get name | findstr /c:"Enterprise" >nul && (
            echo Installing Product Key for Enterprise Edition...
            cscript //nologo slmgr.vbs /ipk 33PXH-7Y6KF-2VJC9-XBBR8-HVTHH >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise N" >nul && (
            echo Installing Product Key for Enterprise N Edition...
            cscript //nologo slmgr.vbs /ipk YDRBP-3D83W-TY26F-D46B2-XCKRJ >nul
            goto :activate
        ) || wmic os get name | findstr /c:"Enterprise E" >nul && (
            echo Installing Product Key for Enterprise E Edition...
            cscript //nologo slmgr.vbs /ipk C29WB-22CC8-VJ326-GHFJW-H9DH4 >nul
            goto :activate
        ) || (
            echo Installing the previous product key...
            cscript //nologo slmgr.vbs /ipk %oldKey% >nul
            set "err=Sorry, your current version is not supported at this time."
            goto :error
        )
    ) else (
        set "err=Please connect to the internet to fully activate your Windows 7."
        goto :error
    )
) else (
    set "err=Please open this file by %ESC%[1mRun as adminstrator."
    goto :error
)
exit /b 0



:activate
echo Prepare for activation...
set cserver=16
set schoose=1

echo Checking %cserver% provider servers...
if %schoose% == 1 set server=kms.digiboy.ir:1688
if %schoose% == 2 set server=hq1.chinancce.com:1688
if %schoose% == 3 set server=54.223.212.31:1688
if %schoose% == 4 set server=kms.cnlic.com:1688
if %schoose% == 5 set server=kms.chinancce.com:1688
if %schoose% == 6 set server=kms.ddns.net:1688
if %schoose% == 7 set server=franklv.ddns.net:1688
if %schoose% == 8 set server=k.zpale.com:1688
if %schoose% == 9 set server=m.zpale.com:1688
if %schoose% == 10 set server=mvg.zpale.com:1688
if %schoose% == 11 set server=kms.shuax.com:1688
if %schoose% == 12 set server=kensol263.imwork.net:1688
if %schoose% == 13 set server=xykz.f3322.org:1688
if %schoose% == 14 set server=kms789.com:1688
if %schoose% == 15 set server=dimanyakms.sytes.net:1688
if %schoose% == 16 set server=kms.03k.org:1688
if %schoose% == 17 (
    echo Installing the previous product key...
    cscript //nologo slmgr.vbs /ipk %oldKey% >nul
    set "err=Sorry, it seems unable to connect to any provider server at all. & echo.Please try again later."
    goto :error
)

echo Connecting to a server serving installed licenses...
cscript //nologo slmgr.vbs /skms %server% >nul

echo Activating your Windows %win%...
cscript //nologo slmgr.vbs /ato | find /i "successfully" >nul && (
    echo %ESC%[92m^<^=^> Complete the process%ESC%[0m
    echo.
    pause

    cls
    call header.bat
    echo Now you are ready to go on.
    echo Now that you've activated your Windows %win%, you can set some features that aren't available. Please explore the new features that are now enabled in your Windows %win%. Have fun.
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
goto :win%win%
exit /b 0

endlocal
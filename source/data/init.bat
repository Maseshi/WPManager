mode con: cols=66 lines=30
chcp 65001

set "v=1.0.2.0"
set "n=Windows Product Manager"
set "verify=0"

for /f "tokens=1,2 delims=#" %%a in ('"prompt #$H#$E# & echo on & for %%b in (1) do rem"') do (
    set ESC=%%b
)
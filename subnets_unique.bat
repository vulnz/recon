@echo off
setlocal enabledelayedexpansion

if not exist dcs.txt (
    echo [!] File dcs.txt not found!
    exit /b
)

echo === DC Subnet Discovery Report === > report.txt
echo Start time: %date% %time% >> report.txt
echo. >> report.txt

del subnets.tmp 2>nul

for /f "usebackq tokens=* delims=" %%A in ("dcs.txt") do (

    echo Checking: %%A >> report.txt
    set "resolvedIP="

    rem --------- TRY TO RESOLVE IP FROM FIRST PING LINE ---------
    for /f "tokens=2 delims=[]" %%I in ('ping -n 1 "%%A" ^| find "["') do (
        set "resolvedIP=%%I"
    )

    if defined resolvedIP (
        echo     Ping IP: !resolvedIP! >> report.txt

        for /f "tokens=1-3 delims=." %%a in ("!resolvedIP!") do (
            echo %%a.%%b.%%c.0/24 >> subnets.tmp
        )
        echo. >> report.txt
        goto nextDC
    )

    rem --------- FALLBACK: NSLOOKUP ---------
    for /f "tokens=2 delims=: " %%N in ('nslookup "%%A" ^| find "Address"') do (
        set "resolvedIP=%%N"
    )

    if defined resolvedIP (
        echo     NSLookup IP: !resolvedIP! >> report.txt

        for /f "tokens=1-3 delims=." %%a in ("!resolvedIP!") do (
            echo %%a.%%b.%%c.0/24 >> subnets.tmp
        )
    ) else (
        echo     ERROR: Cannot resolve IP >> report.txt
    )

    echo. >> report.txt

:nextDC
)

echo === UNIQUE SUBNETS === >> report.txt
echo. >> report.txt

if not exist subnets.tmp (
    echo NO SUBNETS FOUND >> report.txt
    echo Done
    pause
    exit /b
)

sort subnets.tmp > subnets_sorted.tmp
del subnets_unique.tmp 2>nul

for /f "usebackq tokens=* delims=" %%L in ("subnets_sorted.tmp") do (
    findstr /x /c:"%%L" subnets_unique.tmp >nul || echo %%L>>subnets_unique.tmp
)

type subnets_unique.tmp >> report.txt

del subnets.tmp subnets_sorted.tmp subnets_unique.tmp 2>nul

echo Done → report.txt
pause

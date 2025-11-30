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

    rem --- try resolve via ping ---
    set "ip_ping="
    for /f "tokens=2 delims=: " %%p in ('ping -n 1 "%%A" ^| findstr /r /c:"Reply from " /c:"ответ от"') do (
        rem strip spaces
        set "ip_ping=%%p"
    )

    if defined ip_ping (
        echo    Ping IP: !ip_ping! >> report.txt

        rem extract /24 subnet
        for /f "tokens=1-3 delims=." %%a in ("!ip_ping!") do (
            echo %%a.%%b.%%c.0/24 >> subnets.tmp
        )

        goto nextDC
    )

    rem --- fallback via nslookup ---
    set "ip_ns="
    for /f "tokens=2 delims=: " %%n in ('nslookup "%%A" ^| findstr /i "Address"') do (
        set "ip_ns=%%n"
    )

    if defined ip_ns (
        echo    NSLookup IP: !ip_ns! >> report.txt

        rem extract subnet
        for /f "tokens=1-3 delims=." %%a in ("!ip_ns!") do (
            echo %%a.%%b.%%c.0/24 >> subnets.tmp
        )
    ) else (
        echo    ERROR: cannot resolve IP >> report.txt
    )

:nextDC
    echo. >> report.txt
)

echo === UNIQUE SUBNETS === >> report.txt
echo. >> report.txt

if not exist subnets.tmp (
    echo NO SUBNETS FOUND! >> report.txt
    echo Done.
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

echo.
echo DONE → report.txt
pause

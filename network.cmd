@echo off
setlocal enabledelayedexpansion

echo ===============================================
echo   CMD-ONLY Multi-Subnet Discovery Scanner
echo   (Find reachable 172.x and 10.x networks)
echo ===============================================
echo.

set FOUND_SUBNETS=

:: -----------------------------------------
:: 1. Detect LOCAL subnet from active NIC
:: -----------------------------------------
for /f "tokens=2 delims=:" %%A in ('ipconfig ^| findstr /i "IPv4 Address"') do (
    set IP=%%A
    set IP=!IP: =!
)

for /f "tokens=1-3 delims=." %%a in ("!IP!") do (
    set LOCAL_SUBNET=%%a.%%b.%%c
)

echo Local subnet found: %LOCAL_SUBNET%.0/24
set FOUND_SUBNETS=%LOCAL_SUBNET%
echo.

:: -----------------------------------------
:: 2. Extract DNS servers (likely 10.x ranges)
:: -----------------------------------------
echo Detecting DNS subnets...

for /f "tokens=2 delims=:" %%D in ('ipconfig ^| findstr /i "DNS Servers"') do (
    set D1=%%D
    set D1=!D1: =!
)

for /f "tokens=2 delims=:" %%D in ('ipconfig ^| findstr /i /c:"                                " "DNS Servers"') do (
    set D2=%%D
    set D2=!D2: =!
)

for %%X in (!D1! !D2!) do (
    for /f "tokens=1-3 delims=." %%a in ("%%X") do (
        echo DNS subnet found: %%a.%%b.%%c
        echo %%a.%%b.%%c>>subnets.tmp
    )
)

echo.

:: -----------------------------------------
:: 3. Remove duplicates
:: -----------------------------------------
sort subnets.tmp /unique > uniq.tmp
del subnets.tmp

echo Subnets identified for reachability:
type uniq.tmp
echo.

:: -----------------------------------------
:: 4. Test reachability for each subnet
:: -----------------------------------------
echo Testing which subnets respond to pings...
> reachable.tmp echo.

for /f %%S in (uniq.tmp) do (
    echo Pinging %%S.1 ...
    ping -n 1 -w 100 %%S.1 | find "TTL=" >nul
    if not errorlevel 1 (
        echo REACHABLE: %%S.0/24
        echo %%S>>reachable.tmp
    ) else (
        echo No response from %%S.1 (skipped)
    )
)
echo.

:: -----------------------------------------
:: 5. Scan reachable subnets
:: -----------------------------------------
echo ========================================
echo Starting ping sweeps on reachable networks
echo ========================================
echo.

for /f %%N in (reachable.tmp) do (
    echo Scanning subnet %%N.0/24...
    set OUT=scan_%%N.txt
    echo Alive hosts in %%N.0/24: > !OUT!

    for /L %%i in (1,1,254) do (
        ping -n 1 -w 60 %%N.%%i | find "TTL=" >nul
        if not errorlevel 1 (
            echo Host alive: %%N.%%i
            echo %%N.%%i >> !OUT!
        )
    )
    echo Completed subnet: %%N.0/24
    echo.
)

echo All reachable subnets scanned.
echo Results saved as: scan_<subnet>.txt
del uniq.tmp reachable.tmp 2>nul

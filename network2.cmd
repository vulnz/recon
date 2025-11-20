@echo off
setlocal enabledelayedexpansion

echo ===============================================
echo   UNIVERSAL MULTI-SUBNET DISCOVERY SCANNER
echo   (Language-independent, finds 172.x + 10.x)
echo ===============================================
echo.

:: Delete old temp files
del subnets.tmp 2>nul
del uniq.tmp 2>nul
del reachable.tmp 2>nul

echo Searching for ALL IPv4-style addresses...

:: Extract ALL IPv4-like patterns from ipconfig
for /f "tokens=1-4 delims=." %%a in ('ipconfig ^| findstr /r "[0-9]*\.[0-9]*\.[0-9]*\.[0-9]*"') do (
    set A=%%a.%%b.%%c.%%d

    :: Validate IPv4 (simple check)
    echo !A! | findstr /r "^[0-9][0-9]*\.[0-9][0-9]*\.[0-9][0-9]*\.[0-9][0-9]*$" >nul
    if not errorlevel 1 (
        for /f "tokens=1-3 delims=." %%x in ("!A!") do (
            set NET=%%x.%%y.%%z
            echo !NET!>>subnets.tmp
        )
    )
)

:: Add local default subnet again to be safe
for /f "tokens=2 delims=:" %%A in ('ipconfig ^| findstr /i "IPv4"') do (
    set IP=%%A
    set IP=!IP: =!
)
for /f "tokens=1-3 delims=." %%a in ("!IP!") do (
    echo %%a.%%b.%%c>>subnets.tmp
)

echo.
echo All found candidate subnets:
type subnets.tmp
echo.

:: Remove duplicates
sort subnets.tmp /unique > uniq.tmp

echo Clean subnet list (unique):
type uniq.tmp
echo.

:: Test reachability of each subnet
echo Testing which subnets are reachable...
> reachable.tmp echo.

for /f %%S in (uniq.tmp) do (
    echo Checking %%S.1 ...
    ping -n 1 -w 80 %%S.1 | find "TTL=" >nul
    if not errorlevel 1 (
        echo REACHABLE: %%S
        echo %%S>>reachable.tmp
    )
)

echo.
echo Reachable subnets:
type reachable.tmp
echo.

echo ===============================================
echo Scanning all reachable networks (/24)
echo ===============================================

for /f %%N in (reachable.tmp) do (
    set OUT=scan_%%N.txt
    echo Alive hosts in %%N.0/24 > !OUT!

    echo Scanning %%N.0/24 ...
    for /L %%i in (1,1,254) do (
        ping -n 1 -w 60 %%N.%%i | find "TTL=" >nul
        if not errorlevel 1 (
            echo %%N.%%i >> !OUT!
            echo Alive: %%N.%%i
        )
    )
    echo Done: %%N.0/24
    echo.
)

echo ===============================================
echo ALL SCANS COMPLETE
echo Results saved to scan_<subnet>.txt
echo ===============================================


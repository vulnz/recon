@echo off
setlocal enabledelayedexpansion

echo ===============================================
echo   SAFE LOCAL NETWORK DISCOVERY (CMD ONLY)
echo ===============================================
echo.

:: Detect active IPv4
for /f "tokens=2 delims=:" %%A in ('ipconfig ^| findstr /r /c:"IPv4"') do (
    set IP=%%A
    set IP=!IP: =!
)

echo Detected IP: %IP%

:: Compute /24 subnet
for /f "tokens=1-3 delims=." %%a in ("%IP%") do (
    set SUBNET=%%a.%%b.%%c
)

echo Scanning subnet: %SUBNET%.0/24
echo.

set OUTPUT=scan_%SUBNET%.txt
echo Alive hosts in %SUBNET%.0/24 > %OUTPUT%

for /L %%i in (1,1,254) do (
    ping -n 1 -w 70 %SUBNET%.%%i | find "TTL=" >nul
    if not errorlevel 1 ( 
        echo %SUBNET%.%%i >> %OUTPUT%
        echo Alive: %SUBNET%.%%i
    )
)

echo.
echo Scan done. Results saved to %OUTPUT%

@echo off
setlocal enabledelayedexpansion

if not exist dcs.txt (
    echo [!] File dcs.txt not found!
    exit /b
)

echo === Pinging Domain Controllers and Extracting Subnets === > report.txt

del subnets.tmp 2>nul

for /f "usebackq tokens=*" %%A in ("dcs.txt") do (
    echo Checking: %%A >> report.txt

    for /f "skip=1 tokens=2 delims=[]" %%I in ('ping -n 1 %%A ^| find "["') do (
        rem Extract IP
        set ip=%%I

        rem Remove last octet
        for /f "tokens=1-3 delims=." %%a in ("!ip!") do (
            echo %%a.%%b.%%c.0/24 >> subnets.tmp
        )
    )
)

echo. >> report.txt
echo === Unique Subnets === >> report.txt
echo. >> report.txt

rem --- uniq implementation ---
sort subnets.tmp > subnets_sorted.tmp
del subnets_unique.tmp 2>nul

for /f "usebackq tokens=*" %%L in ("subnets_sorted.tmp") do (
    findstr /x /c:"%%L" subnets_unique.tmp >nul || echo %%L>>subnets_unique.tmp
)

type subnets_unique.tmp >> report.txt

del subnets.tmp subnets_sorted.tmp subnets_unique.tmp

echo Done! Output → report.txt

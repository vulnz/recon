@echo off
echo ====================================================
echo   TASK SCHEDULER SECURITY AUDIT (SAFE CHECK)
echo ====================================================
echo.

set out=%temp%\tasks_audit.txt
del "%out%" >nul 2>&1

echo [1] Exporting all tasks... >> "%out%"
schtasks /query /fo LIST /v >> "%out%"

echo ----------------------------------------------------
echo Checking for SYSTEM tasks with writable or missing paths...
echo ----------------------------------------------------

for /f "tokens=1,* delims=:" %%a in ('findstr /i /c:"Task To Run" "%out%"') do (
    set "path=%%b"
    call :trim path "%path%"
    
    if not exist "%path%" (
        echo [MISSING] %%b
    ) else (
        rem Check if path is in writable dirs
        echo "%path%" | findstr /i "AppData Temp ProgramData Public" >nul
        if not errorlevel 1 (
            echo [UNSAFE-PATH] %%b
        )
    )
)

echo ----------------------------------------------------
echo Checking for SYSTEM RunLevel=Highest tasks...
echo ----------------------------------------------------

findstr /i /c:"SYSTEM" /c:"Run Level: Highest" "%out%" 

echo.
echo Done.
goto :eof

:trim
setlocal enableextensions enabledelayedexpansion
set var=%1
set val=%~2
for /f "tokens=* delims= " %%a in ("%val%") do set val=%%a
for /l %%a in (1,1,50) do if "!val:~-1!"==" " set val=!val:~0,-1!
endlocal & set %var%=%val%
goto :eof

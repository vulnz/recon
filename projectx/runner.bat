@echo off
setlocal enabledelayedexpansion

:: ===============================
:: CHECK ARGUMENTS
:: ===============================
if "%~1"=="" (
    echo Usage: run.bat payload.bin [arguments]
    echo Example: run.bat payload.bin --help
    pause
    exit /b
)

:: First argument = payload file
set PAYLOAD=%1

:: Shift so %* now contains only extra params
shift

:: ===============================
:: LOG FILE
:: ===============================
set LOGFILE=run_output.txt

echo ================================ >> %LOGFILE%
echo [Run started] %date% %time% >> %LOGFILE%
echo Payload: %PAYLOAD% >> %LOGFILE%
echo Arguments: %* >> %LOGFILE%
echo ================================ >> %LOGFILE%

echo [*] Running Java with payload: %PAYLOAD%
echo [*] Additional params: %*
echo.

:: ===============================
:: RUN Java → save output to a temp file
:: ===============================
java -cp ".;jna-5.13.0.jar" ShellcodeRunner "%PAYLOAD%" %* > temp_output.txt

:: ===============================
:: DISPLAY OUTPUT
:: ===============================
echo [*] Output:
echo ---------------------------------
type temp_output.txt
echo ---------------------------------

:: ===============================
:: APPEND OUTPUT TO LOG
:: ===============================
type temp_output.txt >> %LOGFILE%

del temp_output.txt

echo. >> %LOGFILE%
echo [Run finished] %date% %time% >> %LOGFILE%
echo =============================================== >> %LOGFILE%
echo.

echo [*] Output saved to %LOGFILE%
echo [*] Press any key to exit...
pause >nul

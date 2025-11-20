@echo off
set OUTPUT=recon_%COMPUTERNAME%_%DATE:~10,4%-%DATE:~4,2%-%DATE:~7,2%.txt

echo ======================================================= > %OUTPUT%
echo   FULL RECON + AD ENUMERATION (CMD ONLY) >> %OUTPUT%
echo ======================================================= >> %OUTPUT%
echo. >> %OUTPUT%

REM ===================== SYSTEM =====================
echo [*] SYSTEM INFO >> %OUTPUT%
systeminfo >> %OUTPUT%
echo. >> %OUTPUT%

echo [*] GENERAL INFO >> %OUTPUT%
hostname >> %OUTPUT%
whoami >> %OUTPUT%
whoami /groups >> %OUTPUT%
whoami /priv >> %OUTPUT%
echo. >> %OUTPUT%

REM ===================== LOCAL USERS =====================
echo [*] LOCAL USERS >> %OUTPUT%
net user >> %OUTPUT%
echo. >> %OUTPUT%

echo [*] LOCAL GROUPS >> %OUTPUT%
net localgroup >> %OUTPUT%
echo. >> %OUTPUT%

echo [*] ADMINISTRATORS GROUP >> %OUTPUT%
net localgroup administrators >> %OUTPUT%
echo. >> %OUTPUT%

REM ===================== AD ENUM (NO POWERSHELL) =====================
echo [*] DOMAIN INFO >> %OUTPUT%
echo (errors mean workstation not joined to domain) >> %OUTPUT%
echo. >> %OUTPUT%

echo --- Current Domain --- >> %OUTPUT%
net config workstation >> %OUTPUT%
echo. >> %OUTPUT%

echo --- Domain Controllers --- >> %OUTPUT%
nltest /dclist:%USERDOMAIN% >> %OUTPUT% 2>&1
echo. >> %OUTPUT%

echo --- Domain Trusts --- >> %OUTPUT%
nltest /domain_trusts >> %OUTPUT% 2>&1
echo. >> %OUTPUT%

echo --- Domain Users --- >> %OUTPUT%
net user /domain >> %OUTPUT% 2>&1
echo. >> %OUTPUT%

echo --- Domain Groups --- >> %OUTPUT%
net group /domain >> %OUTPUT% 2>&1
echo. >> %OUTPUT%

echo --- Domain Admins --- >> %OUTPUT%
net group "Domain Admins" /domain >> %OUTPUT% 2>&1
echo. >> %OUTPUT%

echo --- Enterprise Admins --- >> %OUTPUT%
net group "Enterprise Admins" /domain >> %OUTPUT% 2>&1
echo. >> %OUTPUT%

echo --- Domain Computers --- >> %OUTPUT%
net group "Domain Computers" /domain >> %OUTPUT% 2>&1
echo. >> %OUTPUT%

echo --- Domain Controllers Group --- >> %OUTPUT%
net group "Domain Controllers" /domain >> %OUTPUT% 2>&1
echo. >> %OUTPUT%

echo --- Logon Servers --- >> %OUTPUT%
set logonserver >> %OUTPUT%
echo. >> %OUTPUT%

echo --- Site Info --- >> %OUTPUT%
nltest /dsgetsite >> %OUTPUT% 2>&1
echo. >> %OUTPUT%

echo --- Kerberos Info --- >> %OUTPUT%
nltest /kerberos:%USERDOMAIN% >> %OUTPUT% 2>&1
echo. >> %OUTPUT%

echo --- DC Connection Test --- >> %OUTPUT%
nltest /sc_verify:%USERDOMAIN% >> %OUTPUT% 2>&1
echo. >> %OUTPUT%

REM ===================== NETWORK =====================
echo [*] NETWORK CONFIGURATION >> %OUTPUT%
ipconfig /all >> %OUTPUT%
echo. >> %OUTPUT%

echo [*] ROUTES >> %OUTPUT%
route print >> %OUTPUT%
echo. >> %OUTPUT%

echo [*] ARP CACHE >> %OUTPUT%
arp -a >> %OUTPUT%
echo. >> %OUTPUT%

echo [*] DNS CACHE >> %OUTPUT%
ipconfig /displaydns >> %OUTPUT%
echo. >> %OUTPUT%

echo [*] OPEN CONNECTIONS >> %OUTPUT%
netstat -ano >> %OUTPUT%
echo. >> %OUTPUT%

echo [*] HOSTS FILE >> %OUTPUT%
type %SystemRoot%\System32\drivers\etc\hosts >> %OUTPUT%
echo. >> %OUTPUT%

REM ===================== NETWORK SHARES =====================
echo [*] LOCAL SHARES >> %OUTPUT%
net share >> %OUTPUT%
echo. >> %OUTPUT%

echo [*] SMB NEIGHBORHOOD >> %OUTPUT%
net view >> %OUTPUT%
net view /domain >> %OUTPUT% 2>&1
echo. >> %OUTPUT%

REM ===================== PROCESSES / SERVICES =====================
echo [*] RUNNING PROCESSES >> %OUTPUT%
tasklist /v >> %OUTPUT%
echo. >> %OUTPUT%

echo [*] SERVICES >> %OUTPUT%
sc query type= service state= all >> %OUTPUT%
echo. >> %OUTPUT%

REM ===================== STARTUP / SCHEDULED TASKS =====================
echo [*] STARTUP PROGRAMS >> %OUTPUT%
wmic startup get caption,command >> %OUTPUT% 2>&1
echo. >> %OUTPUT%

echo [*] SCHEDULED TASKS >> %OUTPUT%
schtasks /query /fo LIST /v >> %OUTPUT%
echo. >> %OUTPUT%

echo ======================================================= >> %OUTPUT%
echo DONE! Saved to: %OUTPUT% >> %OUTPUT%
echo =======================================================

echo.
echo ✔ RECON COMPLETE!
echo File saved as: %OUTPUT%

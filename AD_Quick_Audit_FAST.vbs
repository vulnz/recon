' ========================================
' OPTIMIZED AD Security Audit for LARGE Domains (VBS)
' Designed for 100K+ objects
' With timeouts and result limits
' ========================================

Option Explicit

Dim objFSO, objRootDSE, objConnection, objCommand, objRecordSet
Dim strDomainDN, strDomainName, strReportPath
Dim dictFindings
Dim intCritical, intHigh, intMedium
Dim MAX_RESULTS ' Limit results to prevent hanging

' Configuration
MAX_RESULTS = 100 ' Maximum results per query (adjust if needed)

' Initialize
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set dictFindings = CreateObject("Scripting.Dictionary")

intCritical = 0
intHigh = 0
intMedium = 0

' Get domain info
Set objRootDSE = GetObject("LDAP://RootDSE")
strDomainDN = objRootDSE.Get("defaultNamingContext")
strDomainName = Replace(strDomainDN, "DC=", "")
strDomainName = Replace(strDomainName, ",", ".")

WScript.Echo "========================================="
WScript.Echo "OPTIMIZED AD SECURITY AUDIT - LARGE DOMAINS"
WScript.Echo "========================================="
WScript.Echo "Domain: " & strDomainName
WScript.Echo "Max Results Per Check: " & MAX_RESULTS
WScript.Echo ""
WScript.Echo "Starting FAST audit..."
WScript.Echo ""

' ==================================================
' FAST CRITICAL CHECKS ONLY
' ==================================================

WScript.Echo "[+] PHASE 1: Critical Security Checks (Optimized)"
WScript.Echo ""

' 1. Unconstrained Delegation (MOST CRITICAL - Limited)
Call CheckUnconstrainedDelegationFast()

' 2. Kerberoastable Accounts (Limited)
Call CheckKerberoastableFast()

' 3. AS-REP Roastable (Limited)
Call CheckASREPRoastableFast()

' 4. KRBTGT Password Age
Call CheckKRBTGTFast()

' 5. Password Policy
Call CheckPasswordPolicyFast()

' 6. Password Never Expires (Sample)
Call CheckPasswordNeverExpiresFast()

' 7. MachineAccountQuota
Call CheckMachineAccountQuotaFast()

' 8. Domain Admins Count
Call CheckDomainAdminsFast()

' 9. Legacy OS (Sample)
Call CheckLegacyOSFast()

' 10. Domain Controllers
Call CheckDomainControllersFast()

WScript.Echo ""
WScript.Echo "========================================="
WScript.Echo "SCAN COMPLETE!"
WScript.Echo "========================================="
WScript.Echo ""

' Display Summary
Call DisplaySummary()

' Generate Report
Call GenerateQuickReport()

WScript.Echo ""
WScript.Echo "Report saved to: " & strReportPath
WScript.Echo ""
WScript.Echo "========================================="
WScript.Echo "For full audit use:"
WScript.Echo "  - Purple Knight (highly recommended!)"
WScript.Echo "  - PowerShell version for large domains"
WScript.Echo "========================================="

' ==================================================
' OPTIMIZED CHECK FUNCTIONS
' ==================================================

Sub CheckUnconstrainedDelegationFast()
    WScript.Echo "  [1.1] Checking Unconstrained Delegation (CRITICAL)..."
    
    On Error Resume Next
    
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    
    Set objCommand.ActiveConnection = objConnection
    objCommand.Properties("Page Size") = MAX_RESULTS
    objCommand.Properties("Searchscope") = 2
    objCommand.Properties("Timeout") = 60 ' 60 second timeout
    
    ' Query for unconstrained delegation (TRUSTED_FOR_DELEGATION = 524288)
    ' Exclude Domain Controllers (PrimaryGroupID 516)
    objCommand.CommandText = "SELECT sAMAccountName FROM 'LDAP://" & strDomainDN & "' WHERE objectCategory='computer' AND userAccountControl:1.2.840.113556.1.4.803:=524288 AND NOT primaryGroupID=516"
    
    Set objRecordSet = objCommand.Execute
    
    Dim intCount
    intCount = 0
    
    If Not objRecordSet.EOF Then
        objRecordSet.MoveFirst
        
        Dim strComputers
        strComputers = ""
        
        Do Until objRecordSet.EOF Or intCount >= 10 ' Show first 10
            If intCount < 10 Then
                If strComputers <> "" Then strComputers = strComputers & ", "
                strComputers = strComputers & objRecordSet.Fields("sAMAccountName").Value
            End If
            intCount = intCount + 1
            objRecordSet.MoveNext
        Loop
        
        If intCount > 0 Then
            Call AddFinding("CRITICAL", "Unconstrained Delegation", "Found " & intCount & " computer(s) with unconstrained delegation (showing first 10): " & strComputers & vbCrLf & "EXPLOITATION: PrinterBug + Unconstrained Delegation = DC compromise!", intCount)
            WScript.Echo "    WARNING: Found " & intCount & " computers with Unconstrained Delegation!"
            WScript.Echo "    First few: " & strComputers
            WScript.Echo "    REMEDIATION: Set-ADComputer -TrustedForDelegation $false"
        End If
    End If
    
    If intCount = 0 Then
        WScript.Echo "    OK: No unconstrained delegation found"
    End If
    
    objRecordSet.Close
    objConnection.Close
    
    On Error GoTo 0
End Sub

Sub CheckKerberoastableFast()
    WScript.Echo "  [1.2] Checking Kerberoastable Accounts..."
    
    On Error Resume Next
    
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    
    Set objCommand.ActiveConnection = objConnection
    objCommand.Properties("Page Size") = MAX_RESULTS
    objCommand.Properties("Searchscope") = 2
    objCommand.Properties("Timeout") = 60
    
    objCommand.CommandText = "SELECT sAMAccountName FROM 'LDAP://" & strDomainDN & "' WHERE objectCategory='person' AND servicePrincipalName=*"
    
    Set objRecordSet = objCommand.Execute
    
    Dim intCount
    intCount = 0
    
    If Not objRecordSet.EOF Then
        objRecordSet.MoveFirst
        
        Dim strAccounts
        strAccounts = ""
        
        Do Until objRecordSet.EOF Or intCount >= 10
            If intCount < 10 Then
                If strAccounts <> "" Then strAccounts = strAccounts & ", "
                strAccounts = strAccounts & objRecordSet.Fields("sAMAccountName").Value
            End If
            intCount = intCount + 1
            objRecordSet.MoveNext
        Loop
        
        If intCount > 0 Then
            Call AddFinding("CRITICAL", "Kerberoastable Accounts", "Found " & intCount & "+ accounts with SPN (showing first 10): " & strAccounts & vbCrLf & "EXPLOITATION: Rubeus.exe kerberoast /outfile:hashes.txt" & vbCrLf & "REMEDIATION: Convert to gMSA or set 25+ char passwords", intCount)
            WScript.Echo "    WARNING: Found " & intCount & "+ Kerberoastable accounts!"
            WScript.Echo "    First few: " & strAccounts
        End If
    End If
    
    If intCount = 0 Then
        WScript.Echo "    OK: No kerberoastable accounts found"
    End If
    
    objRecordSet.Close
    objConnection.Close
    
    On Error GoTo 0
End Sub

Sub CheckASREPRoastableFast()
    WScript.Echo "  [1.3] Checking AS-REP Roastable Accounts..."
    
    On Error Resume Next
    
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    
    Set objCommand.ActiveConnection = objConnection
    objCommand.Properties("Page Size") = MAX_RESULTS
    objCommand.Properties("Searchscope") = 2
    objCommand.Properties("Timeout") = 60
    
    ' DONT_REQ_PREAUTH = 4194304
    objCommand.CommandText = "SELECT sAMAccountName FROM 'LDAP://" & strDomainDN & "' WHERE objectCategory='person' AND userAccountControl:1.2.840.113556.1.4.803:=4194304"
    
    Set objRecordSet = objCommand.Execute
    
    Dim intCount
    intCount = 0
    
    If Not objRecordSet.EOF Then
        objRecordSet.MoveFirst
        
        Dim strAccounts
        strAccounts = ""
        
        Do Until objRecordSet.EOF Or intCount >= 10
            If intCount < 10 Then
                If strAccounts <> "" Then strAccounts = strAccounts & ", "
                strAccounts = strAccounts & objRecordSet.Fields("sAMAccountName").Value
            End If
            intCount = intCount + 1
            objRecordSet.MoveNext
        Loop
        
        If intCount > 0 Then
            Call AddFinding("CRITICAL", "AS-REP Roastable Accounts", "Found " & intCount & " accounts without pre-auth: " & strAccounts & vbCrLf & "EXPLOITATION: Rubeus.exe asreproast (NO PASSWORD NEEDED!)" & vbCrLf & "REMEDIATION: Set-ADAccountControl -DoesNotRequirePreAuth $false", intCount)
            WScript.Echo "    WARNING: Found " & intCount & " AS-REP Roastable accounts!"
            WScript.Echo "    Accounts: " & strAccounts
        End If
    End If
    
    If intCount = 0 Then
        WScript.Echo "    OK: No AS-REP roastable accounts found"
    End If
    
    objRecordSet.Close
    objConnection.Close
    
    On Error GoTo 0
End Sub

Sub CheckKRBTGTFast()
    WScript.Echo "  [1.4] Checking KRBTGT Password Age (GOLDEN TICKET RISK)..."
    
    On Error Resume Next
    
    Dim objKRBTGT, objPwdLastSet, dtPwdLastSet, intDaysSinceChange
    Set objKRBTGT = GetObject("LDAP://CN=krbtgt,CN=Users," & strDomainDN)
    
    Set objPwdLastSet = objKRBTGT.Get("pwdLastSet")
    
    If objPwdLastSet.HighPart <> 0 Or objPwdLastSet.LowPart <> 0 Then
        dtPwdLastSet = Int(objPwdLastSet.HighPart * 2^32 + objPwdLastSet.LowPart) / 600000000 - 11644473600
        dtPwdLastSet = DateAdd("s", dtPwdLastSet, #1/1/1970#)
        intDaysSinceChange = DateDiff("d", dtPwdLastSet, Now())
        
        If intDaysSinceChange > 180 Then
            Call AddFinding("CRITICAL", "KRBTGT Password Not Rotated", "KRBTGT password is " & intDaysSinceChange & " days old (" & Int(intDaysSinceChange/365) & " years)!" & vbCrLf & "RISK: Golden Ticket attack! Attacker with KRBTGT hash = permanent backdoor" & vbCrLf & "REMEDIATION: New-KrbtgtKeys.ps1 TWICE (wait 10h between resets)", 1)
            WScript.Echo "    CRITICAL: KRBTGT password is " & intDaysSinceChange & " days old!"
            WScript.Echo "    GOLDEN TICKET RISK!"
        Else
            WScript.Echo "    OK: KRBTGT password is " & intDaysSinceChange & " days old"
        End If
    End If
    
    On Error GoTo 0
End Sub

Sub CheckPasswordPolicyFast()
    WScript.Echo "  [1.5] Checking Password Policy..."
    
    On Error Resume Next
    
    Dim objDomain
    Set objDomain = GetObject("LDAP://" & strDomainDN)
    
    Dim intMinPwdLength, intMaxPwdAge
    intMinPwdLength = objDomain.Get("minPwdLength")
    intMaxPwdAge = objDomain.Get("maxPwdAge")
    
    Dim intMaxPwdAgeDays
    If IsNumeric(intMaxPwdAge) Then
        intMaxPwdAgeDays = Abs(intMaxPwdAge) / 864000000000
    Else
        intMaxPwdAgeDays = 0
    End If
    
    If intMinPwdLength < 14 Then
        Call AddFinding("CRITICAL", "Weak Password Policy", "Minimum password length is " & intMinPwdLength & " (recommend 14+)" & vbCrLf & "EXPLOITATION: Password spraying, brute force" & vbCrLf & "REMEDIATION: Set-ADDefaultDomainPasswordPolicy -MinPasswordLength 14", 1)
        WScript.Echo "    WARNING: Min password length: " & intMinPwdLength & " (recommend 14+)"
    Else
        WScript.Echo "    OK: Min password length: " & intMinPwdLength
    End If
    
    If intMaxPwdAgeDays = 0 Then
        Call AddFinding("MEDIUM", "Password Never Expires (Policy)", "Max password age is set to never expire" & vbCrLf & "REMEDIATION: Set-ADDefaultDomainPasswordPolicy -MaxPasswordAge 90", 1)
        WScript.Echo "    WARNING: Max password age: Never expires"
    ElseIf intMaxPwdAgeDays > 90 Then
        WScript.Echo "    INFO: Max password age: " & intMaxPwdAgeDays & " days (consider 90)"
    Else
        WScript.Echo "    OK: Max password age: " & intMaxPwdAgeDays & " days"
    End If
    
    On Error GoTo 0
End Sub

Sub CheckPasswordNeverExpiresFast()
    WScript.Echo "  [1.6] Checking Password Never Expires (sample)..."
    
    On Error Resume Next
    
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    
    Set objCommand.ActiveConnection = objConnection
    objCommand.Properties("Page Size") = 50 ' Small sample
    objCommand.Properties("Searchscope") = 2
    objCommand.Properties("Timeout") = 30
    
    ' DONT_EXPIRE_PASSWORD = 65536, ACCOUNTDISABLE = 2
    objCommand.CommandText = "SELECT sAMAccountName FROM 'LDAP://" & strDomainDN & "' WHERE objectCategory='person' AND userAccountControl:1.2.840.113556.1.4.803:=65536 AND NOT userAccountControl:1.2.840.113556.1.4.803:=2"
    
    Set objRecordSet = objCommand.Execute
    
    Dim intCount
    intCount = 0
    
    If Not objRecordSet.EOF Then
        objRecordSet.MoveFirst
        Do Until objRecordSet.EOF
            intCount = intCount + 1
            objRecordSet.MoveNext
        Loop
        
        If intCount > 0 Then
            Call AddFinding("HIGH", "Password Never Expires", "Found " & intCount & "+ accounts with password never expires (sample of 50)" & vbCrLf & "REMEDIATION: Set-ADUser -PasswordNeverExpires $false", intCount)
            WScript.Echo "    WARNING: Found " & intCount & "+ accounts with password never expires"
        End If
    End If
    
    If intCount = 0 Then
        WScript.Echo "    OK: No accounts with password never expires (in sample)"
    End If
    
    objRecordSet.Close
    objConnection.Close
    
    On Error GoTo 0
End Sub

Sub CheckMachineAccountQuotaFast()
    WScript.Echo "  [1.7] Checking MachineAccountQuota..."
    
    On Error Resume Next
    
    Dim objDomain, intMAQ
    Set objDomain = GetObject("LDAP://" & strDomainDN)
    intMAQ = objDomain.Get("ms-DS-MachineAccountQuota")
    
    If intMAQ > 0 Then
        Call AddFinding("HIGH", "MachineAccountQuota > 0", "MachineAccountQuota is " & intMAQ & " (ANY user can create computers!)" & vbCrLf & "EXPLOITATION: noPAC attack, RBCD abuse" & vbCrLf & "REMEDIATION: Set-ADDomain -Replace @{'ms-DS-MachineAccountQuota'='0'}", 1)
        WScript.Echo "    WARNING: MachineAccountQuota is " & intMAQ & " (should be 0!)"
    Else
        WScript.Echo "    OK: MachineAccountQuota is 0"
    End If
    
    On Error GoTo 0
End Sub

Sub CheckDomainAdminsFast()
    WScript.Echo "  [1.8] Checking Domain Admins..."
    
    On Error Resume Next
    
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    
    Set objCommand.ActiveConnection = objConnection
    objCommand.Properties("Searchscope") = 2
    objCommand.Properties("Timeout") = 30
    
    objCommand.CommandText = "SELECT member FROM 'LDAP://" & strDomainDN & "' WHERE cn='Domain Admins'"
    
    Set objRecordSet = objCommand.Execute
    
    Dim intCount
    intCount = 0
    
    If Not objRecordSet.EOF Then
        If Not IsNull(objRecordSet.Fields("member").Value) Then
            If IsArray(objRecordSet.Fields("member").Value) Then
                intCount = UBound(objRecordSet.Fields("member").Value) + 1
            Else
                intCount = 1
            End If
        End If
    End If
    
    If intCount > 5 Then
        Call AddFinding("HIGH", "Excessive Domain Admins", "Found " & intCount & " Domain Admin accounts (recommend 3-5)" & vbCrLf & "REMEDIATION: Follow least privilege, use tiered admin model", intCount)
        WScript.Echo "    WARNING: " & intCount & " Domain Admins (recommend max 5)"
    Else
        WScript.Echo "    OK: " & intCount & " Domain Admins"
    End If
    
    objRecordSet.Close
    objConnection.Close
    
    On Error GoTo 0
End Sub

Sub CheckLegacyOSFast()
    WScript.Echo "  [1.9] Checking Legacy Operating Systems (sample)..."
    
    On Error Resume Next
    
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    
    Set objCommand.ActiveConnection = objConnection
    objCommand.Properties("Page Size") = 50
    objCommand.Properties("Searchscope") = 2
    objCommand.Properties("Timeout") = 30
    
    ' Check for Windows 7, XP, 2003, 2008, 2012 (all unsupported)
    objCommand.CommandText = "SELECT sAMAccountName, operatingSystem FROM 'LDAP://" & strDomainDN & "' WHERE objectClass='computer' AND (operatingSystem='Windows 7*' OR operatingSystem='*Windows XP*' OR operatingSystem='*2003*' OR operatingSystem='*2008*' OR operatingSystem='*2012*')"
    
    Set objRecordSet = objCommand.Execute
    
    Dim intCount
    intCount = 0
    
    If Not objRecordSet.EOF Then
        objRecordSet.MoveFirst
        Do Until objRecordSet.EOF
            intCount = intCount + 1
            objRecordSet.MoveNext
        Loop
        
        If intCount > 0 Then
            Call AddFinding("CRITICAL", "Unsupported Operating Systems", "Found " & intCount & "+ computers with legacy/unsupported OS (sample)" & vbCrLf & "RISK: Vulnerable to known exploits (EternalBlue, etc.)" & vbCrLf & "REMEDIATION: Upgrade or isolate immediately", intCount)
            WScript.Echo "    WARNING: Found " & intCount & "+ legacy OS systems"
        End If
    End If
    
    If intCount = 0 Then
        WScript.Echo "    OK: No legacy OS found (in sample)"
    End If
    
    objRecordSet.Close
    objConnection.Close
    
    On Error GoTo 0
End Sub

Sub CheckDomainControllersFast()
    WScript.Echo "  [1.10] Checking Domain Controllers..."
    
    On Error Resume Next
    
    Dim strConfigDN
    strConfigDN = objRootDSE.Get("configurationNamingContext")
    
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    
    Set objCommand.ActiveConnection = objConnection
    objCommand.Properties("Searchscope") = 2
    objCommand.Properties("Timeout") = 30
    
    objCommand.CommandText = "SELECT dNSHostName FROM 'LDAP://CN=Sites," & strConfigDN & "' WHERE objectClass='server'"
    
    Set objRecordSet = objCommand.Execute
    
    Dim intCount
    intCount = 0
    
    If Not objRecordSet.EOF Then
        objRecordSet.MoveFirst
        Do Until objRecordSet.EOF
            intCount = intCount + 1
            objRecordSet.MoveNext
        Loop
    End If
    
    WScript.Echo "    INFO: Found " & intCount & " Domain Controllers"
    
    If intCount < 2 Then
        Call AddFinding("HIGH", "Single Point of Failure", "Only " & intCount & " Domain Controller detected" & vbCrLf & "RISK: No redundancy, single point of failure", intCount)
        WScript.Echo "    WARNING: Only 1 DC - consider adding redundant DCs"
    ElseIf intCount > 50 Then
        WScript.Echo "    INFO: Very large number of DCs (" & intCount & ")"
    End If
    
    objRecordSet.Close
    objConnection.Close
    
    On Error GoTo 0
End Sub

' ==================================================
' HELPER FUNCTIONS
' ==================================================

Sub AddFinding(strSeverity, strTitle, strDetails, intCount)
    Dim strKey
    strKey = dictFindings.Count + 1
    dictFindings.Add strKey, strSeverity & "|" & strTitle & "|" & strDetails & "|" & intCount
    
    Select Case UCase(strSeverity)
        Case "CRITICAL"
            intCritical = intCritical + 1
        Case "HIGH"
            intHigh = intHigh + 1
        Case "MEDIUM"
            intMedium = intMedium + 1
    End Select
End Sub

Sub DisplaySummary()
    WScript.Echo "FINDINGS SUMMARY:"
    WScript.Echo "  CRITICAL: " & intCritical
    WScript.Echo "  HIGH:     " & intHigh
    WScript.Echo "  MEDIUM:   " & intMedium
    WScript.Echo "  TOTAL:    " & dictFindings.Count
    WScript.Echo ""
    
    If intCritical > 0 Then
        WScript.Echo "CRITICAL FINDINGS - IMMEDIATE ACTION REQUIRED!"
        WScript.Echo ""
        
        Dim key, arrItem
        For Each key In dictFindings.Keys
            arrItem = Split(dictFindings(key), "|")
            If UCase(arrItem(0)) = "CRITICAL" Then
                WScript.Echo "  [!] " & arrItem(1)
                WScript.Echo "      Count: " & arrItem(3)
            End If
        Next
        WScript.Echo ""
    End If
End Sub

Sub GenerateQuickReport()
    strReportPath = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\AD_Quick_Audit_" & Replace(Replace(Replace(Now(), "/", "-"), ":", "-"), " ", "_") & ".html"
    
    Dim objFile, strHTML
    Set objFile = objFSO.CreateTextFile(strReportPath, True)
    
    strHTML = "<!DOCTYPE html><html><head><meta charset='UTF-8'><title>AD Quick Audit</title>"
    strHTML = strHTML & "<style>"
    strHTML = strHTML & "body{font-family:'Segoe UI',Arial;margin:20px;background:#f5f5f5}"
    strHTML = strHTML & ".header{background:linear-gradient(135deg,#667eea,#764ba2);color:white;padding:30px;border-radius:10px;margin-bottom:20px}"
    strHTML = strHTML & ".finding{background:white;padding:20px;margin:10px 0;border-radius:8px;border-left:5px solid #ddd}"
    strHTML = strHTML & ".critical{border-left-color:#e74c3c;background:#fff5f5}"
    strHTML = strHTML & ".high{border-left-color:#e67e22;background:#fff8f0}"
    strHTML = strHTML & ".medium{border-left-color:#f39c12;background:#fffbf0}"
    strHTML = strHTML & ".badge{display:inline-block;padding:5px 12px;border-radius:15px;color:white;font-size:12px;font-weight:bold;margin-right:10px}"
    strHTML = strHTML & ".badge-critical{background:#e74c3c}"
    strHTML = strHTML & ".badge-high{background:#e67e22}"
    strHTML = strHTML & ".badge-medium{background:#f39c12}"
    strHTML = strHTML & "h2{color:#2c3e50;border-bottom:2px solid #667eea;padding-bottom:10px}"
    strHTML = strHTML & ".summary{display:grid;grid-template-columns:repeat(3,1fr);gap:15px;margin:20px 0}"
    strHTML = strHTML & ".stat-box{background:white;padding:20px;border-radius:8px;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,0.1)}"
    strHTML = strHTML & ".stat-number{font-size:48px;font-weight:bold;margin:10px 0}"
    strHTML = strHTML & ".stat-label{color:#7f8c8d;font-size:14px}"
    strHTML = strHTML & ".critical .stat-number{color:#e74c3c}"
    strHTML = strHTML & ".high .stat-number{color:#e67e22}"
    strHTML = strHTML & ".medium .stat-number{color:#f39c12}"
    strHTML = strHTML & "</style></head><body>"
    
    strHTML = strHTML & "<div class='header'><h1>🚀 AD Quick Security Audit - Large Domain</h1>"
    strHTML = strHTML & "<p><strong>Domain:</strong> " & strDomainName & " | <strong>Generated:</strong> " & Now() & "</p>"
    strHTML = strHTML & "<p style='opacity:0.9'>Optimized scan for large domains (showing first " & MAX_RESULTS & " results per check)</p></div>"
    
    strHTML = strHTML & "<div class='summary'>"
    strHTML = strHTML & "<div class='stat-box critical'><div class='stat-number'>" & intCritical & "</div><div class='stat-label'>CRITICAL</div></div>"
    strHTML = strHTML & "<div class='stat-box high'><div class='stat-number'>" & intHigh & "</div><div class='stat-label'>HIGH</div></div>"
    strHTML = strHTML & "<div class='stat-box medium'><div class='stat-number'>" & intMedium & "</div><div class='stat-label'>MEDIUM</div></div>"
    strHTML = strHTML & "</div>"
    
    strHTML = strHTML & "<h2>🔴 Findings</h2>"
    
    Dim key, arrItem
    For Each key In dictFindings.Keys
        arrItem = Split(dictFindings(key), "|")
        Dim strSeverity, strTitle, strDetails, strCount
        strSeverity = LCase(arrItem(0))
        strTitle = arrItem(1)
        strDetails = Replace(arrItem(2), vbCrLf, "<br>")
        strCount = arrItem(3)
        
        strHTML = strHTML & "<div class='finding " & strSeverity & "'>"
        strHTML = strHTML & "<span class='badge badge-" & strSeverity & "'>" & UCase(arrItem(0)) & "</span>"
        strHTML = strHTML & "<strong>" & strTitle & "</strong>"
        If strCount <> "0" Then
            strHTML = strHTML & " <span style='color:#7f8c8d'>(Count: " & strCount & ")</span>"
        End If
        strHTML = strHTML & "<div style='margin-top:10px;color:#555'>" & strDetails & "</div>"
        strHTML = strHTML & "</div>"
    Next
    
    strHTML = strHTML & "<div style='margin-top:30px;padding:20px;background:#ecf0f1;border-radius:8px'>"
    strHTML = strHTML & "<h3>📌 Top Priority Actions</h3>"
    strHTML = strHTML & "<ol>"
    strHTML = strHTML & "<li>Remove all Unconstrained Delegation immediately</li>"
    strHTML = strHTML & "<li>Rotate KRBTGT if > 180 days old (TWICE!)</li>"
    strHTML = strHTML & "<li>Convert Kerberoastable accounts to gMSA</li>"
    strHTML = strHTML & "<li>Set MachineAccountQuota = 0</li>"
    strHTML = strHTML & "<li>Strengthen password policy (14+ chars)</li>"
    strHTML = strHTML & "</ol>"
    strHTML = strHTML & "<p style='margin-top:15px'><strong>For comprehensive audit use:</strong> Purple Knight, PingCastle, or PowerShell version</p>"
    strHTML = strHTML & "</div>"
    
    strHTML = strHTML & "</body></html>"
    
    objFile.Write strHTML
    objFile.Close
End Sub

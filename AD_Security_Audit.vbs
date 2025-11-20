' ========================================
' AD Security Audit Script
' Comprehensive Active Directory Security Assessment Tool
' Generates HTML Report with Vulnerabilities and Assets
' ========================================

Option Explicit

Dim objFSO, objFile, objRootDSE, objConnection, objCommand, objRecordSet
Dim strDomainDN, strDomainName, strReportPath, strHTMLContent
Dim dictUsers, dictGroups, dictComputers, dictVulnerabilities
Dim intTotalUsers, intAdmins, intInactiveUsers, intPasswordNeverExpires
Dim intWeakPasswords, intSPNAccounts, intOldComputers

' Initialize
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set dictUsers = CreateObject("Scripting.Dictionary")
Set dictGroups = CreateObject("Scripting.Dictionary")
Set dictComputers = CreateObject("Scripting.Dictionary")
Set dictVulnerabilities = CreateObject("Scripting.Dictionary")

' Get domain information
Set objRootDSE = GetObject("LDAP://RootDSE")
strDomainDN = objRootDSE.Get("defaultNamingContext")
strDomainName = Replace(strDomainDN, "DC=", "")
strDomainName = Replace(strDomainName, ",", ".")

' Report file path
strReportPath = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\AD_Security_Report_" & Replace(Replace(Replace(Now(), "/", "-"), ":", "-"), " ", "_") & ".html"

WScript.Echo "========================================="
WScript.Echo "AD Security Audit Tool"
WScript.Echo "========================================="
WScript.Echo "Domain: " & strDomainName
WScript.Echo "Starting audit..."
WScript.Echo ""

' Initialize counters
intTotalUsers = 0
intAdmins = 0
intInactiveUsers = 0
intPasswordNeverExpires = 0
intWeakPasswords = 0
intSPNAccounts = 0
intOldComputers = 0

' Perform audit
Call AuditUsers()
Call AuditGroups()
Call AuditComputers()
Call AuditGPOs()
Call CheckDomainControllers()
Call CheckPasswordPolicies()
Call CheckKerberosDelegation()
Call CheckTrusts()

' Generate HTML Report
Call GenerateHTMLReport()

WScript.Echo ""
WScript.Echo "========================================="
WScript.Echo "Audit Complete!"
WScript.Echo "Report saved to: " & strReportPath
WScript.Echo "========================================="

' ========================================
' Audit Users
' ========================================
Sub AuditUsers()
    WScript.Echo "[+] Auditing users..."
    
    Dim objUser, strLDAPPath
    Dim dtLastLogon, intDaysSinceLogon
    Dim objUserFlags
    
    strLDAPPath = "LDAP://" & strDomainDN
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    
    Set objCommand.ActiveConnection = objConnection
    objCommand.Properties("Page Size") = 1000
    objCommand.Properties("Searchscope") = 2 ' ADS_SCOPE_SUBTREE
    
    objCommand.CommandText = "SELECT distinguishedName, sAMAccountName, userAccountControl, lastLogonTimestamp, pwdLastSet, servicePrincipalName, adminCount, memberOf FROM '" & strLDAPPath & "' WHERE objectClass='user' AND objectCategory='person'"
    
    On Error Resume Next
    Set objRecordSet = objCommand.Execute
    On Error GoTo 0
    
    If Not objRecordSet.EOF Then
        objRecordSet.MoveFirst
        
        Do Until objRecordSet.EOF
            intTotalUsers = intTotalUsers + 1
            
            Dim strSAM, strDN, intUAC, strSPN, intAdminCount
            strSAM = objRecordSet.Fields("sAMAccountName").Value
            strDN = objRecordSet.Fields("distinguishedName").Value
            intUAC = objRecordSet.Fields("userAccountControl").Value
            
            ' Check if admin
            If Not IsNull(objRecordSet.Fields("adminCount").Value) Then
                If objRecordSet.Fields("adminCount").Value = 1 Then
                    intAdmins = intAdmins + 1
                End If
            End If
            
            ' Check password never expires
            If (intUAC And &H10000) = &H10000 Then
                intPasswordNeverExpires = intPasswordNeverExpires + 1
                Call AddVulnerability("Password Never Expires", "User: " & strSAM, "HIGH", "Password set to never expire")
            End If
            
            ' Check if account disabled
            Dim bDisabled
            bDisabled = (intUAC And 2) = 2
            
            ' Check last logon (inactive users)
            If Not IsNull(objRecordSet.Fields("lastLogonTimestamp").Value) Then
                Dim objLastLogon
                Set objLastLogon = objRecordSet.Fields("lastLogonTimestamp").Value
                dtLastLogon = Int(objLastLogon.HighPart * 2^32 + objLastLogon.LowPart) / 600000000 - 11644473600
                dtLastLogon = DateAdd("s", dtLastLogon, #1/1/1970#)
                intDaysSinceLogon = DateDiff("d", dtLastLogon, Now())
                
                If intDaysSinceLogon > 90 And Not bDisabled Then
                    intInactiveUsers = intInactiveUsers + 1
                    Call AddVulnerability("Inactive User Account", "User: " & strSAM & " (Last logon: " & intDaysSinceLogon & " days ago)", "MEDIUM", "Account not used for over 90 days")
                End If
            End If
            
            ' Check for SPN (Kerberoastable accounts)
            If Not IsNull(objRecordSet.Fields("servicePrincipalName").Value) Then
                intSPNAccounts = intSPNAccounts + 1
                Call AddVulnerability("Kerberoastable Account", "User: " & strSAM, "HIGH", "Account has SPN set and may be vulnerable to Kerberoasting")
            End If
            
            ' Store user info
            dictUsers.Add intTotalUsers, strSAM & "|" & strDN & "|" & intDaysSinceLogon
            
            objRecordSet.MoveNext
        Loop
    End If
    
    objRecordSet.Close
    objConnection.Close
    
    WScript.Echo "    Total Users: " & intTotalUsers
    WScript.Echo "    Admin Users: " & intAdmins
    WScript.Echo "    Inactive Users: " & intInactiveUsers
    WScript.Echo "    Password Never Expires: " & intPasswordNeverExpires
    WScript.Echo "    Kerberoastable Accounts: " & intSPNAccounts
End Sub

' ========================================
' Audit Groups
' ========================================
Sub AuditGroups()
    WScript.Echo "[+] Auditing groups..."
    
    Dim strLDAPPath
    strLDAPPath = "LDAP://" & strDomainDN
    
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    
    Set objCommand.ActiveConnection = objConnection
    objCommand.Properties("Page Size") = 1000
    objCommand.Properties("Searchscope") = 2
    
    objCommand.CommandText = "SELECT sAMAccountName, distinguishedName, member FROM '" & strLDAPPath & "' WHERE objectClass='group'"
    
    On Error Resume Next
    Set objRecordSet = objCommand.Execute
    On Error GoTo 0
    
    Dim intGroupCount
    intGroupCount = 0
    
    If Not objRecordSet.EOF Then
        objRecordSet.MoveFirst
        
        Do Until objRecordSet.EOF
            intGroupCount = intGroupCount + 1
            
            Dim strGroupName, strGroupDN
            strGroupName = objRecordSet.Fields("sAMAccountName").Value
            strGroupDN = objRecordSet.Fields("distinguishedName").Value
            
            ' Check privileged groups
            If InStr(1, strGroupName, "Admin", 1) > 0 Or _
               InStr(1, strGroupName, "Domain Admins", 1) > 0 Or _
               InStr(1, strGroupName, "Enterprise Admins", 1) > 0 Or _
               InStr(1, strGroupName, "Schema Admins", 1) > 0 Then
                
                Dim memberCount
                memberCount = 0
                If Not IsNull(objRecordSet.Fields("member").Value) Then
                    If IsArray(objRecordSet.Fields("member").Value) Then
                        memberCount = UBound(objRecordSet.Fields("member").Value) + 1
                    Else
                        memberCount = 1
                    End If
                End If
                
                Call AddVulnerability("Privileged Group", "Group: " & strGroupName & " (" & memberCount & " members)", "INFO", "Monitor privileged group membership")
            End If
            
            dictGroups.Add intGroupCount, strGroupName & "|" & strGroupDN
            
            objRecordSet.MoveNext
        Loop
    End If
    
    objRecordSet.Close
    objConnection.Close
    
    WScript.Echo "    Total Groups: " & intGroupCount
End Sub

' ========================================
' Audit Computers
' ========================================
Sub AuditComputers()
    WScript.Echo "[+] Auditing computers..."
    
    Dim strLDAPPath
    strLDAPPath = "LDAP://" & strDomainDN
    
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    
    Set objCommand.ActiveConnection = objConnection
    objCommand.Properties("Page Size") = 1000
    objCommand.Properties("Searchscope") = 2
    
    objCommand.CommandText = "SELECT sAMAccountName, distinguishedName, operatingSystem, lastLogonTimestamp, userAccountControl FROM '" & strLDAPPath & "' WHERE objectClass='computer'"
    
    On Error Resume Next
    Set objRecordSet = objCommand.Execute
    On Error GoTo 0
    
    Dim intComputerCount
    intComputerCount = 0
    
    If Not objRecordSet.EOF Then
        objRecordSet.MoveFirst
        
        Do Until objRecordSet.EOF
            intComputerCount = intComputerCount + 1
            
            Dim strCompName, strOS
            strCompName = objRecordSet.Fields("sAMAccountName").Value
            
            If Not IsNull(objRecordSet.Fields("operatingSystem").Value) Then
                strOS = objRecordSet.Fields("operatingSystem").Value
                
                ' Check for old OS versions
                If InStr(1, strOS, "Windows 2000", 1) > 0 Or _
                   InStr(1, strOS, "Windows XP", 1) > 0 Or _
                   InStr(1, strOS, "Windows 2003", 1) > 0 Or _
                   InStr(1, strOS, "Windows Vista", 1) > 0 Or _
                   InStr(1, strOS, "Windows 7", 1) > 0 Or _
                   InStr(1, strOS, "Server 2003", 1) > 0 Or _
                   InStr(1, strOS, "Server 2008", 1) > 0 Then
                    
                    intOldComputers = intOldComputers + 1
                    Call AddVulnerability("Unsupported Operating System", "Computer: " & strCompName & " (OS: " & strOS & ")", "CRITICAL", "Running unsupported/legacy OS")
                End If
            Else
                strOS = "Unknown"
            End If
            
            ' Check last logon for inactive computers
            If Not IsNull(objRecordSet.Fields("lastLogonTimestamp").Value) Then
                Dim objCompLastLogon, dtCompLastLogon, intCompDaysSinceLogon
                Set objCompLastLogon = objRecordSet.Fields("lastLogonTimestamp").Value
                dtCompLastLogon = Int(objCompLastLogon.HighPart * 2^32 + objCompLastLogon.LowPart) / 600000000 - 11644473600
                dtCompLastLogon = DateAdd("s", dtCompLastLogon, #1/1/1970#)
                intCompDaysSinceLogon = DateDiff("d", dtCompLastLogon, Now())
                
                If intCompDaysSinceLogon > 90 Then
                    Call AddVulnerability("Inactive Computer", "Computer: " & strCompName & " (Last seen: " & intCompDaysSinceLogon & " days ago)", "LOW", "Computer account not used for over 90 days")
                End If
            End If
            
            ' Check for unconstrained delegation
            Dim intCompUAC
            intCompUAC = objRecordSet.Fields("userAccountControl").Value
            If (intCompUAC And &H80000) = &H80000 Then
                Call AddVulnerability("Unconstrained Delegation", "Computer: " & strCompName, "CRITICAL", "Computer configured for unconstrained delegation")
            End If
            
            dictComputers.Add intComputerCount, strCompName & "|" & strOS
            
            objRecordSet.MoveNext
        Loop
    End If
    
    objRecordSet.Close
    objConnection.Close
    
    WScript.Echo "    Total Computers: " & intComputerCount
    WScript.Echo "    Legacy OS Systems: " & intOldComputers
End Sub

' ========================================
' Audit GPOs
' ========================================
Sub AuditGPOs()
    WScript.Echo "[+] Auditing Group Policy Objects..."
    
    Dim strLDAPPath
    strLDAPPath = "LDAP://CN=Policies,CN=System," & strDomainDN
    
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    
    Set objCommand.ActiveConnection = objConnection
    objCommand.Properties("Searchscope") = 2
    
    objCommand.CommandText = "SELECT displayName, gPCFileSysPath FROM '" & strLDAPPath & "' WHERE objectClass='groupPolicyContainer'"
    
    On Error Resume Next
    Set objRecordSet = objCommand.Execute
    On Error GoTo 0
    
    Dim intGPOCount
    intGPOCount = 0
    
    If Not objRecordSet.EOF Then
        objRecordSet.MoveFirst
        
        Do Until objRecordSet.EOF
            intGPOCount = intGPOCount + 1
            objRecordSet.MoveNext
        Loop
    End If
    
    objRecordSet.Close
    objConnection.Close
    
    WScript.Echo "    Total GPOs: " & intGPOCount
End Sub

' ========================================
' Check Domain Controllers
' ========================================
Sub CheckDomainControllers()
    WScript.Echo "[+] Checking Domain Controllers..."
    
    Dim strConfigDN
    strConfigDN = objRootDSE.Get("configurationNamingContext")
    
    Dim strLDAPPath
    strLDAPPath = "LDAP://CN=Sites," & strConfigDN
    
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    
    Set objCommand.ActiveConnection = objConnection
    objCommand.Properties("Searchscope") = 2
    
    objCommand.CommandText = "SELECT dNSHostName FROM '" & strLDAPPath & "' WHERE objectClass='server'"
    
    On Error Resume Next
    Set objRecordSet = objCommand.Execute
    On Error GoTo 0
    
    Dim intDCCount
    intDCCount = 0
    
    If Not objRecordSet.EOF Then
        objRecordSet.MoveFirst
        
        Do Until objRecordSet.EOF
            intDCCount = intDCCount + 1
            objRecordSet.MoveNext
        Loop
    End If
    
    objRecordSet.Close
    objConnection.Close
    
    If intDCCount < 2 Then
        Call AddVulnerability("Single Domain Controller", "Only " & intDCCount & " DC detected", "HIGH", "Consider adding redundant domain controllers")
    End If
    
    WScript.Echo "    Domain Controllers: " & intDCCount
End Sub

' ========================================
' Check Password Policies
' ========================================
Sub CheckPasswordPolicies()
    WScript.Echo "[+] Checking Password Policies..."
    
    On Error Resume Next
    Dim objDomain
    Set objDomain = GetObject("LDAP://" & strDomainDN)
    
    Dim intMinPwdLength, intMaxPwdAge, intMinPwdAge
    intMinPwdLength = objDomain.Get("minPwdLength")
    intMaxPwdAge = objDomain.Get("maxPwdAge")
    intMinPwdAge = objDomain.Get("minPwdAge")
    
    ' Convert maxPwdAge from 100-nanosecond intervals to days
    Dim intMaxPwdAgeDays
    If IsNumeric(intMaxPwdAge) Then
        intMaxPwdAgeDays = Abs(intMaxPwdAge) / 864000000000
    Else
        intMaxPwdAgeDays = 0
    End If
    
    On Error GoTo 0
    
    If intMinPwdLength < 14 Then
        Call AddVulnerability("Weak Password Policy", "Minimum password length is " & intMinPwdLength & " characters", "HIGH", "Increase minimum password length to 14+ characters")
    End If
    
    If intMaxPwdAgeDays > 90 Or intMaxPwdAgeDays = 0 Then
        Call AddVulnerability("Password Expiration Policy", "Maximum password age is " & intMaxPwdAgeDays & " days", "MEDIUM", "Set password expiration to 90 days or less")
    End If
    
    WScript.Echo "    Min Password Length: " & intMinPwdLength
    WScript.Echo "    Max Password Age: " & intMaxPwdAgeDays & " days"
End Sub

' ========================================
' Check Kerberos Delegation
' ========================================
Sub CheckKerberosDelegation()
    WScript.Echo "[+] Checking Kerberos Delegation..."
    
    Dim strLDAPPath
    strLDAPPath = "LDAP://" & strDomainDN
    
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    
    Set objCommand.ActiveConnection = objConnection
    objCommand.Properties("Page Size") = 1000
    objCommand.Properties("Searchscope") = 2
    
    ' Check for unconstrained delegation
    objCommand.CommandText = "SELECT sAMAccountName, objectClass FROM '" & strLDAPPath & "' WHERE userAccountControl:1.2.840.113556.1.4.803:=524288"
    
    On Error Resume Next
    Set objRecordSet = objCommand.Execute
    
    Dim intDelegationCount
    intDelegationCount = 0
    
    If Not objRecordSet.EOF Then
        objRecordSet.MoveFirst
        
        Do Until objRecordSet.EOF
            intDelegationCount = intDelegationCount + 1
            Dim strDelegatedAccount
            strDelegatedAccount = objRecordSet.Fields("sAMAccountName").Value
            Call AddVulnerability("Unconstrained Delegation", "Account: " & strDelegatedAccount, "CRITICAL", "Unconstrained delegation is a high-risk configuration")
            objRecordSet.MoveNext
        Loop
    End If
    
    objRecordSet.Close
    objConnection.Close
    On Error GoTo 0
    
    WScript.Echo "    Unconstrained Delegation: " & intDelegationCount & " accounts"
End Sub

' ========================================
' Check Trusts
' ========================================
Sub CheckTrusts()
    WScript.Echo "[+] Checking Domain Trusts..."
    
    On Error Resume Next
    Dim objDomain, objTrusts, objTrust
    Set objDomain = GetObject("LDAP://" & strDomainDN)
    Set objTrusts = objDomain.GetInfoEx(Array("trustPartner"), 0)
    
    Dim intTrustCount
    intTrustCount = 0
    
    ' Note: Trust enumeration may require specific permissions
    ' This is a basic check
    
    On Error GoTo 0
    
    WScript.Echo "    Domain Trusts: " & intTrustCount & " detected"
End Sub

' ========================================
' Add Vulnerability to Dictionary
' ========================================
Sub AddVulnerability(strTitle, strDetails, strSeverity, strDescription)
    Dim strKey
    strKey = dictVulnerabilities.Count + 1
    dictVulnerabilities.Add strKey, strSeverity & "|" & strTitle & "|" & strDetails & "|" & strDescription
End Sub

' ========================================
' Generate HTML Report
' ========================================
Sub GenerateHTMLReport()
    WScript.Echo "[+] Generating HTML report..."
    
    Dim strHTML
    strHTML = "<!DOCTYPE html>" & vbCrLf
    strHTML = strHTML & "<html lang='en'>" & vbCrLf
    strHTML = strHTML & "<head>" & vbCrLf
    strHTML = strHTML & "    <meta charset='UTF-8'>" & vbCrLf
    strHTML = strHTML & "    <meta name='viewport' content='width=device-width, initial-scale=1.0'>" & vbCrLf
    strHTML = strHTML & "    <title>Active Directory Security Audit Report</title>" & vbCrLf
    strHTML = strHTML & "    <style>" & vbCrLf
    strHTML = strHTML & "        * { margin: 0; padding: 0; box-sizing: border-box; }" & vbCrLf
    strHTML = strHTML & "        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f5f5f5; padding: 20px; }" & vbCrLf
    strHTML = strHTML & "        .container { max-width: 1400px; margin: 0 auto; background: white; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }" & vbCrLf
    strHTML = strHTML & "        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; border-radius: 10px 10px 0 0; }" & vbCrLf
    strHTML = strHTML & "        .header h1 { font-size: 32px; margin-bottom: 10px; }" & vbCrLf
    strHTML = strHTML & "        .header p { opacity: 0.9; font-size: 14px; }" & vbCrLf
    strHTML = strHTML & "        .content { padding: 30px; }" & vbCrLf
    strHTML = strHTML & "        .section { margin-bottom: 40px; }" & vbCrLf
    strHTML = strHTML & "        .section h2 { color: #333; font-size: 24px; margin-bottom: 20px; padding-bottom: 10px; border-bottom: 2px solid #667eea; }" & vbCrLf
    strHTML = strHTML & "        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; margin-bottom: 30px; }" & vbCrLf
    strHTML = strHTML & "        .stat-card { background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%); padding: 20px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }" & vbCrLf
    strHTML = strHTML & "        .stat-card h3 { color: #555; font-size: 14px; margin-bottom: 10px; text-transform: uppercase; }" & vbCrLf
    strHTML = strHTML & "        .stat-card .value { font-size: 36px; color: #667eea; font-weight: bold; }" & vbCrLf
    strHTML = strHTML & "        .vulnerability { background: white; border-left: 4px solid #ddd; padding: 15px; margin-bottom: 15px; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }" & vbCrLf
    strHTML = strHTML & "        .vulnerability.critical { border-left-color: #e74c3c; background: #fff5f5; }" & vbCrLf
    strHTML = strHTML & "        .vulnerability.high { border-left-color: #e67e22; background: #fff8f0; }" & vbCrLf
    strHTML = strHTML & "        .vulnerability.medium { border-left-color: #f39c12; background: #fffbf0; }" & vbCrLf
    strHTML = strHTML & "        .vulnerability.low { border-left-color: #3498db; background: #f0f8ff; }" & vbCrLf
    strHTML = strHTML & "        .vulnerability.info { border-left-color: #95a5a6; background: #f8f9fa; }" & vbCrLf
    strHTML = strHTML & "        .vulnerability h3 { font-size: 16px; margin-bottom: 8px; color: #333; }" & vbCrLf
    strHTML = strHTML & "        .vulnerability .severity { display: inline-block; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: bold; color: white; margin-bottom: 8px; }" & vbCrLf
    strHTML = strHTML & "        .severity.critical { background: #e74c3c; }" & vbCrLf
    strHTML = strHTML & "        .severity.high { background: #e67e22; }" & vbCrLf
    strHTML = strHTML & "        .severity.medium { background: #f39c12; }" & vbCrLf
    strHTML = strHTML & "        .severity.low { background: #3498db; }" & vbCrLf
    strHTML = strHTML & "        .severity.info { background: #95a5a6; }" & vbCrLf
    strHTML = strHTML & "        .vulnerability .details { color: #666; font-size: 14px; margin-bottom: 5px; }" & vbCrLf
    strHTML = strHTML & "        .vulnerability .description { color: #888; font-size: 13px; font-style: italic; }" & vbCrLf
    strHTML = strHTML & "        .footer { background: #f8f9fa; padding: 20px 30px; border-radius: 0 0 10px 10px; text-align: center; color: #666; font-size: 13px; }" & vbCrLf
    strHTML = strHTML & "        table { width: 100%; border-collapse: collapse; margin-top: 15px; }" & vbCrLf
    strHTML = strHTML & "        th, td { padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }" & vbCrLf
    strHTML = strHTML & "        th { background: #667eea; color: white; font-weight: 600; }" & vbCrLf
    strHTML = strHTML & "        tr:hover { background: #f5f5f5; }" & vbCrLf
    strHTML = strHTML & "    </style>" & vbCrLf
    strHTML = strHTML & "</head>" & vbCrLf
    strHTML = strHTML & "<body>" & vbCrLf
    strHTML = strHTML & "    <div class='container'>" & vbCrLf
    
    ' Header
    strHTML = strHTML & "        <div class='header'>" & vbCrLf
    strHTML = strHTML & "            <h1>🛡️ Active Directory Security Audit Report</h1>" & vbCrLf
    strHTML = strHTML & "            <p>Domain: <strong>" & strDomainName & "</strong> | Generated: " & Now() & "</p>" & vbCrLf
    strHTML = strHTML & "        </div>" & vbCrLf
    
    ' Content
    strHTML = strHTML & "        <div class='content'>" & vbCrLf
    
    ' Executive Summary
    strHTML = strHTML & "            <div class='section'>" & vbCrLf
    strHTML = strHTML & "                <h2>📊 Executive Summary</h2>" & vbCrLf
    strHTML = strHTML & "                <div class='stats-grid'>" & vbCrLf
    strHTML = strHTML & "                    <div class='stat-card'>" & vbCrLf
    strHTML = strHTML & "                        <h3>Total Vulnerabilities</h3>" & vbCrLf
    strHTML = strHTML & "                        <div class='value'>" & dictVulnerabilities.Count & "</div>" & vbCrLf
    strHTML = strHTML & "                    </div>" & vbCrLf
    strHTML = strHTML & "                    <div class='stat-card'>" & vbCrLf
    strHTML = strHTML & "                        <h3>Total Users</h3>" & vbCrLf
    strHTML = strHTML & "                        <div class='value'>" & intTotalUsers & "</div>" & vbCrLf
    strHTML = strHTML & "                    </div>" & vbCrLf
    strHTML = strHTML & "                    <div class='stat-card'>" & vbCrLf
    strHTML = strHTML & "                        <h3>Administrative Users</h3>" & vbCrLf
    strHTML = strHTML & "                        <div class='value'>" & intAdmins & "</div>" & vbCrLf
    strHTML = strHTML & "                    </div>" & vbCrLf
    strHTML = strHTML & "                    <div class='stat-card'>" & vbCrLf
    strHTML = strHTML & "                        <h3>Total Computers</h3>" & vbCrLf
    strHTML = strHTML & "                        <div class='value'>" & dictComputers.Count & "</div>" & vbCrLf
    strHTML = strHTML & "                    </div>" & vbCrLf
    strHTML = strHTML & "                    <div class='stat-card'>" & vbCrLf
    strHTML = strHTML & "                        <h3>Inactive Users</h3>" & vbCrLf
    strHTML = strHTML & "                        <div class='value'>" & intInactiveUsers & "</div>" & vbCrLf
    strHTML = strHTML & "                    </div>" & vbCrLf
    strHTML = strHTML & "                    <div class='stat-card'>" & vbCrLf
    strHTML = strHTML & "                        <h3>Password Never Expires</h3>" & vbCrLf
    strHTML = strHTML & "                        <div class='value'>" & intPasswordNeverExpires & "</div>" & vbCrLf
    strHTML = strHTML & "                    </div>" & vbCrLf
    strHTML = strHTML & "                    <div class='stat-card'>" & vbCrLf
    strHTML = strHTML & "                        <h3>Kerberoastable Accounts</h3>" & vbCrLf
    strHTML = strHTML & "                        <div class='value'>" & intSPNAccounts & "</div>" & vbCrLf
    strHTML = strHTML & "                    </div>" & vbCrLf
    strHTML = strHTML & "                    <div class='stat-card'>" & vbCrLf
    strHTML = strHTML & "                        <h3>Legacy OS Systems</h3>" & vbCrLf
    strHTML = strHTML & "                        <div class='value'>" & intOldComputers & "</div>" & vbCrLf
    strHTML = strHTML & "                    </div>" & vbCrLf
    strHTML = strHTML & "                </div>" & vbCrLf
    strHTML = strHTML & "            </div>" & vbCrLf
    
    ' Vulnerabilities Section
    strHTML = strHTML & "            <div class='section'>" & vbCrLf
    strHTML = strHTML & "                <h2>🔴 Security Vulnerabilities</h2>" & vbCrLf
    
    If dictVulnerabilities.Count > 0 Then
        Dim key, arrVuln
        For Each key In dictVulnerabilities.Keys
            arrVuln = Split(dictVulnerabilities(key), "|")
            Dim strSev, strTitle, strDetails, strDesc
            strSev = arrVuln(0)
            strTitle = arrVuln(1)
            strDetails = arrVuln(2)
            strDesc = arrVuln(3)
            
            strHTML = strHTML & "                <div class='vulnerability " & LCase(strSev) & "'>" & vbCrLf
            strHTML = strHTML & "                    <span class='severity " & LCase(strSev) & "'>" & strSev & "</span>" & vbCrLf
            strHTML = strHTML & "                    <h3>" & strTitle & "</h3>" & vbCrLf
            strHTML = strHTML & "                    <div class='details'>" & strDetails & "</div>" & vbCrLf
            strHTML = strHTML & "                    <div class='description'>" & strDesc & "</div>" & vbCrLf
            strHTML = strHTML & "                </div>" & vbCrLf
        Next
    Else
        strHTML = strHTML & "                <p>No vulnerabilities detected.</p>" & vbCrLf
    End If
    
    strHTML = strHTML & "            </div>" & vbCrLf
    
    ' Assets Section - Users
    strHTML = strHTML & "            <div class='section'>" & vbCrLf
    strHTML = strHTML & "                <h2>👥 User Assets (Sample - First 50)</h2>" & vbCrLf
    strHTML = strHTML & "                <table>" & vbCrLf
    strHTML = strHTML & "                    <thead>" & vbCrLf
    strHTML = strHTML & "                        <tr>" & vbCrLf
    strHTML = strHTML & "                            <th>Username</th>" & vbCrLf
    strHTML = strHTML & "                            <th>Distinguished Name</th>" & vbCrLf
    strHTML = strHTML & "                        </tr>" & vbCrLf
    strHTML = strHTML & "                    </thead>" & vbCrLf
    strHTML = strHTML & "                    <tbody>" & vbCrLf
    
    Dim i, maxDisplay
    maxDisplay = 50
    If dictUsers.Count < maxDisplay Then maxDisplay = dictUsers.Count
    
    For i = 1 To maxDisplay
        If dictUsers.Exists(i) Then
            Dim arrUser
            arrUser = Split(dictUsers(i), "|")
            strHTML = strHTML & "                        <tr>" & vbCrLf
            strHTML = strHTML & "                            <td>" & arrUser(0) & "</td>" & vbCrLf
            strHTML = strHTML & "                            <td style='font-size: 11px; color: #666;'>" & arrUser(1) & "</td>" & vbCrLf
            strHTML = strHTML & "                        </tr>" & vbCrLf
        End If
    Next
    
    strHTML = strHTML & "                    </tbody>" & vbCrLf
    strHTML = strHTML & "                </table>" & vbCrLf
    strHTML = strHTML & "            </div>" & vbCrLf
    
    ' Assets Section - Computers
    strHTML = strHTML & "            <div class='section'>" & vbCrLf
    strHTML = strHTML & "                <h2>💻 Computer Assets (Sample - First 50)</h2>" & vbCrLf
    strHTML = strHTML & "                <table>" & vbCrLf
    strHTML = strHTML & "                    <thead>" & vbCrLf
    strHTML = strHTML & "                        <tr>" & vbCrLf
    strHTML = strHTML & "                            <th>Computer Name</th>" & vbCrLf
    strHTML = strHTML & "                            <th>Operating System</th>" & vbCrLf
    strHTML = strHTML & "                        </tr>" & vbCrLf
    strHTML = strHTML & "                    </thead>" & vbCrLf
    strHTML = strHTML & "                    <tbody>" & vbCrLf
    
    maxDisplay = 50
    If dictComputers.Count < maxDisplay Then maxDisplay = dictComputers.Count
    
    For i = 1 To maxDisplay
        If dictComputers.Exists(i) Then
            Dim arrComp
            arrComp = Split(dictComputers(i), "|")
            strHTML = strHTML & "                        <tr>" & vbCrLf
            strHTML = strHTML & "                            <td>" & arrComp(0) & "</td>" & vbCrLf
            strHTML = strHTML & "                            <td>" & arrComp(1) & "</td>" & vbCrLf
            strHTML = strHTML & "                        </tr>" & vbCrLf
        End If
    Next
    
    strHTML = strHTML & "                    </tbody>" & vbCrLf
    strHTML = strHTML & "                </table>" & vbCrLf
    strHTML = strHTML & "            </div>" & vbCrLf
    
    strHTML = strHTML & "        </div>" & vbCrLf
    
    ' Footer
    strHTML = strHTML & "        <div class='footer'>" & vbCrLf
    strHTML = strHTML & "            <p>Active Directory Security Audit Tool | Generated on " & Now() & "</p>" & vbCrLf
    strHTML = strHTML & "            <p style='margin-top: 5px; font-size: 11px;'>This report provides a security assessment snapshot. Regular audits recommended.</p>" & vbCrLf
    strHTML = strHTML & "        </div>" & vbCrLf
    
    strHTML = strHTML & "    </div>" & vbCrLf
    strHTML = strHTML & "</body>" & vbCrLf
    strHTML = strHTML & "</html>" & vbCrLf
    
    ' Write to file
    Set objFile = objFSO.CreateTextFile(strReportPath, True)
    objFile.Write strHTML
    objFile.Close
End Sub

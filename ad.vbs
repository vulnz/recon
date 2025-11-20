Option Explicit

' ==========================
' Global objects / setup
' ==========================
Dim rootDSE, domainDN, configDN
Dim con, cmd, rs
Dim fso, reportsPath

Set fso = CreateObject("Scripting.FileSystemObject")

Dim scriptPath
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)
reportsPath = scriptPath & "\Reports"

If Not fso.FolderExists(reportsPath) Then
    fso.CreateFolder reportsPath
End If

WScript.Echo "========================================="
WScript.Echo "     ACTIVE DIRECTORY AUDIT (VBS)"
WScript.Echo "     Using current domain credentials"
WScript.Echo "     Reports folder: " & reportsPath
WScript.Echo "========================================="

Set rootDSE = GetObject("LDAP://RootDSE")
domainDN  = rootDSE.Get("defaultNamingContext")
configDN  = rootDSE.Get("configurationNamingContext")

Set con = CreateObject("ADODB.Connection")
Set cmd = CreateObject("ADODB.Command")

con.Provider = "ADsDSOObject"
con.Open "Active Directory Provider"
Set cmd.ActiveConnection = con

' ==========================
' Helpers
' ==========================
Sub WriteLine(fileName, text)
    Dim file
    Dim fullPath
    fullPath = reportsPath & "\" & fileName

    Set file = fso.OpenTextFile(fullPath, 8, True) ' 8 = Append
    file.WriteLine text
    file.Close
End Sub

Function GetRS(baseDN, filter, attrs, scope)
    If scope = "" Then scope = "subtree"
    cmd.CommandText = "<LDAP://" & baseDN & ">;" & filter & ";" & attrs & ";" & scope
    Set GetRS = cmd.Execute
End Function

Function SafeField(rs, fld)
    On Error Resume Next
    If rs.Fields(fld).Value = "" Then
        SafeField = ""
    Else
        SafeField = rs.Fields(fld).Value
    End If
    Err.Clear
End Function

Function MultiToString(val)
    Dim out, v
    out = ""
    If IsArray(val) Then
        For Each v In val
            If out <> "" Then out = out & "; "
            out = out & CStr(v)
        Next
    Else
        out = CStr(val)
    End If
    MultiToString = out
End Function

Function PingResolveIP(hostname)
    On Error Resume Next
    Dim colPing, objStatus
    PingResolveIP = ""
    Set colPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery( _
        "select * from Win32_PingStatus where Address='" & hostname & "'")
    For Each objStatus In colPing
        If Not IsNull(objStatus.ProtocolAddress) Then
            PingResolveIP = objStatus.ProtocolAddress
            Exit For
        End If
    Next
    Err.Clear
End Function

' ==========================
' 1) Domain info
' ==========================
Sub Audit_DomainInfo()
    Dim rsLocal
    Set rsLocal = GetRS(domainDN, "(& (objectClass=domainDNS))", _
        "name,whenCreated,whenChanged,lockoutThreshold,pwdHistoryLength,pwdMaxAge,pwdMinPwdLength", "base")

    WriteLine "AD_DomainInfo.txt", "=== DOMAIN INFO ==="
    If Not rsLocal.EOF Then
        WriteLine "AD_DomainInfo.txt", "Domain DN: " & domainDN
        WriteLine "AD_DomainInfo.txt", "Name: " & SafeField(rsLocal, "name")
        WriteLine "AD_DomainInfo.txt", "Created: " & SafeField(rsLocal, "whenCreated")
        WriteLine "AD_DomainInfo.txt", "Changed: " & SafeField(rsLocal, "whenChanged")
        WriteLine "AD_DomainInfo.txt", "LockoutThreshold: " & SafeField(rsLocal, "lockoutThreshold")
        WriteLine "AD_DomainInfo.txt", "PasswordHistoryLength: " & SafeField(rsLocal, "pwdHistoryLength")
        WriteLine "AD_DomainInfo.txt", "PasswordMaxAge (raw): " & SafeField(rsLocal, "pwdMaxAge")
        WriteLine "AD_DomainInfo.txt", "PasswordMinLength: " & SafeField(rsLocal, "pwdMinPwdLength")
    End If
    WScript.Echo "[+] Domain info collected."
End Sub

' ==========================
' 2) Domain Controllers
' ==========================
Sub Audit_DomainControllers()
    Dim rsLocal
    Set rsLocal = GetRS(domainDN, _
        "(&(objectClass=computer)(userAccountControl:1.2.840.113556.1.4.803:=8192))", _
        "name,dNSHostName,operatingSystem", "subtree")

    WriteLine "AD_DomainControllers.txt", "=== DOMAIN CONTROLLERS ==="
    Do Until rsLocal.EOF
        WriteLine "AD_DomainControllers.txt", _
            SafeField(rsLocal, "name") & " | " & _
            SafeField(rsLocal, "dNSHostName") & " | " & _
            SafeField(rsLocal, "operatingSystem")
        rsLocal.MoveNext
    Loop
    WScript.Echo "[+] Domain controllers listed."
End Sub

' ==========================
' 3) All Users
' ==========================
Sub Audit_AllUsers()
    Dim rsLocal
    Set rsLocal = GetRS(domainDN, _
        "(&(objectClass=user)(!(objectClass=computer)))", _
        "sAMAccountName,displayName,mail,userAccountControl", "subtree")

    WriteLine "AD_Users_All.txt", "=== ALL AD USERS ==="
    Dim count: count = 0
    Do Until rsLocal.EOF
        count = count + 1
        WriteLine "AD_Users_All.txt", _
            SafeField(rsLocal, "sAMAccountName") & " | " & _
            SafeField(rsLocal, "displayName") & " | " & _
            SafeField(rsLocal, "mail") & " | UAC=" & SafeField(rsLocal, "userAccountControl")
        rsLocal.MoveNext
    Loop
    WriteLine "AD_Users_All.txt", "TOTAL USERS: " & count
    WScript.Echo "[+] All users exported (" & count & ")."
End Sub

' ==========================
' 4) Test Accounts
' ==========================
Sub Audit_TestAccounts()
    Dim rsLocal
    ' crude pattern search
    Set rsLocal = GetRS(domainDN, _
        "(&(objectClass=user)(!(objectClass=computer))(|(sAMAccountName=*test*)(sAMAccountName=*qa*)(sAMAccountName=*lab*)(sAMAccountName=*demo*)(displayName=*test*)(displayName=*qa*)))", _
        "sAMAccountName,displayName,mail", "subtree")

    WriteLine "AD_Users_TestAccounts.txt", "=== POSSIBLE TEST ACCOUNTS ==="
    Do Until rsLocal.EOF
        WriteLine "AD_Users_TestAccounts.txt", _
            SafeField(rsLocal, "sAMAccountName") & " | " & _
            SafeField(rsLocal, "displayName") & " | " & _
            SafeField(rsLocal, "mail")
        rsLocal.MoveNext
    Loop
    WScript.Echo "[+] Test accounts list generated."
End Sub

' ==========================
' 5) Temp Accounts
' ==========================
Sub Audit_TempAccounts()
    Dim rsLocal
    Set rsLocal = GetRS(domainDN, _
        "(&(objectClass=user)(!(objectClass=computer))(|(sAMAccountName=*temp*)(sAMAccountName=*tmp*)(sAMAccountName=*guest*)(sAMAccountName=*trial*)(displayName=*temp*)(displayName=*guest*)))", _
        "sAMAccountName,displayName,mail", "subtree")

    WriteLine "AD_Users_TempAccounts.txt", "=== POSSIBLE TEMP / GUEST ACCOUNTS ==="
    Do Until rsLocal.EOF
        WriteLine "AD_Users_TempAccounts.txt", _
            SafeField(rsLocal, "sAMAccountName") & " | " & _
            SafeField(rsLocal, "displayName") & " | " & _
            SafeField(rsLocal, "mail")
        rsLocal.MoveNext
    Loop
    WScript.Echo "[+] Temp/guest accounts list generated."
End Sub

' ==========================
' 6) Disabled Accounts
' ==========================
Sub Audit_DisabledAccounts()
    Dim rsLocal
    ' userAccountControl bit 0x2 (ACCOUNTDISABLE)
    Set rsLocal = GetRS(domainDN, _
        "(&(objectClass=user)(!(objectClass=computer))(userAccountControl:1.2.840.113556.1.4.803:=2))", _
        "sAMAccountName,displayName,mail,userAccountControl", "subtree")

    WriteLine "AD_Users_Disabled.txt", "=== DISABLED USER ACCOUNTS ==="
    Do Until rsLocal.EOF
        WriteLine "AD_Users_Disabled.txt", _
            SafeField(rsLocal, "sAMAccountName") & " | " & _
            SafeField(rsLocal, "displayName") & " | " & _
            SafeField(rsLocal, "mail") & " | UAC=" & SafeField(rsLocal, "userAccountControl")
        rsLocal.MoveNext
    Loop
    WScript.Echo "[+] Disabled accounts listed."
End Sub

' ==========================
' 7) Password Never Expires
' ==========================
Sub Audit_PwdNeverExpires()
    Dim rsLocal
    ' userAccountControl bit 0x10000 (DONT_EXPIRE_PASSWORD)
    Set rsLocal = GetRS(domainDN, _
        "(&(objectClass=user)(!(objectClass=computer))(userAccountControl:1.2.840.113556.1.4.803:=65536))", _
        "sAMAccountName,displayName,mail,userAccountControl", "subtree")

    WriteLine "AD_Users_PwdNeverExpires.txt", "=== ACCOUNTS WITH PASSWORD NEVER EXPIRES ==="
    Do Until rsLocal.EOF
        WriteLine "AD_Users_PwdNeverExpires.txt", _
            SafeField(rsLocal, "sAMAccountName") & " | " & _
            SafeField(rsLocal, "displayName") & " | " & _
            SafeField(rsLocal, "mail") & " | UAC=" & SafeField(rsLocal, "userAccountControl")
        rsLocal.MoveNext
    Loop
    WScript.Echo "[+] Pwd-never-expires accounts listed."
End Sub

' ==========================
' 8) Service Accounts (SPNs)
' ==========================
Sub Audit_ServiceAccounts_SPN()
    Dim rsLocal
    Set rsLocal = GetRS(domainDN, _
        "(&(servicePrincipalName=*)(!(objectClass=computer)))", _
        "sAMAccountName,displayName,servicePrincipalName", "subtree")

    WriteLine "AD_Users_ServiceAccounts_SPNs.txt", "=== USERS WITH SPNs (SERVICE ACCOUNTS) ==="
    Do Until rsLocal.EOF
        WriteLine "AD_Users_ServiceAccounts_SPNs.txt", _
            SafeField(rsLocal, "sAMAccountName") & " | " & _
            SafeField(rsLocal, "displayName") & " | SPNs=" & _
            MultiToString(rsLocal.Fields("servicePrincipalName").Value)
        rsLocal.MoveNext
    Loop
    WScript.Echo "[+] Service accounts (SPNs) listed."
End Sub

' ==========================
' 9) Unconstrained Delegation
' ==========================
Sub Audit_Delegation_Unconstrained()
    Dim rsLocal
    ' TRUSTED_FOR_DELEGATION (0x80000) on users & computers
    Set rsLocal = GetRS(domainDN, _
        "(&(|(objectClass=user)(objectClass=computer))(userAccountControl:1.2.840.113556.1.4.803:=524288))", _
        "sAMAccountName,name,displayName,userAccountControl", "subtree")

    WriteLine "AD_Delegation_Unconstrained.txt", "=== ACCOUNTS WITH UNCONSTRAINED DELEGATION ==="
    Do Until rsLocal.EOF
        WriteLine "AD_Delegation_Unconstrained.txt", _
            SafeField(rsLocal, "name") & " | " & SafeField(rsLocal, "sAMAccountName") & " | UAC=" & SafeField(rsLocal, "userAccountControl")
        rsLocal.MoveNext
    Loop
    WScript.Echo "[+] Unconstrained delegation accounts listed."
End Sub

' ==========================
' 10) Constrained Delegation
' ==========================
Sub Audit_Delegation_Constrained()
    Dim rsLocal
    Set rsLocal = GetRS(domainDN, _
        "(&(|(objectClass=user)(objectClass=computer))(msDS-AllowedToDelegateTo=*))", _
        "sAMAccountName,name,msDS-AllowedToDelegateTo", "subtree")

    WriteLine "AD_Delegation_Constrained.txt", "=== ACCOUNTS WITH CONSTRAINED DELEGATION ==="
    Do Until rsLocal.EOF
        WriteLine "AD_Delegation_Constrained.txt", _
            SafeField(rsLocal, "name") & " | " & SafeField(rsLocal, "sAMAccountName") & _
            " | AllowedToDelegateTo=" & MultiToString(rsLocal.Fields("msDS-AllowedToDelegateTo").Value)
        rsLocal.MoveNext
    Loop
    WScript.Echo "[+] Constrained delegation accounts listed."
End Sub

' ==========================
' 11) All Groups
' ==========================
Sub Audit_Groups_All()
    Dim rsLocal
    Set rsLocal = GetRS(domainDN, _
        "(&(objectClass=group))", _
        "cn,sAMAccountName,member,description", "subtree")

    WriteLine "AD_Groups_All.txt", "=== ALL GROUPS ==="
    Do Until rsLocal.EOF
        Dim memberCount
        On Error Resume Next
        If IsArray(rsLocal.Fields("member").Value) Then
            memberCount = UBound(rsLocal.Fields("member").Value) + 1
        ElseIf rsLocal.Fields("member").Value <> "" Then
            memberCount = 1
        Else
            memberCount = 0
        End If
        Err.Clear

        WriteLine "AD_Groups_All.txt", _
            SafeField(rsLocal, "cn") & " | " & _
            SafeField(rsLocal, "sAMAccountName") & " | Members=" & memberCount & _
            " | Desc=" & SafeField(rsLocal, "description")
        rsLocal.MoveNext
    Loop
    WScript.Echo "[+] All groups exported."
End Sub

' ==========================
' 12) Empty Groups
' ==========================
Sub Audit_Groups_Empty()
    Dim rsLocal
    Set rsLocal = GetRS(domainDN, _
        "(&(objectClass=group)(!(member=*)))", _
        "cn,sAMAccountName,description", "subtree")

    WriteLine "AD_Groups_Empty.txt", "=== EMPTY GROUPS (NO MEMBERS) ==="
    Do Until rsLocal.EOF
        WriteLine "AD_Groups_Empty.txt", _
            SafeField(rsLocal, "cn") & " | " & _
            SafeField(rsLocal, "sAMAccountName") & " | Desc=" & SafeField(rsLocal, "description")
        rsLocal.MoveNext
    Loop
    WScript.Echo "[+] Empty groups listed."
End Sub

' ==========================
' 13) Large Groups (>50 members)
' ==========================
Sub Audit_Groups_Large()
    Dim rsLocal
    Set rsLocal = GetRS(domainDN, _
        "(&(objectClass=group)(member=*))", _
        "cn,sAMAccountName,member,description", "subtree")

    WriteLine "AD_Groups_Large.txt", "=== LARGE GROUPS (>=50 MEMBERS) ==="
    Do Until rsLocal.EOF
        Dim memberCount
        On Error Resume Next
        If IsArray(rsLocal.Fields("member").Value) Then
            memberCount = UBound(rsLocal.Fields("member").Value) + 1
        Else
            memberCount = 1
        End If
        Err.Clear

        If memberCount >= 50 Then
            WriteLine "AD_Groups_Large.txt", _
                SafeField(rsLocal, "cn") & " | " & _
                SafeField(rsLocal, "sAMAccountName") & " | Members=" & memberCount & _
                " | Desc=" & SafeField(rsLocal, "description")
        End If

        rsLocal.MoveNext
    Loop
    WScript.Echo "[+] Large groups (>=50 members) listed."
End Sub

' ==========================
' 14–16) Privileged Groups
' ==========================
Sub Audit_PrivGroupMembers(groupCN, outFile)
    Dim rsGroup, dn, rsMembers
    Set rsGroup = GetRS(domainDN, "(&(objectClass=group)(cn=" & groupCN & "))", "distinguishedName,member", "subtree")

    WriteLine outFile, "=== MEMBERS OF " & groupCN & " ==="
    If rsGroup.EOF Then Exit Sub

    Dim members, m
    members = rsGroup.Fields("member").Value
    If IsArray(members) Then
        For Each m In members
            ' m is DN, we can just print DN
            WriteLine outFile, m
        Next
    ElseIf members <> "" Then
        WriteLine outFile, members
    End If
End Sub

' ==========================
' 17) All Computers
' ==========================
Sub Audit_Computers_All()
    Dim rsLocal
    Set rsLocal = GetRS(domainDN, _
        "(&(objectClass=computer))", _
        "name,dNSHostName,operatingSystem,operatingSystemVersion", "subtree")

    WriteLine "AD_Computers_All.txt", "=== ALL COMPUTERS ==="
    Do Until rsLocal.EOF
        WriteLine "AD_Computers_All.txt", _
            SafeField(rsLocal, "name") & " | " & _
            SafeField(rsLocal, "dNSHostName") & " | " & _
            SafeField(rsLocal, "operatingSystem") & " | " & _
            SafeField(rsLocal, "operatingSystemVersion")
        rsLocal.MoveNext
    Loop
    WScript.Echo "[+] All computers exported."
End Sub

' ==========================
' 18) Old OS Computers
' ==========================
Sub Audit_Computers_OldOS()
    Dim rsLocal
    Set rsLocal = GetRS(domainDN, _
        "(&(objectClass=computer)(operatingSystem=*))", _
        "name,dNSHostName,operatingSystem", "subtree")

    WriteLine "AD_Computers_OldOS.txt", "=== COMPUTERS WITH POSSIBLY OLD OS ==="
    Do Until rsLocal.EOF
        Dim os
        os = LCase(SafeField(rsLocal, "operatingSystem"))
        If InStr(os, "windows 7") > 0 Or _
           InStr(os, "windows xp") > 0 Or _
           InStr(os, "2003") > 0 Or _
           InStr(os, "2008") > 0 Then
            WriteLine "AD_Computers_OldOS.txt", _
                SafeField(rsLocal, "name") & " | " & _
                SafeField(rsLocal, "dNSHostName") & " | " & _
                SafeField(rsLocal, "operatingSystem")
        End If
        rsLocal.MoveNext
    Loop
    WScript.Echo "[+] Old OS computers listed."
End Sub

' ==========================
' 19) DNS Zones (AD-integrated)
' ==========================
Sub Audit_DNS_Zones()
    On Error Resume Next
    Dim baseDNS, rsLocal
    baseDNS = "CN=MicrosoftDNS,CN=System," & domainDN

    Set rsLocal = GetRS(baseDNS, "(&(objectClass=dnsZone))", "name,distinguishedName", "subtree")

    If Err.Number <> 0 Then
        WriteLine "AD_DNS_Zones.txt", "Cannot access AD-integrated DNS at: " & baseDNS
        Err.Clear
        Exit Sub
    End If

    WriteLine "AD_DNS_Zones.txt", "=== AD-INTEGRATED DNS ZONES ==="
    Do Until rsLocal.EOF
        WriteLine "AD_DNS_Zones.txt", SafeField(rsLocal, "name") & " | " & SafeField(rsLocal, "distinguishedName")
        rsLocal.MoveNext
    Loop
    WScript.Echo "[+] DNS zones exported (if AD-integrated)."
End Sub

' ==========================
' 20) DNS Records (dnsNode)
' ==========================
Sub Audit_DNS_Records()
    On Error Resume Next
    Dim baseDNS, rsLocal
    baseDNS = "CN=MicrosoftDNS,CN=System," & domainDN

    Set rsLocal = GetRS(baseDNS, "(&(objectClass=dnsNode))", "name,distinguishedName", "subtree")

    If Err.Number <> 0 Then
        WriteLine "AD_DNS_Records.txt", "Cannot access AD-integrated DNS nodes."
        Err.Clear
        Exit Sub
    End If

    WriteLine "AD_DNS_Records.txt", "=== DNS NODES (HOSTNAMES) ==="
    Do Until rsLocal.EOF
        WriteLine "AD_DNS_Records.txt", SafeField(rsLocal, "name") & " | " & SafeField(rsLocal, "distinguishedName")
        rsLocal.MoveNext
    Loop
    WScript.Echo "[+] DNS records exported (may be large)."
End Sub

' ==========================
' 21) OUs
' ==========================
Sub Audit_OUs()
    Dim rsLocal
    Set rsLocal = GetRS(domainDN, "(&(objectClass=organizationalUnit))", "ou,distinguishedName,description", "subtree")

    WriteLine "AD_OUs.txt", "=== ORGANIZATIONAL UNITS ==="
    Do Until rsLocal.EOF
        WriteLine "AD_OUs.txt", _
            SafeField(rsLocal, "ou") & " | " & _
            SafeField(rsLocal, "distinguishedName") & " | Desc=" & SafeField(rsLocal, "description")
        rsLocal.MoveNext
    Loop
    WScript.Echo "[+] OUs listed."
End Sub

' ==========================
' 22) Project-like Objects (OUs + groups)
' ==========================
Sub Audit_Projects()
    Dim rsLocal
    ' search OUs and groups where name or description hints project/app/service/prod/dev/qa
    Set rsLocal = GetRS(domainDN, _
        "(|(&(objectClass=organizationalUnit)(|(ou=*project*)(ou=*prod*)(ou=*dev*)(ou=*app*)(ou=*service*)(ou=*qa*)(description=*project*)(description=*service*)))" & _
        "(&(objectClass=group)(|(cn=*project*)(cn=*prod*)(cn=*dev*)(cn=*app*)(cn=*service*)(cn=*qa*)(description=*project*)(description=*service*))))", _
        "cn,ou,distinguishedName,description,objectClass", "subtree")

    WriteLine "AD_Projects.txt", "=== POSSIBLE INTERNAL PROJECTS (OUs/GROUPS) ==="
    Do Until rsLocal.EOF
        WriteLine "AD_Projects.txt", _
            "ObjectClass=" & MultiToString(rsLocal.Fields("objectClass").Value) & " | " & _
            "CN/OU=" & SafeField(rsLocal, "cn") & SafeField(rsLocal, "ou") & " | " & _
            "DN=" & SafeField(rsLocal, "distinguishedName") & " | Desc=" & SafeField(rsLocal, "description")
        rsLocal.MoveNext
    Loop
    WScript.Echo "[+] Project-like OUs/groups listed."
End Sub

' ==========================
' 23) IP list (dNSHostName -> IP)
' ==========================
Sub Audit_IPs_FromDNS()
    Dim rsLocal
    Set rsLocal = GetRS(domainDN, _
        "(&(objectClass=computer)(dNSHostName=*))", _
        "name,dNSHostName", "subtree")

    WriteLine "AD_IPs_FromDNS.txt", "=== COMPUTER DNS TO IP RESOLUTION ==="
    Do Until rsLocal.EOF
        Dim host, ip
        host = SafeField(rsLocal, "dNSHostName")
        If host <> "" Then
            ip = PingResolveIP(host)
            WriteLine "AD_IPs_FromDNS.txt", _
                SafeField(rsLocal, "name") & " | " & host & " | IP=" & ip
        End If
        rsLocal.MoveNext
    Loop
    WScript.Echo "[+] IP list created via DNS/WMI ping."
End Sub

' ==========================
' MAIN
' ==========================
WScript.Echo vbCrLf & "[*] Starting AD audit..."

Audit_DomainInfo
Audit_DomainControllers
Audit_AllUsers
Audit_TestAccounts
Audit_TempAccounts
Audit_DisabledAccounts
Audit_PwdNeverExpires
Audit_ServiceAccounts_SPN
Audit_Delegation_Unconstrained
Audit_Delegation_Constrained
Audit_Groups_All
Audit_Groups_Empty
Audit_Groups_Large

Audit_PrivGroupMembers "Domain Admins", "AD_Group_DomainAdmins.txt"
Audit_PrivGroupMembers "Enterprise Admins", "AD_Group_EnterpriseAdmins.txt"
Audit_PrivGroupMembers "Schema Admins", "AD_Group_SchemaAdmins.txt"

Audit_Computers_All
Audit_Computers_OldOS
Audit_DNS_Zones
Audit_DNS_Records
Audit_OUs
Audit_Projects
Audit_IPs_FromDNS

WScript.Echo vbCrLf & "=== AD AUDIT COMPLETE ==="
WScript.Echo "Reports stored in: " & reportsPath

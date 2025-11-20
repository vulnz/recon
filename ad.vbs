Option Explicit

' ==========================
' Global objects / setup
' ==========================
Dim rootDSE, domainDN, configDN
Dim con, cmd
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

' -------- Get domain naming context safely --------
On Error Resume Next
Set rootDSE = GetObject("LDAP://RootDSE")
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Cannot access LDAP RootDSE. Are you joined to a domain?"
    WScript.Echo "Details: " & Err.Description
    WScript.Quit
End If
domainDN  = rootDSE.Get("defaultNamingContext")
configDN  = rootDSE.Get("configurationNamingContext")
Err.Clear

' -------- Open ADODB connection to AD using ADsDSOObject --------
Set con = CreateObject("ADODB.Connection")
Set cmd = CreateObject("ADODB.Command")

con.Open "Provider=ADsDSOObject;"
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Cannot open ADsDSOObject provider."
    WScript.Echo "Details: " & Err.Description
    WScript.Quit
End If
Err.Clear

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
    Set rsLocal = GetRS(domainDN, "(&(objectClass=domainDNS))", _
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
' 3) All Users (basic + flags)
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
' Test / temp / disabled / pwd-never-expires / etc.
' (same as before, omitted here for brevity)
' ==========================
' --- you can keep using the rest of the big script I gave,
' --- just replace everything ABOVE with this fixed header
' --- up to and including Audit_AllUsers()
' ==========================

' ===========================================================
' MAIN – call the audit functions you kept in the script
' ===========================================================
WScript.Echo vbCrLf & "[*] Starting AD audit..."

Audit_DomainInfo
Audit_DomainControllers
Audit_AllUsers

' call other Audit_* subs here (TestAccounts, TempAccounts, Groups, DNS, etc.)

WScript.Echo vbCrLf & "=== AD AUDIT COMPLETE ==="
WScript.Echo "Reports stored in: " & reportsPath

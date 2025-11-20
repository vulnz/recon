Option Explicit

' ==========================
' Global objects / setup
' ==========================
Dim rootDSE, domainDN
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
WScript.Echo "   WEAK CREDENTIALS AD AUDIT (READ-ONLY)"
WScript.Echo "   Using current domain credentials"
WScript.Echo "   Reports folder: " & reportsPath
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
    If IsNull(rs.Fields(fld).Value) Or rs.Fields(fld).Value = "" Then
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

' Convert AD 64-bit file time to VBScript Date
Function FileTimeToDate(ft)
    On Error Resume Next
    If IsNull(ft) Then
        FileTimeToDate = Null
        Exit Function
    End If
    If ft = 0 Then
        FileTimeToDate = Null
        Exit Function
    End If

    Dim seconds
    seconds = ft / 10000000 ' 10^7 100-ns intervals per second

    FileTimeToDate = DateAdd("s", seconds, #1/1/1601#)
    Err.Clear
End Function

' ==========================
' 1) Main weak-credential audit over all users
' ==========================
Sub Audit_Users_WeakCreds()
    Dim rsLocal
    Dim filter, attrs
    filter = "(&(objectClass=user)(!(objectClass=computer)))"
    attrs  = "sAMAccountName,displayName,mail,userAccountControl,pwdLastSet,lastLogonTimestamp,servicePrincipalName"

    Set rsLocal = GetRS(domainDN, filter, attrs, "subtree")

    WriteLine "Weak_PwdNeverExpires.txt",     "=== PASSWORD NEVER EXPIRES ==="
    WriteLine "Weak_PwdNotRequired.txt",     "=== PASSWORD NOT REQUIRED FLAG ==="
    WriteLine "Weak_ReversibleEncryption.txt","=== REVERSIBLE ENCRYPTION ALLOWED ==="
    WriteLine "Weak_PwdNeverSet.txt",        "=== PASSWORD NEVER SET (pwdLastSet=0) ==="
    WriteLine "Weak_PwdOlderThan365.txt",    "=== PASSWORD OLDER THAN 365 DAYS ==="
    WriteLine "Weak_StaleLogon90.txt",       "=== LAST LOGON OLDER THAN 90 DAYS ==="
    WriteLine "Weak_ServiceAccounts_SPN.txt","=== USERS WITH SPNs (SERVICE ACCOUNTS) ==="

    Dim nowDate: nowDate = Now

    Do Until rsLocal.EOF
        Dim sam, dn, disp, mail, uac
        Dim pwdLastSet, lastLogonTS
        Dim pwdDate, logonDate
        Dim line

        sam   = SafeField(rsLocal, "sAMAccountName")
        disp  = SafeField(rsLocal, "displayName")
        mail  = SafeField(rsLocal, "mail")
        uac   = 0
        If SafeField(rsLocal, "userAccountControl") <> "" Then
            uac = CLng(SafeField(rsLocal, "userAccountControl"))
        End If

        pwdLastSet   = SafeField(rsLocal, "pwdLastSet")
        lastLogonTS  = SafeField(rsLocal, "lastLogonTimestamp")

        pwdDate   = FileTimeToDate(pwdLastSet)
        logonDate = FileTimeToDate(lastLogonTS)

        line = sam & " | " & disp & " | " & mail & " | UAC=" & uac

        ' --- 1. Password never expires (0x10000) ---
        If (uac And 65536) <> 0 Then
            WriteLine "Weak_PwdNeverExpires.txt", line
        End If

        ' --- 2. Password not required (0x20) ---
        If (uac And 32) <> 0 Then
            WriteLine "Weak_PwdNotRequired.txt", line
        End If

        ' --- 3. Reversible encryption allowed (0x80) ---
        If (uac And 128) <> 0 Then
            WriteLine "Weak_ReversibleEncryption.txt", line
        End If

        ' --- 4. Password never set (pwdLastSet = 0) ---
        If (Not IsNull(pwdDate)) Then
            ' has some date
        Else
            ' Null means 0 or not set
            WriteLine "Weak_PwdNeverSet.txt", line & " | pwdLastSet=0"
        End If

        ' --- 5. Password older than 365 days ---
        If Not IsNull(pwdDate) Then
            If DateDiff("d", pwdDate, nowDate) >= 365 Then
                WriteLine "Weak_PwdOlderThan365.txt", _
                    line & " | pwdLastSet=" & CStr(pwdDate) & " | AgeDays=" & DateDiff("d", pwdDate, nowDate)
            End If
        End If

        ' --- 6. Last logon older than 90 days (stale account) ---
        If Not IsNull(logonDate) Then
            If DateDiff("d", logonDate, nowDate) >= 90 Then
                WriteLine "Weak_StaleLogon90.txt", _
                    line & " | lastLogon=" & CStr(logonDate) & " | AgeDays=" & DateDiff("d", logonDate, nowDate)
            End If
        End If

        ' --- 7. Service accounts (have SPNs) ---
        If Not IsNull(rsLocal.Fields("servicePrincipalName").Value) Then
            WriteLine "Weak_ServiceAccounts_SPN.txt", _
                line & " | SPNs=" & MultiToString(rsLocal.Fields("servicePrincipalName").Value)
        End If

        rsLocal.MoveNext
    Loop

    WScript.Echo "[+] Weak-credential user indicators exported."
End Sub

' ==========================
' 2) Privileged groups (Domain / Enterprise / Schema Admins)
' ==========================
Sub Audit_Privileged()
    Dim rsGroup
    Dim filter, attrs

    filter = "(&(objectClass=group)(|(cn=Domain Admins)(cn=Enterprise Admins)(cn=Schema Admins)))"
    attrs  = "cn,member"

    Set rsGroup = GetRS(domainDN, filter, attrs, "subtree")

    WriteLine "Weak_Privileged.txt", "=== PRIVILEGED GROUP MEMBERS (DA/EA/SA) ==="

    Do Until rsGroup.EOF
        Dim groupName, members, m
        groupName = SafeField(rsGroup, "cn")
        members   = rsGroup.Fields("member").Value

        If IsArray(members) Then
            For Each m In members
                WriteLine "Weak_Privileged.txt", groupName & " | " & m
            Next
        ElseIf Not IsNull(members) And members <> "" Then
            WriteLine "Weak_Privileged.txt", groupName & " | " & members
        End If

        rsGroup.MoveNext
    Loop

    WScript.Echo "[+] Privileged group memberships exported."
End Sub

' ==========================
' MAIN
' ==========================
WScript.Echo vbCrLf & "[*] Starting weak-credential AD audit..."

Audit_Users_WeakCreds
Audit_Privileged

WScript.Echo vbCrLf & "=== WEAK CREDENTIALS AUDIT COMPLETE ==="
WScript.Echo "Reports stored in: " & reportsPath

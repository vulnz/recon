Option Explicit

' ----------------------------------------
' INITIAL SETUP
' ----------------------------------------
Dim rootDSE, domainDN, conn, rs, q, fso, html
Set rootDSE = GetObject("LDAP://RootDSE")
domainDN = rootDSE.Get("defaultNamingContext")

Set conn = CreateObject("ADODB.Connection")
conn.Provider = "ADsDSOObject"
conn.Open "Active Directory Provider"

Set fso = CreateObject("Scripting.FileSystemObject")
Set html = fso.CreateTextFile("critical_findings.html", True)

' ----------------------------------------
' HTML HEADER
' ----------------------------------------
html.WriteLine "<html><head><title>CRITICAL Active Directory Findings</title>"
html.WriteLine "<style>body{font-family:Arial;} h2{color:#a00;} table{border-collapse:collapse;} td,th{border:1px solid #ccc;padding:4px;}</style>"
html.WriteLine "</head><body>"
html.WriteLine "<h1>🔴 CRITICAL Active Directory Findings</h1>"
html.WriteLine "<p>Domain: <b>" & domainDN & "</b></p>"
html.WriteLine "<hr>"

' ----------------------------------------
' CONSTANTS
' ----------------------------------------
Const UF_DONT_EXPIRE_PASSWD = 65536
Const UF_PASSWD_NOTREQD = 32
Const UF_ENCRYPTED_TEXT_PWD_ALLOWED = 128
Const UF_DONT_REQUIRE_PREAUTH = 4194304
Const UF_TRUSTED_FOR_DELEGATION = 524288

' ------------------------------------------------------
' 1) USERS WITH DO NOT REQUIRE PREAUTH (AS-REP ROAST)
' ------------------------------------------------------
html.WriteLine "<h2>1. 🔥 Users with DoNotRequirePreauth (AS-REP Roast)</h2>"
html.WriteLine "<table><tr><th>sAMAccountName</th><th>DN</th><th>UAC</th></tr>"

q = "<LDAP://" & domainDN & ">;" & _
    "(&(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=4194304));" & _
    "sAMAccountName,distinguishedName,userAccountControl;subtree"

Set rs = conn.Execute(q)
Do Until rs.EOF
    html.WriteLine "<tr><td>" & rs("sAMAccountName").Value & "</td><td>" & rs("distinguishedName").Value & "</td><td>" & rs("userAccountControl").Value & "</td></tr>"
    rs.MoveNext
Loop
html.WriteLine "</table><hr>"

' ------------------------------------------------------
' 2) UNCONSTRAINED DELEGATION (CRITICAL)
' ------------------------------------------------------
html.WriteLine "<h2>2. 🔥 Unconstrained Delegation Accounts</h2>"
html.WriteLine "<table><tr><th>Name</th><th>DN</th><th>UAC</th></tr>"

q = "<LDAP://" & domainDN & ">;" & _
    "(&(userAccountControl:1.2.840.113556.1.4.803:=524288));" & _
    "sAMAccountName,distinguishedName,userAccountControl;subtree"

Set rs = conn.Execute(q)
Do Until rs.EOF
    html.WriteLine "<tr><td>" & rs("sAMAccountName").Value & "</td><td>" & rs("distinguishedName").Value & "</td><td>" & rs("userAccountControl").Value & "</td></tr>"
    rs.MoveNext
Loop
html.WriteLine "</table><hr>"

' ------------------------------------------------------
' 3) COMPUTERS WITH UNCONSTRAINED DELEGATION
' ------------------------------------------------------
html.WriteLine "<h2>3. 🔥 Machines with Unconstrained Delegation</h2>"
html.WriteLine "<table><tr><th>Name</th><th>DN</th><th>UAC</th></tr>"

q = "<LDAP://" & domainDN & ">;" & _
    "(&(objectClass=computer)(userAccountControl:1.2.840.113556.1.4.803:=524288));" & _
    "sAMAccountName,distinguishedName,userAccountControl;subtree"

Set rs = conn.Execute(q)
Do Until rs.EOF
    html.WriteLine "<tr><td>" & rs("sAMAccountName").Value & "</td><td>" & rs("distinguishedName").Value & "</td><td>" & rs("userAccountControl").Value & "</td></tr>"
    rs.MoveNext
Loop
html.WriteLine "</table><hr>"

' ------------------------------------------------------
' 4) REVERSIBLE PASSWORD ENCRYPTION (CRITICAL)
' ------------------------------------------------------
html.WriteLine "<h2>4. 🔥 Reversible Password Encryption Enabled</h2>"
html.WriteLine "<table><tr><th>Name</th><th>DN</th><th>UAC</th></tr>"

q = "<LDAP://" & domainDN & ">;" & _
    "(&(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=128));" & _
    "sAMAccountName,distinguishedName,userAccountControl;subtree"

Set rs = conn.Execute(q)
Do Until rs.EOF
    html.WriteLine "<tr><td>" & rs("sAMAccountName").Value & "</td><td>" & rs("distinguishedName").Value & "</td><td>" & rs("userAccountControl").Value & "</td></tr>"
    rs.MoveNext
Loop
html.WriteLine "</table><hr>"

' ------------------------------------------------------
' 5) PASSWORD NEVER EXPIRES + PRIVILEGE = CRITICAL
' ------------------------------------------------------
html.WriteLine "<h2>5. 🔥 Privileged Accounts with Password Never Expires</h2>"
html.WriteLine "<table><tr><th>Name</th><th>DN</th><th>UAC</th><th>AdminCount</th></tr>"

q = "<LDAP://" & domainDN & ">;" & _
    "(&(objectClass=user)(adminCount=1));" & _
    "sAMAccountName,distinguishedName,userAccountControl,adminCount;subtree"

Set rs = conn.Execute(q)
Do Until rs.EOF
    Dim uac
    uac = rs("userAccountControl").Value

    If (uac And UF_DONT_EXPIRE_PASSWD) <> 0 Then
        html.WriteLine "<tr><td>" & rs("sAMAccountName").Value & "</td><td>" & rs("distinguishedName").Value & "</td><td>" & uac & "</td><td>" & rs("adminCount").Value & "</td></tr>"
    End If

    rs.MoveNext
Loop

html.WriteLine "</table><hr>"

' ------------------------------------------------------
' 6) DUPLICATE SPN (REAL-WORLD HIGH RISK ISSUE)
' ------------------------------------------------------
html.WriteLine "<h2>6. 🔥 Duplicate SPN Entries (Massively Critical)</h2>"
html.WriteLine "<p>Duplicate SPNs break Kerberos and enable credential relay attacks.</p>"
html.WriteLine "<table><tr><th>SPN</th><th>Account</th></tr>"

Dim dict, spn, arr, dn

Set dict = CreateObject("Scripting.Dictionary")

q = "<LDAP://" & domainDN & ">;" & _
    "(servicePrincipalName=*);" & _
    "distinguishedName,servicePrincipalName;subtree"

Set rs = conn.Execute(q)
Do Until rs.EOF
    dn = rs("distinguishedName").Value

    If Not IsNull(rs("servicePrincipalName").Value) Then
        arr = rs("servicePrincipalName").Value
        If IsArray(arr) Then
            Dim s
            For Each s In arr
                If dict.Exists(LCase(s)) Then
                    html.WriteLine "<tr><td>" & s & "</td><td>" & dn & "</td></tr>"
                    html.WriteLine "<tr><td>" & s & "</td><td>" & dict(LCase(s)) & "</td></tr>"
                Else
                    dict.Add LCase(s), dn
                End If
            Next
        End If
    End If

    rs.MoveNext
Loop

html.WriteLine "</table><hr>"
html.WriteLine "<h3>Audit Completed</h3>"

html.WriteLine "</body></html>"
html.Close

WScript.Echo "Report generated: critical_findings.html"

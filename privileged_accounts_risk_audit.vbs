Option Explicit

' ----------------------------------------
' SETUP
' ----------------------------------------
Dim rootDSE, domainDN, fso, csvPriv, csvSvc, csvAC, html, conn, rs, q
Set rootDSE = GetObject("LDAP://RootDSE")
domainDN = rootDSE.Get("defaultNamingContext")

Set fso = CreateObject("Scripting.FileSystemObject")
Set csvPriv = fso.CreateTextFile("privileged_accounts.csv", True)
Set csvSvc  = fso.CreateTextFile("privileged_accounts_service.csv", True)
Set csvAC   = fso.CreateTextFile("privileged_accounts_admincount.csv", True)
Set html    = fso.CreateTextFile("privileged_accounts_risk.html", True)

csvPriv.WriteLine "AccountType,sAMAccountName,distinguishedName,userAccountControl,AdminCount"
csvSvc.WriteLine  "ServiceAccount,sAMAccountName,distinguishedName,userAccountControl"
csvAC.WriteLine   "AdminCount1,sAMAccountName,distinguishedName"

' ----------------------------------------
' HTML HEADER
' ----------------------------------------
html.WriteLine "<html><head><title>Privileged Account Risk Audit</title></head><body>"
html.WriteLine "<h2>Privileged Account Risk Audit - " & domainDN & "</h2>"
html.WriteLine "<table border='1' cellspacing='0' cellpadding='4'>"
html.WriteLine "<tr><th>User</th><th>Risk</th><th>DN</th></tr>"

' ----------------------------------------
' LDAP CONNECT
' ----------------------------------------
Set conn = CreateObject("ADODB.Connection")
conn.Provider = "ADsDSOObject"
conn.Open "Active Directory Provider"

' ----------------------------------------
' QUERY PRIVILEGED ACCOUNTS
' ----------------------------------------
q = "<LDAP://" & domainDN & ">;" & _
    "(&(objectCategory=person)(objectClass=user));" & _
    "distinguishedName,sAMAccountName,userAccountControl,adminCount;subtree"

Set rs = conn.Execute(q)

' ----------------------------------------
' CONSTANTS
' ----------------------------------------
Const UF_DONT_EXPIRE_PASSWD = 65536
Const UF_PASSWD_NOTREQD     = 32
Const UF_ENCRYPTED_TEXT_PWD_ALLOWED = 128
Const UF_ACCOUNTDISABLE = 2

' ----------------------------------------
' PROCESS ACCOUNTS
' ----------------------------------------
Do Until rs.EOF
    Dim sam, dn, uac, ac, risk

    sam = rs("sAMAccountName").Value
    dn  = rs("distinguishedName").Value

    uac = 0 : ac = 0
    If Not IsNull(rs("userAccountControl").Value) Then uac = rs("userAccountControl").Value
    If Not IsNull(rs("adminCount").Value) Then ac = rs("adminCount").Value

    ' ------------------------------------------------
    ' PRIVILEGED GROUPS CHECK
    ' ------------------------------------------------
    If InStr(1, LCase(dn), "domain admins") > 0 Or _
       InStr(1, LCase(dn), "enterprise admins") > 0 Or _
       InStr(1, LCase(dn), "schema admins") > 0 Or _
       InStr(1, LCase(dn), "account operators") > 0 Or _
       ac = 1 Then

        csvPriv.WriteLine "Privileged," & sam & ",""" & dn & """," & uac & "," & ac

        ' ------------------------------------------------
        ' RISK FLAGS
        ' ------------------------------------------------
        risk = ""

        If (uac And UF_DONT_EXPIRE_PASSWD) <> 0 Then
            risk = risk & "PasswordNeverExpires; "
        End If

        If (uac And UF_PASSWD_NOTREQD) <> 0 Then
            risk = risk & "NoPasswordRequired; "
        End If

        If (uac And UF_ENCRYPTED_TEXT_PWD_ALLOWED) <> 0 Then
            risk = risk & "ReversiblePassword; "
        End If

        If (uac And UF_ACCOUNTDISABLE) <> 0 Then
            risk = risk & "DisabledPrivUser; "
        End If

        If InStr(LCase(sam), "svc") > 0 Or InStr(LCase(sam), "service") > 0 Then
            csvSvc.WriteLine "Service," & sam & ",""" & dn & """," & uac
            risk = risk & "ServiceAccount; "
        End If

        html.WriteLine "<tr><td>" & sam & "</td><td>" & risk & "</td><td>" & dn & "</td></tr>"
    End If

    ' ------------------------------------------------
    ' adminCount=1 (protected admin)
    ' ------------------------------------------------
    If ac = 1 Then
        csvAC.WriteLine sam & ",""" & dn & """"
    End If

    rs.MoveNext
Loop

' ----------------------------------------
' HTML FOOTER
' ----------------------------------------
html.WriteLine "</table></body></html>"
html.Close

csvPriv.Close
csvSvc.Close
csvAC.Close
rs.Close
conn.Close

WScript.Echo "Generated:"
WScript.Echo " - privileged_accounts.csv"
WScript.Echo " - privileged_accounts_service.csv"
WScript.Echo " - privileged_accounts_admincount.csv"
WScript.Echo " - privileged_accounts_risk.html"

Option Explicit
Const DONT_EXPIRE    = 65536
Const PASSWD_NOTREQD = 32

Dim rootDSE, domainDN, conn, rs, q, fso, out, uac, sam

Set rootDSE = GetObject("LDAP://RootDSE")
domainDN = rootDSE.Get("defaultNamingContext")

Set conn = CreateObject("ADODB.Connection")
conn.Provider = "ADsDSOObject"
conn.Open "Active Directory Provider"

q = "<LDAP://" & domainDN & ">;" & _
    "(&(objectCategory=person)(objectClass=user)(!(objectClass=computer)));" & _
    "distinguishedName,sAMAccountName,userAccountControl;subtree"

Set rs = conn.Execute(q)

Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile("ad_audit09_service_accounts_risky.csv", True)
out.WriteLine "sAMAccountName,distinguishedName,userAccountControl,PasswordNeverExpires,NoPasswordRequired"

Do Until rs.EOF
    sam = LCase(rs("sAMAccountName").Value)
    If InStr(sam, "svc") > 0 Or InStr(sam, "service") > 0 Then
        uac = 0
        If Not IsNull(rs("userAccountControl").Value) Then uac = rs("userAccountControl").Value
        Dim pne, npr
        pne = IIf((uac And DONT_EXPIRE) <> 0, "YES", "")
        npr = IIf((uac And PASSWD_NOTREQD) <> 0, "YES", "")

        If pne <> "" Or npr <> "" Then
            out.WriteLine rs("sAMAccountName").Value & ",""" & rs("distinguishedName").Value & """," & uac & "," & pne & "," & npr
        End If
    End If
    rs.MoveNext
Loop

rs.Close : conn.Close : out.Close
WScript.Echo "Report: ad_audit09_service_accounts_risky.csv"

Option Explicit
Const ENCRYPTED_TEXT_PWD_ALLOWED = 128

Dim rootDSE, domainDN, conn, rs, q, fso, out, uac

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
Set out = fso.CreateTextFile("ad_audit04_reversible_encryption.csv", True)
out.WriteLine "sAMAccountName,distinguishedName,userAccountControl"

Do Until rs.EOF
    uac = 0
    If Not IsNull(rs("userAccountControl").Value) Then uac = rs("userAccountControl").Value
    If (uac And ENCRYPTED_TEXT_PWD_ALLOWED) <> 0 Then
        out.WriteLine rs("sAMAccountName").Value & ",""" & rs("distinguishedName").Value & """," & uac
    End If
    rs.MoveNext
Loop

rs.Close : conn.Close : out.Close
WScript.Echo "Report: ad_audit04_reversible_encryption.csv"

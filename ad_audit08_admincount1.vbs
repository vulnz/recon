Option Explicit

Dim rootDSE, domainDN, conn, rs, q, fso, out

Set rootDSE = GetObject("LDAP://RootDSE")
domainDN = rootDSE.Get("defaultNamingContext")

Set conn = CreateObject("ADODB.Connection")
conn.Provider = "ADsDSOObject"
conn.Open "Active Directory Provider"

q = "<LDAP://" & domainDN & ">;" & _
    "(&(objectCategory=person)(objectClass=user)(adminCount=1));" & _
    "distinguishedName,sAMAccountName;subtree"

Set rs = conn.Execute(q)

Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile("ad_audit08_admincount1.csv", True)
out.WriteLine "sAMAccountName,distinguishedName"

Do Until rs.EOF
    out.WriteLine rs("sAMAccountName").Value & ",""" & rs("distinguishedName").Value & """"
    rs.MoveNext
Loop

rs.Close : conn.Close : out.Close
WScript.Echo "Report: ad_audit08_admincount1.csv"

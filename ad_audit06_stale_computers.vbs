Option Explicit

Const STALE_DAYS = 90

Dim rootDSE, domainDN, conn, rs, q, fso, out
Dim nowFileTime, thresholdFileTime

Set rootDSE = GetObject("LDAP://RootDSE")
domainDN = rootDSE.Get("defaultNamingContext")

Set conn = CreateObject("ADODB.Connection")
conn.Provider = "ADsDSOObject"
conn.Open "Active Directory Provider"

q = "<LDAP://" & domainDN & ">;" & _
    "(&(objectClass=computer));" & _
    "distinguishedName,sAMAccountName,lastLogonTimestamp;subtree"

Set rs = conn.Execute(q)

Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile("ad_audit06_stale_computers.csv", True)
out.WriteLine "sAMAccountName,distinguishedName,lastLogonTimestamp,ApproxLastLogon"

nowFileTime = NowToFileTime(Now)
thresholdFileTime = nowFileTime - (STALE_DAYS * 864000000000#)

Do Until rs.EOF
    Dim llt, approx
    llt = 0
    If Not IsNull(rs("lastLogonTimestamp").Value) Then llt = CDbl(rs("lastLogonTimestamp").Value)

    If llt > 0 And llt < thresholdFileTime Then
        approx = FileTimeToDate(llt)
        out.WriteLine rs("sAMAccountName").Value & ",""" & rs("distinguishedName").Value & """," & llt & ",""" & approx & """"
    End If

    rs.MoveNext
Loop

rs.Close : conn.Close : out.Close
WScript.Echo "Report: ad_audit06_stale_computers.csv"

Function NowToFileTime(d)
    NowToFileTime = (CDbl(d) * 864000000000#) + 621355968000000000#
End Function

Function FileTimeToDate(ft)
    FileTimeToDate = CDate((ft - 621355968000000000#) / 864000000000#)
End Function

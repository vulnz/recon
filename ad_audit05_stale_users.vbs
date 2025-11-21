Option Explicit

Dim rootDSE, domainDN, conn, rs, q, fso, out
Dim nowFileTime, thresholdFileTime

' days to consider stale
Const STALE_DAYS = 90

Set rootDSE = GetObject("LDAP://RootDSE")
domainDN = rootDSE.Get("defaultNamingContext")

Set conn = CreateObject("ADODB.Connection")
conn.Provider = "ADsDSOObject"
conn.Open "Active Directory Provider"

q = "<LDAP://" & domainDN & ">;" & _
    "(&(objectCategory=person)(objectClass=user)(!(objectClass=computer)));" & _
    "distinguishedName,sAMAccountName,lastLogonTimestamp;subtree"

Set rs = conn.Execute(q)

Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile("ad_audit05_stale_users.csv", True)
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
WScript.Echo "Report: ad_audit05_stale_users.csv"

' --- helper: convert VB Date to FILETIME (approx) ---
Function NowToFileTime(d)
    NowToFileTime = (CDbl(d) * 864000000000#) + 621355968000000000#
End Function

' --- helper: FILETIME -> Date ---
Function FileTimeToDate(ft)
    FileTimeToDate = CDate((ft - 621355968000000000#) / 864000000000#)
End Function

Option Explicit

Dim criticalGroups
criticalGroups = Array( _
  "CN=Domain Admins,CN=Users,", _
  "CN=Enterprise Admins,CN=Users,", _
  "CN=Schema Admins,CN=Users,", _
  "CN=Administrators,CN=Builtin,", _
  "CN=Account Operators,CN=Builtin,", _
  "CN=Backup Operators,CN=Builtin," _
)

Dim rootDSE, domainDN, fso, out, baseDN, groupDN, grp, mem

Set rootDSE = GetObject("LDAP://RootDSE")
domainDN = rootDSE.Get("defaultNamingContext")

Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile("ad_audit07_critical_groups.csv", True)
out.WriteLine "GroupDN,MemberDN,MemberSAM"

For Each baseDN In criticalGroups
    groupDN = baseDN & domainDN
    On Error Resume Next
    Set grp = GetObject("LDAP://" & groupDN)
    If Err.Number = 0 Then
        Dim mems
        If Not IsNull(grp.Get("member")) Then
            mems = grp.GetEx("member")
            For Each mem In mems
                Dim memObj, sam
                sam = ""
                On Error Resume Next
                Set memObj = GetObject("LDAP://" & mem)
                If Err.Number = 0 Then
                    If Not IsNull(memObj.Get("sAMAccountName")) Then sam = memObj.Get("sAMAccountName")
                End If
                Err.Clear
                out.WriteLine """" & groupDN & """,""" & mem & """,""" & sam & """"
            Next
        End If
    End If
    Err.Clear
    On Error GoTo 0
Next

out.Close
WScript.Echo "Report: ad_audit07_critical_groups.csv"

Option Explicit
Dim net, domain, fso, out, grpName, grp, m
Set net = CreateObject("WScript.Network")
domain = net.UserDomain

Dim groups
groups = Array("Domain Admins", "Enterprise Admins", "Schema Admins")

Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile("audit03_privileged_groups.csv", True)
out.WriteLine "Domain,Group,Member"

For Each grpName In groups
    On Error Resume Next
    Set grp = GetObject("WinNT://" & domain & "/" & grpName & ",group")
    If Err.Number = 0 Then
        For Each m In grp.Members
            out.WriteLine domain & "," & """" & grpName & """,""" & m.Name & """"
        Next
    End If
    Err.Clear
    On Error GoTo 0
Next

out.Close
WScript.Echo "Report: audit03_privileged_groups.csv"

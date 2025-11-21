Option Explicit
Const PASSWD_NOTREQD = 32

Dim net, domain, fso, out, d, u
Set net = CreateObject("WScript.Network")
domain = net.UserDomain

Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile("audit02_no_password_required.csv", True)
out.WriteLine "Domain,User,Flags"

Set d = GetObject("WinNT://" & domain)
d.Filter = Array("User")

For Each u In d
    If (u.Flags And PASSWD_NOTREQD) <> 0 Then
        out.WriteLine domain & "," & u.Name & "," & u.Flags
    End If
Next

out.Close
WScript.Echo "Report: audit02_no_password_required.csv"

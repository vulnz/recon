Option Explicit
Const DONT_EXPIRE = 65536

Dim net, domain, fso, out, d, u
Set net = CreateObject("WScript.Network")
domain = net.UserDomain

Set fso  = CreateObject("Scripting.FileSystemObject")
Set out  = fso.CreateTextFile("audit01_pw_never_expires.csv", True)
out.WriteLine "Domain,User,Flags"

Set d = GetObject("WinNT://" & domain)
d.Filter = Array("User")

For Each u In d
    If (u.Flags And DONT_EXPIRE) <> 0 Then
        out.WriteLine domain & "," & u.Name & "," & u.Flags
    End If
Next

out.Close
WScript.Echo "Report: audit01_pw_never_expires.csv"

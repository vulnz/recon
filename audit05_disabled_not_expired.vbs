Option Explicit
Const DISABLED = 2

Dim net, domain, fso, out, d, u
Set net = CreateObject("WScript.Network")
domain = net.UserDomain

Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile("audit05_disabled_not_expired.csv", True)
out.WriteLine "Domain,User,Flags"

Set d = GetObject("WinNT://" & domain)
d.Filter = Array("User")

For Each u In d
    On Error Resume Next
    If (u.Flags And DISABLED) <> 0 Then
        ' No expiration property here, treat all disabled as needing review
        out.WriteLine domain & "," & u.Name & "," & u.Flags
    End If
    Err.Clear
    On Error GoTo 0
Next

out.Close
WScript.Echo "Report: audit05_disabled_not_expired.csv"

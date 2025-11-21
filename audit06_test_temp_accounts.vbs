Option Explicit
Dim net, domain, fso, out, d, u, n, ln
Set net = CreateObject("WScript.Network")
domain = net.UserDomain

Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile("audit06_test_temp_accounts.csv", True)
out.WriteLine "Domain,User"

Set d = GetObject("WinNT://" & domain)
d.Filter = Array("User")

For Each u In d
    n = LCase(u.Name)
    If InStr(n, "test") > 0 Or InStr(n, "temp") > 0 Or InStr(n, "guest") > 0 Then
        out.WriteLine domain & "," & u.Name
    End If
Next

out.Close
WScript.Echo "Report: audit06_test_temp_accounts.csv"

Option Explicit
Dim net, domain, fso, out, d, u, threshold
Set net = CreateObject("WScript.Network")
domain = net.UserDomain

threshold = Now - 90   ' 90 days

Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile("audit04_stale_users.csv", True)
out.WriteLine "Domain,User,LastLogin"

Set d = GetObject("WinNT://" & domain)
d.Filter = Array("User")

For Each u In d
    On Error Resume Next
    If IsDate(u.LastLogin) Then
        If CDate(u.LastLogin) < threshold Then
            out.WriteLine domain & "," & u.Name & ",""" & u.LastLogin & """"
        End If
    End If
    Err.Clear
    On Error GoTo 0
Next

out.Close
WScript.Echo "Report: audit04_stale_users.csv"

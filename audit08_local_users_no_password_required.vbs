Option Explicit
Const PASSWD_NOTREQD = 32

Dim net, computer, fso, out, d, u
Set net = CreateObject("WScript.Network")
computer = net.ComputerName

Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile("audit08_local_users_no_password_required.csv", True)
out.WriteLine "Computer,User,Flags"

Set d = GetObject("WinNT://" & computer)
d.Filter = Array("User")

For Each u In d
    If (u.Flags And PASSWD_NOTREQD) <> 0 Then
        out.WriteLine computer & "," & u.Name & "," & u.Flags
    End If
Next

out.Close
WScript.Echo "Report: audit08_local_users_no_password_required.csv"

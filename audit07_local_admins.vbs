Option Explicit
Dim net, computer, fso, out, grp, m
Set net = CreateObject("WScript.Network")
computer = net.ComputerName

Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile("audit07_local_admins.csv", True)
out.WriteLine "Computer,Member"

On Error Resume Next
Set grp = GetObject("WinNT://" & computer & "/Administrators,group")
If Err.Number = 0 Then
    For Each m In grp.Members
        out.WriteLine computer & "," & m.Name
    Next
End If
Err.Clear
On Error GoTo 0

out.Close
WScript.Echo "Report: audit07_local_admins.csv"

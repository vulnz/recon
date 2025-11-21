Option Explicit
Dim shell, fso, out, tmp, net, computer, txt

Set net    = CreateObject("WScript.Network")
computer   = net.ComputerName
Set shell  = CreateObject("WScript.Shell")
Set fso    = CreateObject("Scripting.FileSystemObject")

tmp = fso.GetSpecialFolder(2) & "\rdp_" & computer & ".txt"
shell.Run "reg query ""HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp"" /v UserAuthentication > """ & tmp & """", 0, True

Set out = fso.CreateTextFile("audit11_rdp_nla.csv", True)
out.WriteLine "Computer,NLA_Enabled"

If fso.FileExists(tmp) Then
    Dim f
    Set f = fso.OpenTextFile(tmp, 1)
    txt = f.ReadAll
    f.Close
    If InStr(LCase(txt), "0x0") > 0 Then
        out.WriteLine computer & ",NO"
    ElseIf InStr(LCase(txt), "0x1") > 0 Then
        out.WriteLine computer & ",YES"
    Else
        out.WriteLine computer & ",UNKNOWN"
    End If
    fso.DeleteFile tmp
Else
    out.WriteLine computer & ",UNKNOWN"
End If

out.Close
WScript.Echo "Report: audit11_rdp_nla.csv"

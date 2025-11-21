Option Explicit
Dim shell, fso, out, tmp, net, computer, f, line

Set net    = CreateObject("WScript.Network")
computer   = net.ComputerName
Set shell  = CreateObject("WScript.Shell")
Set fso    = CreateObject("Scripting.FileSystemObject")

tmp = fso.GetSpecialFolder(2) & "\fw_" & computer & ".txt"
shell.Run "cmd /c netsh advfirewall show allprofiles > """ & tmp & """", 0, True

Set out = fso.CreateTextFile("audit12_firewall_status.csv", True)
out.WriteLine "Computer,Profile,State"

If fso.FileExists(tmp) Then
    Set f = fso.OpenTextFile(tmp, 1)
    Dim currentProfile
    currentProfile = ""
    Do Until f.AtEndOfStream
        line = Trim(f.ReadLine)
        If InStr(line, "Domain Profile Settings") > 0 Then currentProfile = "Domain"
        If InStr(line, "Private Profile Settings") > 0 Then currentProfile = "Private"
        If InStr(line, "Public Profile Settings") > 0 Then currentProfile = "Public"

        If LCase(Left(line, 11)) = "state      " Then
            out.WriteLine computer & "," & currentProfile & "," & Trim(Mid(line, 12))
        End If
    Loop
    f.Close
    fso.DeleteFile tmp
Else
    out.WriteLine computer & ",Unknown,Unknown"
End If

out.Close
WScript.Echo "Report: audit12_firewall_status.csv"

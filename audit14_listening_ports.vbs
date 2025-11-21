Option Explicit
Dim shell, fso, out, tmp, net, computer, f, line, parts

Set net    = CreateObject("WScript.Network")
computer   = net.ComputerName
Set shell  = CreateObject("WScript.Shell")
Set fso    = CreateObject("Scripting.FileSystemObject")

tmp = fso.GetSpecialFolder(2) & "\netstat_" & computer & ".txt"
shell.Run "cmd /c netstat -ano -p tcp > """ & tmp & """", 0, True

Set out = fso.CreateTextFile("audit14_listening_ports.csv", True)
out.WriteLine "Computer,LocalAddress,LocalPort,State,PID"

If fso.FileExists(tmp) Then
    Set f = fso.OpenTextFile(tmp, 1)
    Do Until f.AtEndOfStream
        line = Trim(f.ReadLine)
        If LCase(Left(line, 3)) = "tcp" Then
            parts = Split(line)
            If UBound(parts) >= 4 Then
                Dim local, state, pid, hostPort
                local = parts(1)
                state = parts(3)
                pid   = parts(4)
                If InStr(local, ":") > 0 Then
                    hostPort = Split(local, ":")
                    out.WriteLine computer & "," & hostPort(0) & "," & hostPort(1) & "," & state & "," & pid
                End If
            End If
        End If
    Loop
    f.Close
    fso.DeleteFile tmp
End If

out.Close
WScript.Echo "Report: audit14_listening_ports.csv"

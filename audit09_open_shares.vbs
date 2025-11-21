Option Explicit
Dim shell, fso, out, tmp, f, line, net, computer

Set net   = CreateObject("WScript.Network")
computer  = net.ComputerName
Set shell = CreateObject("WScript.Shell")
Set fso   = CreateObject("Scripting.FileSystemObject")

tmp = fso.GetSpecialFolder(2) & "\shares_" & computer & ".txt"
shell.Run "cmd /c net share > """ & tmp & """", 0, True

Set out = fso.CreateTextFile("audit09_open_shares.csv", True)
out.WriteLine "Computer,Line"

Set f = fso.OpenTextFile(tmp, 1)
Do Until f.AtEndOfStream
    line = f.ReadLine
    If InStr(LCase(line), "everyone") > 0 Then
        out.WriteLine computer & ",""" & Replace(line, """", "'") & """"
    End If
Loop
f.Close

out.Close
fso.DeleteFile tmp

WScript.Echo "Report: audit09_open_shares.csv"

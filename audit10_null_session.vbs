Option Explicit
Dim shell, net, computer, fso, out, rc

Set net    = CreateObject("WScript.Network")
computer   = net.ComputerName
Set shell  = CreateObject("WScript.Shell")
Set fso    = CreateObject("Scripting.FileSystemObject")

Set out = fso.CreateTextFile("audit10_null_session.csv", True)
out.WriteLine "Computer,NullSessionAllowed"

rc = shell.Run("cmd /c net use \\" & computer & "\ipc$ """" /user:"""":", 0, True)

If rc = 0 Then
    out.WriteLine computer & ",YES"
    ' Clean up mapping
    shell.Run "cmd /c net use \\" & computer & "\ipc$ /delete /y", 0, True
Else
    out.WriteLine computer & ",NO"
End If

out.Close
WScript.Echo "Report: audit10_null_session.csv"

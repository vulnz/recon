Option Explicit
Dim fso, out, shell, tmp, net, computer, f, line

Set net    = CreateObject("WScript.Network")
computer   = net.ComputerName
Set shell  = CreateObject("WScript.Shell")
Set fso    = CreateObject("Scripting.FileSystemObject")

tmp = fso.GetSpecialFolder(2) & "\services_" & computer & ".txt"
' sc qc dumps service configuration, we filter later
shell.Run "cmd /c sc qc type= service > """ & tmp & """", 0, True

Set out = fso.CreateTextFile("audit13_unquoted_service_paths.csv", True)
out.WriteLine "Computer,Service,ImagePath"

If fso.FileExists(tmp) Then
    Set f = fso.OpenTextFile(tmp, 1)
    Dim svcName, img
    svcName = ""
    Do Until f.AtEndOfStream
        line = Trim(f.ReadLine)
        If Left(line, 8) = "SERVICE_" Then
            svcName = Mid(line, InStr(line, "NAME:") + 5)
            svcName = Trim(svcName)
        ElseIf InStr(LCase(line), "binary_path_name") > 0 Then
            img = Trim(Mid(line, InStr(line, ":") + 1))
            ' unquoted with space
            If InStr(img, " ") > 0 And Left(img, 1) <> """" Then
                out.WriteLine computer & ",""" & svcName & """,""" & Replace(img, """", "'") & """"
            End If
        End If
    Loop
    f.Close
    fso.DeleteFile tmp
End If

out.Close
WScript.Echo "Report: audit13_unquoted_service_paths.csv"

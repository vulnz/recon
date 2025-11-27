Option Explicit

Dim ports, hosts, i, j, port, host, sock

ports = Array(389, 636, 3268, 3269, 999)
hosts = Array("172.31.198.3", "172.31.214.1")

For i = 0 To UBound(hosts)
    host = hosts(i)
    WScript.Echo "=== Проверка хоста: " & host & " ==="
    
    For j = 0 To UBound(ports)
        port = ports(j)
        On Error Resume Next
        Set sock = CreateObject("WinHttp.WinHttpRequest.5.1")
        sock.SetTimeouts 2000, 2000, 2000, 2000
        sock.Open "POST", "http://" & host & ":" & port, False
        sock.Send

        If Err.Number = 0 Then
            WScript.Echo "  [✓] Порт " & port & " ОТКРЫТ"
        Else
            WScript.Echo "  [✗] Порт " & port & " ЗАКРЫТ или не отвечает"
        End If

        Set sock = Nothing
        Err.Clear
    Next

    WScript.Echo ""
Next

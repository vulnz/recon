' ldap_check.vbs — Проверка LDAP TCP соединения через порт 999
Option Explicit

Const LDAP_HOST = "127.0.0.1"
Const LDAP_PORT = 999
Const LDAP_BIND_REQUEST = _
    Chr(&H30) & Chr(&H0C) & _              ' LDAPMessage (SEQUENCE)
    Chr(&H02) & Chr(&H01) & Chr(&H01) & _   ' Message ID: 1
    Chr(&H60) & Chr(&H07) & _               ' BindRequest (application 0)
    Chr(&H02) & Chr(&H01) & Chr(&H03) & _   ' LDAP v3
    Chr(&H04) & Chr(&H00) & _               ' name (empty)
    Chr(&H80) & Chr(&H00)                   ' authentication (simple, empty)

Dim sock, connected, response

Set sock = CreateObject("MSWinsock.Winsock")
sock.RemoteHost = LDAP_HOST
sock.RemotePort = LDAP_PORT

WScript.Echo "[*] Connecting to " & LDAP_HOST & ":" & LDAP_PORT & "..."
sock.Connect

' Wait for connection
connected = False
Dim i
For i = 0 To 50 ' wait max 5 seconds
    If sock.Connect = 7 Then
        connected = True
        Exit For
    End If
    WScript.Sleep 100
Next

If Not connected Then
    WScript.Echo "[!] Connection failed."
    WScript.Quit(1)
End If

WScript.Echo "[+] Connected. Sending LDAP bind request..."
sock.SendData LDAP_BIND_REQUEST

' Wait for response
WScript.Sleep 500

If sock.BytesReceived > 0 Then
    sock.GetData response
    WScript.Echo "[✓] Got response from LDAP server!"
    WScript.Echo "[>] Raw response (hex): " & ToHex(response)
Else
    WScript.Echo "[✗] No response received from server."
End If

sock.Close

Function ToHex(str)
    Dim i, hex, c
    For i = 1 To Len(str)
        c = AscB(MidB(str, i, 1))
        hex = hex & Right("0" & Hex(c), 2) & " "
    Next
    ToHex = hex
End Function

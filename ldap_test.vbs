Option Explicit

Const HOST = "127.0.0.1"
Const PORT = 999
Const BASE_DN = "DC=corp,DC=local" ' Можешь не менять, просто placeholder

Dim sock, connected, response, data, i, line
Dim foundUsers, foundDCs, userCount, dcCount

Set sock = CreateObject("MSWinsock.Winsock")
sock.RemoteHost = HOST
sock.RemotePort = PORT

WScript.Echo "[*] Connecting to " & HOST & ":" & PORT & "..."
sock.Connect

connected = False
For i = 0 To 50
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

WScript.Echo "[+] Connected. Sending BIND request..."

Dim BIND_REQ
BIND_REQ = Chr(&H30) & Chr(&H0C) & _
           Chr(&H02) & Chr(&H01) & Chr(&H01) & _
           Chr(&H60) & Chr(&H07) & _
           Chr(&H02) & Chr(&H01) & Chr(&H03) & _
           Chr(&H04) & Chr(&H00) & _
           Chr(&H80) & Chr(&H00)

sock.SendData BIND_REQ
WScript.Sleep 500

If sock.BytesReceived > 0 Then
    sock.GetData response
    WScript.Echo "[✓] BIND successful."
Else
    WScript.Echo "[✗] No BIND response. Exiting."
    WScript.Quit(1)
End If

WScript.Echo "[*] Sending SearchRequest for users..."

Dim SEARCH_REQ
SEARCH_REQ = Chr(&H30) & Chr(&H3F) & _
             Chr(&H02) & Chr(&H01) & Chr(&H02) & _
             Chr(&H63) & Chr(&H3A) & _
             Chr(&H04) & Chr(Len(BASE_DN)) & BASE_DN & _
             Chr(&H0A) & Chr(&H01) & Chr(&H02) & _
             Chr(&H0A) & Chr(&H01) & Chr(&H00) & _
             Chr(&H02) & Chr(&H01) & Chr(&H00) & _
             Chr(&H02) & Chr(&H01) & Chr(&H00) & _
             Chr(&H01) & Chr(&H01) & Chr(&H00) & _
             Chr(&H87) & Chr(&H0F) & "(objectClass=user)" & _
             Chr(&H30) & Chr(&H00)

sock.SendData SEARCH_REQ
WScript.Sleep 1500

If sock.BytesReceived = 0 Then
    WScript.Echo "[✗] No SearchResult received."
    sock.Close
    WScript.Quit(1)
End If

sock.GetData data
WScript.Echo "[✓] SearchResult received. Parsing..."

Dim lines
lines = Split(data, vbCrLf)
userCount = 0
dcCount = 0
foundUsers = ""
foundDCs = ""

For Each line In lines
    If InStr(line, "CN=") > 0 And userCount < 5 Then
        foundUsers = foundUsers & line & vbCrLf
        userCount = userCount + 1
    End If
    If InStr(line, "DC=") > 0 And dcCount < 3 Then
        foundDCs = foundDCs & line & vbCrLf
        dcCount = dcCount + 1
    End If
Next

WScript.Echo vbCrLf & "=== Users (up to 5) ==="
If foundUsers = "" Then
    WScript.Echo "[!] CN= not found"
Else
    WScript.Echo foundUsers
End If

WScript.Echo vbCrLf & "=== Domain Components (up to 3) ==="
If foundDCs = "" Then
    WScript.Echo "[!] DC= not found"
Else
    WScript.Echo foundDCs
End If

sock.Close

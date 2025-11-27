Option Explicit

Const HOST = "127.0.0.1"
Const PORT = 999

Dim http, bindPacket, searchPacket, BASE_DN, searchLen

' Anonymous BIND packet
bindPacket = Chr(&H30) & Chr(&H0C) & _
             Chr(&H02) & Chr(&H01) & Chr(&H01) & _
             Chr(&H60) & Chr(&H07) & _
             Chr(&H02) & Chr(&H01) & Chr(&H03) & _
             Chr(&H04) & Chr(&H00) & _
             Chr(&H80) & Chr(&H00)

' Sample base DN (replace as needed, but can stay as placeholder)
BASE_DN = "DC=corp,DC=local"

' Build raw SearchRequest packet for (objectClass=*) under BASE_DN
searchLen = Len(BASE_DN)
searchPacket = Chr(&H30) & Chr(&H3F) & _
               Chr(&H02) & Chr(&H01) & Chr(&H02) & _
               Chr(&H63) & Chr(&H3A) & _
               Chr(&H04) & Chr(searchLen) & BASE_DN & _
               Chr(&H0A) & Chr(&H01) & Chr(&H02) & _
               Chr(&H0A) & Chr(&H01) & Chr(&H00) & _
               Chr(&H02) & Chr(&H01) & Chr(&H00) & _
               Chr(&H02) & Chr(&H01) & Chr(&H00) & _
               Chr(&H01) & Chr(&H01) & Chr(&H00) & _
               Chr(&H87) & Chr(&H0F) & "(objectClass=user)" & _
               Chr(&H30) & Chr(&H00)

Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

WScript.Echo "[*] Connecting to " & HOST & ":" & PORT
http.Open "POST", "http://" & HOST & ":" & PORT, False
http.SetRequestHeader "Content-Type", "application/octet-stream"

' Send Bind request
http.Send bindPacket
WScript.Sleep 300
If http.Status = 200 Or http.Status = 0 Then
    WScript.Echo "[+] Bind likely accepted."
Else
    WScript.Echo "[!] Bind failed: HTTP " & http.Status
    WScript.Quit(1)
End If

' Send SearchRequest
WScript.Echo "[*] Sending LDAP SearchRequest for users..."
http.Open "POST", "http://" & HOST & ":" & PORT, False
http.SetRequestHeader "Content-Type", "application/octet-stream"
http.Send searchPacket
WScript.Sleep 500

Dim raw, line, i, found, lines
raw = http.ResponseBody

If LenB(raw) = 0 Then
    WScript.Echo "[!] No response to SearchRequest."
    WScript.Quit(1)
End If

' Try to extract any CN= or DC= strings
Dim str, b, ch
str = ""
For i = 1 To LenB(raw)
    b = AscB(MidB(raw, i, 1))
    ch = Chr(b)
    If b >= 32 And b < 127 Then
        str = str & ch
    Else
        str = str & " "
    End If
Next

WScript.Echo "[✓] Raw LDAP response string:"
lines = Split(str, " ")
found = 0
For Each line In lines
    If InStr(line, "CN=") > 0 Or InStr(line, "DC=") > 0 Then
        WScript.Echo "  → " & line
        found = found + 1
        If found >= 10 Then Exit For
    End If
Next

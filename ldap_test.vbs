' Подключается по http://127.0.0.1:9999 и шлёт RAW LDAP BindRequest
' Затем парсит ответ и вытаскивает все строки, содержащие CN= или DC=

Option Explicit

Dim host, port, bindPacket, http, resp, i, matches
host = "127.0.0.1"
port = 9999

bindPacket = Chr(&H30) & Chr(&H0C) & _
             Chr(&H02) & Chr(&H01) & Chr(&H01) & _
             Chr(&H60) & Chr(&H07) & _
             Chr(&H02) & Chr(&H01) & Chr(&H03) & _
             Chr(&H04) & Chr(&H00) & _
             Chr(&H80) & Chr(&H00)

On Error Resume Next
Set http = CreateObject("MSXML2.ServerXMLHTTP")
http.Open "POST", "http://" & host & ":" & port, False
http.setRequestHeader "Content-Type", "application/octet-stream"
http.Send bindPacket

If http.status <> 200 And http.status <> 0 Then
    WScript.Echo "[✗] LDAP через прокси не ответил. Статус: " & http.status
    WScript.Quit
End If

WScript.Echo "[✓] Получен ответ от LDAP через прокси. Ищу CN= и DC=..."

Dim raw, result
raw = http.responseText
result = ""

For i = 1 To Len(raw) - 4
    If Mid(raw, i, 3) = "CN=" Then
        result = result & Mid(raw, i, 30) & vbCrLf
    End If
    If Mid(raw, i, 3) = "DC=" Then
        result = result & Mid(raw, i, 30) & vbCrLf
    End If
Next

If result = "" Then
    WScript.Echo "[!] CN=/DC= не найдены в ответе."
Else
    WScript.Echo result
End If

Option Explicit

Const HOST = "127.0.0.1"
Const PORT = 999

Dim http, bindPacket

bindPacket = Chr(&H30) & Chr(&H0C) & _
             Chr(&H02) & Chr(&H01) & Chr(&H01) & _
             Chr(&H60) & Chr(&H07) & _
             Chr(&H02) & Chr(&H01) & Chr(&H03) & _
             Chr(&H04) & Chr(&H00) & _
             Chr(&H80) & Chr(&H00)

On Error Resume Next
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
If Err.Number <> 0 Then
    WScript.Echo "[!] Ошибка: WinHttp не доступен."
    WScript.Quit 1
End If

WScript.Echo "[*] Подключаюсь к " & HOST & ":" & PORT & "..."

http.Open "POST", "http://" & HOST & ":" & PORT, False
http.SetRequestHeader "Content-Type", "application/octet-stream"

On Error Resume Next
http.Send bindPacket

If Err.Number <> 0 Then
    WScript.Echo "[✗] Ошибка при отправке: " & Err.Description
Else
    WScript.Echo "[✓] Отправлено. Код ответа: " & http.Status
End If

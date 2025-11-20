' Port Scanner VBScript
' Сканирует топ 100 портов на указанном IP-адресе

Option Explicit

' Топ 100 наиболее используемых портов
Dim ports
ports = Array(21, 22, 23, 25, 53, 80, 110, 111, 135, 139, _
              143, 443, 445, 993, 995, 1723, 3306, 3389, 5900, 8080, _
              20, 69, 137, 138, 161, 162, 389, 636, 989, 990, _
              1433, 1434, 1521, 2049, 2082, 2083, 2086, 2087, 2095, 2096, _
              3128, 5432, 5800, 5801, 8000, 8008, 8081, 8443, 8888, 9090, _
              514, 515, 631, 873, 1080, 1194, 1645, 1646, 3690, 5060, _
              5061, 5222, 5269, 5353, 6379, 6660, 6661, 6662, 6663, 6664, _
              6665, 6666, 6667, 6668, 6669, 7000, 7001, 8765, 9091, 9100, _
              9418, 9999, 10000, 11211, 27017, 27018, 27019, 28017, 50000, 50001, _
              1025, 1026, 1027, 1028, 1029, 1030, 1031, 1032, 1033, 1034)

' Получаем IP-адрес от пользователя
Dim targetIP
targetIP = InputBox("Введите IP-адрес для сканирования:", "Port Scanner", "127.0.0.1")

If targetIP = "" Then
    MsgBox "Сканирование отменено", vbInformation
    WScript.Quit
End If

' Проверяем формат IP
If Not IsValidIP(targetIP) Then
    MsgBox "Неверный формат IP-адреса!", vbCritical
    WScript.Quit
End If

' Создаем файл лога
Dim fso, logFile, logPath
Set fso = CreateObject("Scripting.FileSystemObject")
logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\scan_log_" & Replace(Replace(Replace(Now(), ":", "-"), " ", "_"), "/", "-") & ".txt"
Set logFile = fso.CreateTextFile(logPath, True)

' Записываем заголовок
logFile.WriteLine "========================================="
logFile.WriteLine "Port Scanner Log"
logFile.WriteLine "Целевой IP: " & targetIP
logFile.WriteLine "Дата сканирования: " & Now()
logFile.WriteLine "========================================="
logFile.WriteLine ""

MsgBox "Начинается сканирование " & targetIP & vbCrLf & "Это может занять несколько минут..." & vbCrLf & "Лог будет сохранен в:" & vbCrLf & logPath, vbInformation

' Переменные для статистики
Dim openPorts, closedPorts, i
openPorts = 0
closedPorts = 0

' Сканируем порты
For i = 0 To UBound(ports)
    Dim port, isOpen
    port = ports(i)
    isOpen = CheckPort(targetIP, port)
    
    If isOpen Then
        logFile.WriteLine "Порт " & port & " - ОТКРЫТ"
        openPorts = openPorts + 1
    Else
        logFile.WriteLine "Порт " & port & " - ЗАКРЫТ"
        closedPorts = closedPorts + 1
    End If
Next

' Записываем статистику
logFile.WriteLine ""
logFile.WriteLine "========================================="
logFile.WriteLine "Результаты сканирования:"
logFile.WriteLine "Всего портов просканировано: " & (openPorts + closedPorts)
logFile.WriteLine "Открытых портов: " & openPorts
logFile.WriteLine "Закрытых портов: " & closedPorts
logFile.WriteLine "========================================="

logFile.Close
Set logFile = Nothing
Set fso = Nothing

MsgBox "Сканирование завершено!" & vbCrLf & vbCrLf & _
       "Открытых портов: " & openPorts & vbCrLf & _
       "Закрытых портов: " & closedPorts & vbCrLf & vbCrLf & _
       "Лог сохранен в:" & vbCrLf & logPath, vbInformation

' Функция проверки порта
Function CheckPort(ip, port)
    On Error Resume Next
    
    Dim socket, connected
    connected = False
    
    ' Создаем объект Winsock
    Set socket = CreateObject("MSWinsock.Winsock")
    
    If Err.Number = 0 Then
        ' Устанавливаем таймаут
        socket.Connect ip, port
        
        ' Ждем немного
        WScript.Sleep 500
        
        ' Проверяем состояние подключения
        If socket.State = 7 Then ' sckConnected
            connected = True
            socket.Close
        End If
    Else
        ' Если MSWinsock недоступен, используем альтернативный метод
        Err.Clear
        connected = CheckPortAlternative(ip, port)
    End If
    
    Set socket = Nothing
    CheckPort = connected
End Function

' Альтернативный метод проверки порта (через WMI)
Function CheckPortAlternative(ip, port)
    On Error Resume Next
    
    Dim shell, result
    Set shell = CreateObject("WScript.Shell")
    
    ' Используем PowerShell для проверки порта
    Dim cmd
    cmd = "powershell -Command ""$tcpClient = New-Object System.Net.Sockets.TcpClient; " & _
          "try { $tcpClient.Connect('" & ip & "', " & port & "); " & _
          "$tcpClient.Close(); exit 0 } catch { exit 1 }"""
    
    result = shell.Run(cmd, 0, True)
    
    Set shell = Nothing
    
    CheckPortAlternative = (result = 0)
End Function

' Функция проверки формата IP
Function IsValidIP(ip)
    Dim parts, i, part
    IsValidIP = False
    
    parts = Split(ip, ".")
    If UBound(parts) <> 3 Then Exit Function
    
    For i = 0 To 3
        If Not IsNumeric(parts(i)) Then Exit Function
        part = CInt(parts(i))
        If part < 0 Or part > 255 Then Exit Function
    Next
    
    IsValidIP = True
End Function

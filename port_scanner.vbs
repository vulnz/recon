' Enhanced Port Scanner VBScript with Multiple Validation Methods
' Scans top 100 ports on specified IP address

Option Explicit

' Top 100 most commonly used ports
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

' Get IP address from user
Dim targetIP
targetIP = InputBox("Enter IP address to scan:", "Port Scanner", "127.0.0.1")

If targetIP = "" Then
    MsgBox "Scan cancelled", vbInformation
    WScript.Quit
End If

' Validate IP format
If Not IsValidIP(targetIP) Then
    MsgBox "Invalid IP address format!", vbCritical
    WScript.Quit
End If

' Create log file
Dim fso, logFile, logPath
Set fso = CreateObject("Scripting.FileSystemObject")
logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\scan_log_" & Replace(Replace(Replace(Now(), ":", "-"), " ", "_"), "/", "-") & ".txt"
Set logFile = fso.CreateTextFile(logPath, True)

' Write header
logFile.WriteLine "========================================="
logFile.WriteLine "Port Scanner Log"
logFile.WriteLine "Target IP: " & targetIP
logFile.WriteLine "Scan Date: " & Now()
logFile.WriteLine "Validation: Multiple methods (3 checks)"
logFile.WriteLine "========================================="
logFile.WriteLine ""

MsgBox "Starting scan of " & targetIP & vbCrLf & "This may take several minutes..." & vbCrLf & "Log will be saved to:" & vbCrLf & logPath, vbInformation

' Statistics variables
Dim openPorts, closedPorts, i, scannedCount
openPorts = 0
closedPorts = 0
scannedCount = 0

' Create progress tracking
Dim totalPorts
totalPorts = UBound(ports) + 1

' Scan ports
For i = 0 To UBound(ports)
    Dim port, isOpen, checkResults
    port = ports(i)
    
    ' Use multiple validation methods
    isOpen = CheckPortMultipleWays(targetIP, port, checkResults)
    
    scannedCount = scannedCount + 1
    
    If isOpen Then
        logFile.WriteLine "Port " & FormatPort(port) & " - OPEN " & checkResults
        openPorts = openPorts + 1
    Else
        logFile.WriteLine "Port " & FormatPort(port) & " - CLOSED " & checkResults
        closedPorts = closedPorts + 1
    End If
    
    ' Progress indicator every 10 ports
    If scannedCount Mod 10 = 0 Then
        WScript.Echo "Progress: " & scannedCount & "/" & totalPorts & " ports scanned..."
    End If
Next

' Write statistics
logFile.WriteLine ""
logFile.WriteLine "========================================="
logFile.WriteLine "Scan Results:"
logFile.WriteLine "Total ports scanned: " & (openPorts + closedPorts)
logFile.WriteLine "Open ports: " & openPorts
logFile.WriteLine "Closed ports: " & closedPorts
logFile.WriteLine "========================================="

logFile.Close
Set logFile = Nothing
Set fso = Nothing

MsgBox "Scan completed!" & vbCrLf & vbCrLf & _
       "Open ports: " & openPorts & vbCrLf & _
       "Closed ports: " & closedPorts & vbCrLf & vbCrLf & _
       "Log saved to:" & vbCrLf & logPath, vbInformation

' Function to check port using multiple validation methods
Function CheckPortMultipleWays(ip, port, ByRef resultDetails)
    Dim method1, method2, method3, totalSuccess
    
    ' Method 1: PowerShell Test-NetConnection
    method1 = CheckPortPowerShell(ip, port)
    
    ' Method 2: PowerShell TcpClient
    method2 = CheckPortTcpClient(ip, port)
    
    ' Method 3: PowerShell with Socket
    method3 = CheckPortSocket(ip, port)
    
    ' Count successful connections
    totalSuccess = 0
    If method1 Then totalSuccess = totalSuccess + 1
    If method2 Then totalSuccess = totalSuccess + 1
    If method3 Then totalSuccess = totalSuccess + 1
    
    ' Build result details string
    resultDetails = "[M1:" & IIf(method1, "✓", "✗") & " M2:" & IIf(method2, "✓", "✗") & " M3:" & IIf(method3, "✓", "✗") & "]"
    
    ' Port is considered open if at least 2 out of 3 methods confirm it
    CheckPortMultipleWays = (totalSuccess >= 2)
End Function

' Method 1: PowerShell Test-NetConnection (most reliable for modern Windows)
Function CheckPortPowerShell(ip, port)
    On Error Resume Next
    Dim shell, cmd, result
    Set shell = CreateObject("WScript.Shell")
    
    ' Use Test-NetConnection with timeout
    cmd = "powershell -NoProfile -NonInteractive -Command """ & _
          "$result = Test-NetConnection -ComputerName '" & ip & "' -Port " & port & " -WarningAction SilentlyContinue -InformationLevel Quiet -ErrorAction SilentlyContinue; " & _
          "if($result) { exit 0 } else { exit 1 }"""
    
    result = shell.Run(cmd, 0, True)
    
    Set shell = Nothing
    CheckPortPowerShell = (Err.Number = 0 And result = 0)
End Function

' Method 2: PowerShell TcpClient with timeout
Function CheckPortTcpClient(ip, port)
    On Error Resume Next
    Dim shell, cmd, result
    Set shell = CreateObject("WScript.Shell")
    
    cmd = "powershell -NoProfile -NonInteractive -Command """ & _
          "$tcpClient = New-Object System.Net.Sockets.TcpClient; " & _
          "$tcpClient.ReceiveTimeout = 1000; " & _
          "$tcpClient.SendTimeout = 1000; " & _
          "try { " & _
          "$tcpClient.Connect('" & ip & "', " & port & "); " & _
          "Start-Sleep -Milliseconds 100; " & _
          "if($tcpClient.Connected) { " & _
          "$tcpClient.Close(); exit 0 } else { exit 1 } " & _
          "} catch { exit 1 } finally { $tcpClient.Dispose() }"""
    
    result = shell.Run(cmd, 0, True)
    
    Set shell = Nothing
    CheckPortTcpClient = (Err.Number = 0 And result = 0)
End Function

' Method 3: PowerShell Socket connection
Function CheckPortSocket(ip, port)
    On Error Resume Next
    Dim shell, cmd, result
    Set shell = CreateObject("WScript.Shell")
    
    cmd = "powershell -NoProfile -NonInteractive -Command """ & _
          "$socket = New-Object System.Net.Sockets.Socket([System.Net.Sockets.AddressFamily]::InterNetwork, [System.Net.Sockets.SocketType]::Stream, [System.Net.Sockets.ProtocolType]::Tcp); " & _
          "$socket.ReceiveTimeout = 1000; " & _
          "$socket.SendTimeout = 1000; " & _
          "try { " & _
          "$endpoint = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Parse('" & ip & "'), " & port & "); " & _
          "$result = $socket.BeginConnect($endpoint, $null, $null); " & _
          "$success = $result.AsyncWaitHandle.WaitOne(1000, $true); " & _
          "if($success -and $socket.Connected) { " & _
          "$socket.EndConnect($result); " & _
          "$socket.Close(); exit 0 } else { exit 1 } " & _
          "} catch { exit 1 } finally { $socket.Dispose() }"""
    
    result = shell.Run(cmd, 0, True)
    
    Set shell = Nothing
    CheckPortSocket = (Err.Number = 0 And result = 0)
End Function

' Helper function to format port number with padding
Function FormatPort(port)
    FormatPort = Right("     " & port, 5)
End Function

' Function to validate IP format
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

' Helper function for IIf replacement (VBScript doesn't have IIf)
Function IIf(condition, trueValue, falseValue)
    If condition Then
        IIf = trueValue
    Else
        IIf = falseValue
    End If
End Function

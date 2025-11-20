Option Explicit

Dim target, ports, port, alive, objHTTP, startTick, endTick, i

' =========================================
' 1. Ask user for target subnet
' =========================================
target = InputBox("Enter subnet to scan (example: 172.18.96)", "Subnet Scanner")

If target = "" Then
    WScript.Echo "No subnet entered. Exiting."
    WScript.Quit
End If

' =========================================
' 2. Top 100 common ports list
' =========================================
ports = Array(21,22,23,25,53,67,68,69,80,110,111,113,119,123,135,137,138,139,143,161,162,179,389,443,445,465,514,587,593,631,636,873,902,989,990,993,995,1025,1026,1027,1028,1029,1080,1194,1214,1241,1433,1434,1521,1723,1730,1812,1813,1883,2000,2049,2082,2083,2086,2087,2095,2096,2100,2222,2483,2484,3128,3268,3269,3306,3389,3478,3632,3690,4333,4500,4567,4662,4899,5000,5001,5060,5222,5432,5500,5631,5632,5800,5900,6000,6001,6881,8080,8081,8082,8443,8888,9000,9090,9100)

' =========================================
' 3. Scan entire /24 for alive hosts
' =========================================

WScript.Echo "Scanning alive hosts on " & target & ".0/24 ..."

Dim aliveHosts() : ReDim aliveHosts(0)
alive = 0

For i = 1 To 254
    If PingHost(target & "." & i) Then
        alive = alive + 1
        ReDim Preserve aliveHosts(alive)
        aliveHosts(alive - 1) = target & "." & i
        WScript.Echo "Alive: " & target & "." & i
    End If
Next

WScript.Echo vbCrLf & "Alive hosts found: " & alive & vbCrLf

If alive = 0 Then
    WScript.Echo "No hosts alive. Exiting."
    WScript.Quit
End If

' =========================================
' 4. Scan top 100 ports for each alive host
' =========================================

For i = 0 To UBound(aliveHosts)
    WScript.Echo vbCrLf & "===== Host " & aliveHosts(i) & " ====="
    ScanPorts aliveHosts(i), ports
Next

WScript.Echo vbCrLf & "SCAN COMPLETE."

' =========================================
' FUNCTIONS
' =========================================

Function PingHost(ip)
    Dim objPing, status
    Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where Address='" & ip & "'")
    For Each status In objPing
        If Not IsNull(status.StatusCode) Then
            If status.StatusCode = 0 Then
                PingHost = True
                Exit Function
            End If
        End If
    Next
    PingHost = False
End Function

Function ScanPorts(ip, portsArray)
    Dim p, result
    For Each p In portsArray
        If CheckTCP(ip, p) Then
            WScript.Echo "   [+] Port " & p & " OPEN"
        Else
            ' Uncomment to show closed ports:
            ' WScript.Echo "   [-] Port " & p & " closed"
        End If
    Next
End Function

Function CheckTCP(ip, port)
    On Error Resume Next
    Dim http
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "GET", "http://" & ip & ":" & port & "/", False
    http.SetTimeouts 200, 200, 200, 200
    http.Send

    If Err.Number = 0 Then
        CheckTCP = True
    Else
        CheckTCP = False
    End If

    Err.Clear
End Function

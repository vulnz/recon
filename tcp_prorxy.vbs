' ==============================
' Socat-like TCP Proxy in VBScript
' Author: Valera (ChatGPT)
' ==============================

Option Explicit
On Error Resume Next

Const LOCAL_PORT = 9999
Const REMOTE_HOST = "172.31.198.3"
Const REMOTE_PORT = 389

Dim listener, client, remote
Dim data, rdata

WScript.Echo "==[ TCP Proxy - Listening on port " & LOCAL_PORT & " ]=="

Set listener = CreateObject("MSWinsock.Winsock")
listener.Protocol = 0  ' TCP
listener.LocalPort = LOCAL_PORT
listener.Listen

Do
    If listener.ConnectionState = 7 Then ' Listening and ready
        Set client = listener.Accept()
        If Not client Is Nothing Then
            WScript.Echo "[+] Incoming connection from " & client.RemoteHostIP

            ' Connect to remote server
            Set remote = CreateObject("MSWinsock.Winsock")
            remote.Protocol = 0
            remote.RemoteHost = REMOTE_HOST
            remote.RemotePort = REMOTE_PORT
            remote.Connect

            ' Wait for connection
            Do While remote.Connect <> 7
                WScript.Sleep 50
            Loop
            WScript.Echo "[>] Connected to " & REMOTE_HOST & ":" & REMOTE_PORT

            ' === Forward client → remote
            Do While True
                If client.BytesReceived > 0 Then
                    data = client.GetData()
                    WScript.Echo "[>] Client → Remote: " & Len(data) & " bytes"
                    remote.SendData data
                End If

                If remote.BytesReceived > 0 Then
                    rdata = remote.GetData()
                    WScript.Echo "[<] Remote → Client: " & Len(rdata) & " bytes"
                    client.SendData rdata
                End If

                If client.ConnectionState <> 7 Or remote.ConnectionState <> 7 Then
                    Exit Do
                End If

                WScript.Sleep 50
            Loop

            WScript.Echo "[-] Connection closed."
            client.Close
            remote.Close
            Set client = Nothing
            Set remote = Nothing
        End If
    End If
    WScript.Sleep 100
Loop

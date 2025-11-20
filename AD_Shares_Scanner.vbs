' ========================================
' AD Network Shares Enumeration Script (VBS)
' Find all computers, IPs, shared folders and files
' ========================================

Option Explicit

Dim objFSO, objShell, objNetwork, objRootDSE
Dim objConnection, objCommand, objRecordSet
Dim strDomainDN, strOutputPath, strTimestamp
Dim dictShares, dictComputers
Dim intTotalComputers, intOnlineComputers, intComputersWithShares
Dim intTotalShares, intTotalFiles

' Configuration
Const MAX_COMPUTERS = 100  ' Limit to prevent long scans
Const TIMEOUT = 2          ' Seconds
Const DEEP_SCAN = False    ' Set to True for recursive scan

' Initialize
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")
Set objNetwork = CreateObject("WScript.Network")
Set dictShares = CreateObject("Scripting.Dictionary")
Set dictComputers = CreateObject("Scripting.Dictionary")

intTotalComputers = 0
intOnlineComputers = 0
intComputersWithShares = 0
intTotalShares = 0
intTotalFiles = 0

' Get domain DN
Set objRootDSE = GetObject("LDAP://RootDSE")
strDomainDN = objRootDSE.Get("defaultNamingContext")

' Output path
strTimestamp = Replace(Replace(Replace(Now(), "/", "-"), ":", "-"), " ", "_")
strOutputPath = objFSO.GetParentFolderName(WScript.ScriptFullName)

WScript.Echo "========================================="
WScript.Echo "AD NETWORK SHARES ENUMERATION"
WScript.Echo "========================================="
WScript.Echo ""
WScript.Echo "Domain: " & strDomainDN
WScript.Echo "Max Computers: " & MAX_COMPUTERS
WScript.Echo "Timeout: " & TIMEOUT & " seconds"
WScript.Echo ""

' ==================================================
' Get Computers from AD
' ==================================================

WScript.Echo "[*] Getting computers from Active Directory..."

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")

objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"

Set objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = MAX_COMPUTERS
objCommand.Properties("Searchscope") = 2
objCommand.Properties("Timeout") = 60

' Query for enabled computers with IP
objCommand.CommandText = "SELECT dNSHostName, name, operatingSystem FROM 'LDAP://" & strDomainDN & _
                         "' WHERE objectCategory='computer' AND NOT userAccountControl:1.2.840.113556.1.4.803:=2"

Set objRecordSet = objCommand.Execute

If Not objRecordSet.EOF Then
    objRecordSet.MoveFirst
    
    Do Until objRecordSet.EOF Or intTotalComputers >= MAX_COMPUTERS
        Dim strComputerName, strDNSName, strOS
        
        strComputerName = objRecordSet.Fields("name").Value
        strDNSName = objRecordSet.Fields("dNSHostName").Value
        strOS = objRecordSet.Fields("operatingSystem").Value
        
        If Not IsNull(strDNSName) Then
            dictComputers.Add intTotalComputers, strComputerName & "|" & strDNSName & "|" & strOS
            intTotalComputers = intTotalComputers + 1
        End If
        
        objRecordSet.MoveNext
    Loop
End If

objRecordSet.Close
objConnection.Close

WScript.Echo "[+] Found " & intTotalComputers & " computers"
WScript.Echo ""

' ==================================================
' Scan Each Computer
' ==================================================

WScript.Echo "[*] Scanning computers for shared folders..."
WScript.Echo ""

Dim key, arrComputer, intCounter
intCounter = 0

For Each key In dictComputers.Keys
    intCounter = intCounter + 1
    arrComputer = Split(dictComputers(key), "|")
    
    Dim strCompName, strDNS, strCompOS
    strCompName = arrComputer(0)
    strDNS = arrComputer(1)
    strCompOS = arrComputer(2)
    
    WScript.Echo "  [" & intCounter & "/" & intTotalComputers & "] Checking: " & strDNS
    
    ' Ping computer
    If Not PingComputer(strDNS) Then
        WScript.Echo "    └─ Offline or unreachable"
        WScript.Echo ""
        intOnlineComputers = intOnlineComputers + 0
    Else
        intOnlineComputers = intOnlineComputers + 1
        
        ' Get shares
        Dim arrShares, share
        arrShares = GetComputerShares(strDNS)
        
        If IsNull(arrShares) Or UBound(arrShares) < 0 Then
            WScript.Echo "    └─ No accessible shares"
        Else
            intComputersWithShares = intComputersWithShares + 1
            WScript.Echo "    └─ Found " & (UBound(arrShares) + 1) & " share(s)"
            
            For Each share in arrShares
                Dim arrShareInfo
                arrShareInfo = Split(share, "|")
                
                Dim strShareName, strSharePath
                strShareName = arrShareInfo(0)
                strSharePath = arrShareInfo(1)
                
                WScript.Echo "       ├─ Share: " & strShareName & " (" & strSharePath & ")"
                
                ' Get files
                Dim intFileCount
                intFileCount = GetFileCount(strDNS, strShareName)
                
                If intFileCount >= 0 Then
                    WScript.Echo "       │  └─ " & intFileCount & " file(s)/folder(s)"
                    intTotalFiles = intTotalFiles + intFileCount
                Else
                    WScript.Echo "       │  └─ Access denied"
                End If
                
                intTotalShares = intTotalShares + 1
                
                ' Store share info
                Dim strShareKey
                strShareKey = dictShares.Count
                dictShares.Add strShareKey, strDNS & "|" & strShareName & "|" & strSharePath & "|" & intFileCount & "|" & strCompOS
            Next
        End If
        
        WScript.Echo ""
    End If
Next

' ==================================================
' Generate Reports
' ==================================================

WScript.Echo ""
WScript.Echo "========================================="
WScript.Echo "SCAN COMPLETE - GENERATING REPORTS"
WScript.Echo "========================================="
WScript.Echo ""

' Statistics
WScript.Echo "STATISTICS:"
WScript.Echo "  Total Computers:     " & intTotalComputers
WScript.Echo "  Online Computers:    " & intOnlineComputers
WScript.Echo "  With Shares:         " & intComputersWithShares
WScript.Echo "  Total Shares:        " & intTotalShares
WScript.Echo "  Total Files/Folders: " & intTotalFiles
WScript.Echo ""

' Generate CSV
Call GenerateCSV()

' Generate HTML
Call GenerateHTML()

WScript.Echo ""
WScript.Echo "========================================="
WScript.Echo "SCAN COMPLETE!"
WScript.Echo "========================================="

' ==================================================
' FUNCTIONS
' ==================================================

Function PingComputer(strComputer)
    On Error Resume Next
    
    Dim objPing, objStatus
    Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set objStatus = objPing.ExecQuery("SELECT * FROM Win32_PingStatus WHERE Address = '" & strComputer & "' AND Timeout = " & (TIMEOUT * 1000))
    
    Dim objResult
    For Each objResult In objStatus
        If objResult.StatusCode = 0 Then
            PingComputer = True
            Exit Function
        End If
    Next
    
    PingComputer = False
End Function

Function GetComputerShares(strComputer)
    On Error Resume Next
    
    Dim objWMI, colShares, objShare
    Dim arrResults()
    Dim intIndex
    
    intIndex = -1
    
    Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    
    If Err.Number <> 0 Then
        GetComputerShares = Null
        Exit Function
    End If
    
    Set colShares = objWMI.ExecQuery("SELECT * FROM Win32_Share WHERE Type = 0")
    
    If Err.Number <> 0 Then
        GetComputerShares = Null
        Exit Function
    End If
    
    For Each objShare In colShares
        ' Skip admin shares (ending with $)
        If Right(objShare.Name, 1) <> "$" Then
            intIndex = intIndex + 1
            ReDim Preserve arrResults(intIndex)
            arrResults(intIndex) = objShare.Name & "|" & objShare.Path
        End If
    Next
    
    If intIndex >= 0 Then
        GetComputerShares = arrResults
    Else
        GetComputerShares = Null
    End If
End Function

Function GetFileCount(strComputer, strShareName)
    On Error Resume Next
    
    Dim objFolder, colFiles
    Dim strUNCPath
    Dim intCount
    
    strUNCPath = "\\" & strComputer & "\" & strShareName
    intCount = 0
    
    Set objFolder = objFSO.GetFolder(strUNCPath)
    
    If Err.Number <> 0 Then
        GetFileCount = -1
        Exit Function
    End If
    
    ' Count files and folders in root
    Set colFiles = objFolder.Files
    intCount = colFiles.Count
    
    Set colFiles = objFolder.SubFolders
    intCount = intCount + colFiles.Count
    
    GetFileCount = intCount
End Function

Sub GenerateCSV()
    Dim strCSVPath, objFile, key, arrShare
    
    strCSVPath = strOutputPath & "\AD_Shares_" & strTimestamp & ".csv"
    Set objFile = objFSO.CreateTextFile(strCSVPath, True)
    
    ' Header
    objFile.WriteLine "Computer,ShareName,SharePath,FileCount,OS"
    
    ' Data
    For Each key In dictShares.Keys
        arrShare = Split(dictShares(key), "|")
        objFile.WriteLine arrShare(0) & "," & arrShare(1) & "," & arrShare(2) & "," & arrShare(3) & "," & arrShare(4)
    Next
    
    objFile.Close
    
    WScript.Echo "[+] CSV exported to: " & strCSVPath
End Sub

Sub GenerateHTML()
    Dim strHTMLPath, objFile, strHTML
    
    strHTMLPath = strOutputPath & "\AD_Shares_Report_" & strTimestamp & ".html"
    Set objFile = objFSO.CreateTextFile(strHTMLPath, True)
    
    strHTML = "<!DOCTYPE html><html><head><meta charset='UTF-8'><title>AD Shares Report</title>"
    strHTML = strHTML & "<style>"
    strHTML = strHTML & "body{font-family:'Segoe UI',Arial;margin:20px;background:#f5f5f5}"
    strHTML = strHTML & ".header{background:linear-gradient(135deg,#667eea,#764ba2);color:white;padding:30px;border-radius:10px;margin-bottom:20px}"
    strHTML = strHTML & ".stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:15px;margin:20px 0}"
    strHTML = strHTML & ".stat-box{background:white;padding:20px;border-radius:8px;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,0.1)}"
    strHTML = strHTML & ".stat-number{font-size:36px;font-weight:bold;color:#667eea}"
    strHTML = strHTML & ".stat-label{color:#7f8c8d;font-size:14px;margin-top:5px}"
    strHTML = strHTML & "table{width:100%;border-collapse:collapse;background:white;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.1);margin:20px 0}"
    strHTML = strHTML & "th{background:#667eea;color:white;padding:15px;text-align:left}"
    strHTML = strHTML & "td{padding:12px 15px;border-bottom:1px solid #eee}"
    strHTML = strHTML & "tr:hover{background:#f8f9fa}"
    strHTML = strHTML & "h2{color:#2c3e50;border-bottom:2px solid #667eea;padding-bottom:10px}"
    strHTML = strHTML & "</style></head><body>"
    
    strHTML = strHTML & "<div class='header'><h1>🌐 AD Network Shares Report</h1>"
    strHTML = strHTML & "<p><strong>Generated:</strong> " & Now() & " | <strong>Scanned:</strong> " & intTotalComputers & " computers</p></div>"
    
    strHTML = strHTML & "<div class='stats'>"
    strHTML = strHTML & "<div class='stat-box'><div class='stat-number'>" & intTotalComputers & "</div><div class='stat-label'>Total Computers</div></div>"
    strHTML = strHTML & "<div class='stat-box'><div class='stat-number'>" & intOnlineComputers & "</div><div class='stat-label'>Online</div></div>"
    strHTML = strHTML & "<div class='stat-box'><div class='stat-number'>" & intComputersWithShares & "</div><div class='stat-label'>With Shares</div></div>"
    strHTML = strHTML & "<div class='stat-box'><div class='stat-number'>" & intTotalShares & "</div><div class='stat-label'>Total Shares</div></div>"
    strHTML = strHTML & "<div class='stat-box'><div class='stat-number'>" & intTotalFiles & "</div><div class='stat-label'>Files/Folders</div></div>"
    strHTML = strHTML & "</div>"
    
    strHTML = strHTML & "<h2>📁 All Shares (" & dictShares.Count & ")</h2>"
    strHTML = strHTML & "<table><thead><tr><th>Computer</th><th>Share Name</th><th>Path</th><th>Files/Folders</th><th>OS</th></tr></thead><tbody>"
    
    Dim key, arrShare
    For Each key In dictShares.Keys
        arrShare = Split(dictShares(key), "|")
        strHTML = strHTML & "<tr><td>" & arrShare(0) & "</td><td><strong>" & arrShare(1) & "</strong></td><td>" & arrShare(2) & "</td><td>" & arrShare(3) & "</td><td>" & arrShare(4) & "</td></tr>"
    Next
    
    strHTML = strHTML & "</tbody></table>"
    
    strHTML = strHTML & "<div style='margin-top:30px;padding:20px;background:#ecf0f1;border-radius:8px'>"
    strHTML = strHTML & "<h3>📌 Security Recommendations</h3><ul>"
    strHTML = strHTML & "<li>Review all shares for unnecessary exposure</li>"
    strHTML = strHTML & "<li>Remove shares that are no longer needed</li>"
    strHTML = strHTML & "<li>Implement proper access controls</li>"
    strHTML = strHTML & "<li>Avoid storing sensitive data in network shares</li>"
    strHTML = strHTML & "<li>Regular audits of network shares (monthly)</li>"
    strHTML = strHTML & "</ul></div>"
    
    strHTML = strHTML & "</body></html>"
    
    objFile.Write strHTML
    objFile.Close
    
    WScript.Echo "[+] HTML report generated: " & strHTMLPath
End Sub

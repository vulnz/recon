'====================================================================
' AD Password Audit Tool
' Checks: NULL session and one weak password (10+ chars)
' IMPORTANT: Use only with proper authorization
'====================================================================

Option Explicit

Const ADS_SECURE_AUTHENTICATION = 1

Dim objFSO, objFile
Dim strUserList, strDomain
Dim strPassword1, strPassword2
Dim strLine, strUser
Dim intTotalUsers, intCurrentUser, intValidFound
Dim objOutput, objSuccess
Dim dtStart, dtEnd

' Configuration
strUserList = "users.txt"
strDomain = ""  ' Auto-detect domain

' Two passwords to check:
strPassword1 = ""  ' NULL/Empty password
strPassword2 = "P@ssw0rd123"  ' 11 characters

' Initialize
Set objFSO = CreateObject("Scripting.FileSystemObject")
intValidFound = 0
intCurrentUser = 0

dtStart = Now()

' Create output files
Set objOutput = objFSO.CreateTextFile("audit_report.txt", True)
Set objSuccess = objFSO.CreateTextFile("success.txt", True)

WScript.Echo "=========================================="
WScript.Echo "  AD PASSWORD AUDIT TOOL"
WScript.Echo "=========================================="
WScript.Echo "Start Time: " & dtStart
WScript.Echo ""

objSuccess.WriteLine "=========================================="
objSuccess.WriteLine "VALID CREDENTIALS FOUND"
objSuccess.WriteLine "Date: " & Now()
objSuccess.WriteLine "=========================================="
objSuccess.WriteLine ""

' Get domain
If strDomain = "" Then
    strDomain = GetDomain()
    If strDomain = "" Then
        WScript.Echo "[ERROR] Could not determine domain."
        objOutput.Close
        objSuccess.Close
        WScript.Quit
    End If
End If

WScript.Echo "[INFO] Domain: " & strDomain
WScript.Echo "[INFO] Checking 2 passwords:"
WScript.Echo "  1. [NULL/EMPTY]"
WScript.Echo "  2. " & strPassword2
WScript.Echo ""

' Check user list file
If Not objFSO.FileExists(strUserList) Then
    WScript.Echo "[ERROR] File '" & strUserList & "' not found!"
    objOutput.Close
    objSuccess.Close
    WScript.Quit
End If

' Count users
intTotalUsers = CountUsers(strUserList)
WScript.Echo "[INFO] Total users: " & intTotalUsers
WScript.Echo ""
WScript.Echo "=========================================="
WScript.Echo "STARTING AUDIT..."
WScript.Echo "=========================================="
WScript.Echo ""

objOutput.WriteLine "Domain: " & strDomain
objOutput.WriteLine "Total users: " & intTotalUsers
objOutput.WriteLine ""
objOutput.WriteLine "RESULTS:"
objOutput.WriteLine "=========================================="
objOutput.WriteLine ""

' Check each user
Set objFile = objFSO.OpenTextFile(strUserList, 1)

Do Until objFile.AtEndOfStream
    strLine = Trim(objFile.ReadLine)
    
    If strLine <> "" And Left(strLine, 1) <> "#" Then
        intCurrentUser = intCurrentUser + 1
        
        WScript.Echo "[" & intCurrentUser & "/" & intTotalUsers & "] Checking: " & strLine
        objOutput.WriteLine "User: " & strLine
        
        ' Check Password 1 (NULL)
        WScript.Echo "  -> Trying NULL password..."
        If CheckPassword(strDomain, strLine, strPassword1) Then
            WScript.Echo "  [VALID] NULL password works!"
            WScript.Echo ""
            objOutput.WriteLine "  [VALID] NULL password"
            objSuccess.WriteLine "[" & Now() & "] " & strLine & " : [NULL]"
            intValidFound = intValidFound + 1
        Else
            WScript.Echo "  [NOT VALID] NULL password"
            objOutput.WriteLine "  [NOT VALID] NULL password"
        End If
        
        ' Check Password 2
        WScript.Echo "  -> Trying password: " & strPassword2 & "..."
        If CheckPassword(strDomain, strLine, strPassword2) Then
            WScript.Echo "  [VALID] Password works: " & strPassword2
            WScript.Echo ""
            objOutput.WriteLine "  [VALID] Password: " & strPassword2
            objSuccess.WriteLine "[" & Now() & "] " & strLine & " : " & strPassword2
            intValidFound = intValidFound + 1
        Else
            WScript.Echo "  [NOT VALID] Password: " & strPassword2
            objOutput.WriteLine "  [NOT VALID] Password: " & strPassword2
        End If
        
        WScript.Echo ""
        objOutput.WriteLine ""
    End If
Loop

objFile.Close

dtEnd = Now()

' Summary
WScript.Echo "=========================================="
WScript.Echo "SUMMARY"
WScript.Echo "=========================================="
WScript.Echo "Total users checked: " & intTotalUsers
WScript.Echo "Valid credentials found: " & intValidFound
WScript.Echo "Start time: " & dtStart
WScript.Echo "End time: " & dtEnd
WScript.Echo ""
WScript.Echo "Reports saved:"
WScript.Echo "  - audit_report.txt (full log)"
WScript.Echo "  - success.txt (valid credentials)"
WScript.Echo "=========================================="

objOutput.WriteLine "=========================================="
objOutput.WriteLine "SUMMARY"
objOutput.WriteLine "=========================================="
objOutput.WriteLine "Total users checked: " & intTotalUsers
objOutput.WriteLine "Valid credentials found: " & intValidFound
objOutput.WriteLine "Start time: " & dtStart
objOutput.WriteLine "End time: " & dtEnd

objOutput.Close
objSuccess.Close

Set objFSO = Nothing

WScript.Echo ""
WScript.Echo "Press Enter to exit..."
WScript.StdIn.ReadLine

'====================================================================
' FUNCTIONS
'====================================================================

Function CheckPassword(strDomainPath, strUsername, strPassword)
    On Error Resume Next
    
    Dim objADsPath, objUser
    Dim strUserDN
    
    CheckPassword = False
    
    ' Build user path
    strUserDN = strUsername & "@" & GetDomainName()
    
    ' Try to authenticate
    Set objADsPath = GetObject("LDAP:")
    Set objUser = objADsPath.OpenDSObject(strDomainPath, strUserDN, strPassword, ADS_SECURE_AUTHENTICATION)
    
    If Err.Number = 0 Then
        CheckPassword = True
        Set objUser = Nothing
    End If
    
    Err.Clear
    On Error GoTo 0
End Function

Function GetDomain()
    On Error Resume Next
    
    Dim objRootDSE, strDN
    
    GetDomain = ""
    
    Set objRootDSE = GetObject("LDAP://RootDSE")
    If Err.Number = 0 Then
        strDN = objRootDSE.Get("defaultNamingContext")
        GetDomain = "LDAP://" & strDN
        Set objRootDSE = Nothing
    End If
    
    On Error GoTo 0
End Function

Function GetDomainName()
    On Error Resume Next
    
    Dim objRootDSE, strDN, arrParts, strPart
    Dim strResult
    
    GetDomainName = ""
    strResult = ""
    
    Set objRootDSE = GetObject("LDAP://RootDSE")
    If Err.Number = 0 Then
        strDN = objRootDSE.Get("defaultNamingContext")
        arrParts = Split(strDN, ",")
        
        For Each strPart In arrParts
            If Left(UCase(strPart), 3) = "DC=" Then
                If strResult <> "" Then
                    strResult = strResult & "."
                End If
                strResult = strResult & Mid(strPart, 4)
            End If
        Next
        
        GetDomainName = strResult
        Set objRootDSE = Nothing
    End If
    
    On Error GoTo 0
End Function

Function CountUsers(strFile)
    Dim objF, strL, intCount
    intCount = 0
    
    Set objF = objFSO.OpenTextFile(strFile, 1)
    Do Until objF.AtEndOfStream
        strL = Trim(objF.ReadLine)
        If strL <> "" And Left(strL, 1) <> "#" Then
            intCount = intCount + 1
        End If
    Loop
    objF.Close
    
    CountUsers = intCount
End Function

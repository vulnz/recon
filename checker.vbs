Option Explicit

Dim userSam, domain, rootDSE, ldapPath, conn, cmd, rs

If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript check_user.vbs <sAMAccountName>"
    WScript.Quit
End If

userSam = WScript.Arguments.Item(0)

' Get domain DN
Set rootDSE = GetObject("LDAP://RootDSE")
domain = rootDSE.Get("defaultNamingContext")

ldapPath = "LDAP://" & domain

Set conn = CreateObject("ADODB.Connection")
Set cmd = CreateObject("ADODB.Command")

conn.Provider = "ADsDSOObject"
conn.Open "Active Directory Provider"

Set cmd.ActiveConnection = conn

cmd.CommandText = _
    "<" & ldapPath & ">;(&(objectClass=user)(sAMAccountName=" & userSam & "));" & _
    "distinguishedName,sAMAccountName,userAccountControl,displayName,mail,enabled;subtree"

Set rs = cmd.Execute

If rs.EOF Then
    WScript.Echo "❌ User not found: " & userSam
Else
    Dim uac
    uac = rs.Fields("userAccountControl").Value

    WScript.Echo "Found user: " & rs.Fields("sAMAccountName").Value
    WScript.Echo "DN: " & rs.Fields("distinguishedName").Value
    WScript.Echo "UAC value: " & uac

    ' Check PasswordNotRequired flag (0x20 = 32)
    If (uac And 32) <> 0 Then
        WScript.Echo "⚠️ PasswordNotRequired = TRUE (VULNERABLE)"
    Else
        WScript.Echo "✔ PasswordNotRequired = FALSE"
    End If

    ' Check enabled/disabled
    If (uac And 2) <> 0 Then
        WScript.Echo "❌ Account is DISABLED"
    Else
        WScript.Echo "✔ Account is ENABLED"
    End If
End If

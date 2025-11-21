Option Explicit

Dim rootDSE, domainDN, domainObj, fso, out, htmlPath

Set rootDSE = GetObject("LDAP://RootDSE")
domainDN = rootDSE.Get("defaultNamingContext")
Set domainObj = GetObject("LDAP://" & domainDN)

Set fso = CreateObject("Scripting.FileSystemObject")
htmlPath = "ad_audit01_password_policy.html"
Set out = fso.CreateTextFile(htmlPath, True)

out.WriteLine "<html><head><title>AD Password Policy</title></head><body>"
out.WriteLine "<h2>AD Password Policy - " & domainDN & "</h2>"
out.WriteLine "<table border='1' cellpadding='3' cellspacing='0'>"
out.WriteLine "<tr><th>Setting</th><th>Value</th></tr>"

On Error Resume Next

out.WriteLine "<tr><td>minPwdLength</td><td>" & domainObj.Get("minPwdLength") & "</td></tr>"
out.WriteLine "<tr><td>maxPwdAge (days)</td><td>" & CStr(Abs(domainObj.Get("maxPwdAge") / 864000000000#)) & "</td></tr>"
out.WriteLine "<tr><td>minPwdAge (days)</td><td>" & CStr(Abs(domainObj.Get("minPwdAge") / 864000000000#)) & "</td></tr>"
out.WriteLine "<tr><td>pwdHistoryLength</td><td>" & domainObj.Get("pwdHistoryLength") & "</td></tr>"
out.WriteLine "<tr><td>pwdProperties</td><td>" & domainObj.Get("pwdProperties") & " (bitmask)</td></tr>"
out.WriteLine "<tr><td>lockoutThreshold</td><td>" & domainObj.Get("lockoutThreshold") & "</td></tr>"
out.WriteLine "<tr><td>lockoutDuration (minutes)</td><td>" & CStr(Abs(domainObj.Get("lockoutDuration") / 600000000#)) & "</td></tr>"
out.WriteLine "<tr><td>lockoutObservationWindow (minutes)</td><td>" & CStr(Abs(domainObj.Get("lockoutObservationWindow") / 600000000#)) & "</td></tr>"

On Error GoTo 0

out.WriteLine "</table>"
out.WriteLine "<p>Note: Review weak values (e.g., minPwdLength &lt; 12, maxPwdAge very high, lockoutThreshold = 0 means no lockout).</p>"
out.WriteLine "</body></html>"
out.Close

WScript.Echo "Report: " & htmlPath

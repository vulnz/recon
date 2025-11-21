Option Explicit
Const DONT_EXPIRE     = 65536
Const PASSWD_NOTREQD  = 32
Const DISABLED        = 2

Dim net, domain, fso, out, d, u, htmlPath
Set net = CreateObject("WScript.Network")
domain = net.UserDomain

Set fso = CreateObject("Scripting.FileSystemObject")
htmlPath = "audit15_ad_users_overview.html"
Set out = fso.CreateTextFile(htmlPath, True)

out.WriteLine "<html><head><title>AD User Overview - " & domain & "</title></head><body>"
out.WriteLine "<h2>AD User Overview - " & domain & "</h2>"
out.WriteLine "<table border='1' cellspacing='0' cellpadding='3'>"
out.WriteLine "<tr><th>User</th><th>Disabled</th><th>PW Never Expires</th><th>No PW Required</th></tr>"

Set d = GetObject("WinNT://" & domain)
d.Filter = Array("User")

For Each u In d
    Dim flags, dis, neverExp, noPw
    flags   = u.Flags
    dis     = IIf((flags And DISABLED) <> 0, "YES", "")
    neverExp= IIf((flags And DONT_EXPIRE) <> 0, "YES", "")
    noPw    = IIf((flags And PASSWD_NOTREQD) <> 0, "YES", "")

    out.WriteLine "<tr><td>" & u.Name & "</td><td>" & dis & "</td><td>" & neverExp & "</td><td>" & noPw & "</td></tr>"
Next

out.WriteLine "</table></body></html>"
out.Close

WScript.Echo "Report: " & htmlPath

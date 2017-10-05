Dim PuValue, StValue, DoValue
Dim disableFirewall
settings = "settings.xml"

' On Error Resume Next

Set xmlDoc = CreateObject("Microsoft.XMLDOM")

xmlDoc.Async = "False"
xmlDoc.Load(settings)

Set colNodes=xmlDoc.selectNodes("//domain/name")

' For Each objNode in colNodes
'     printw "XML : " & objNode.Text
' Next

' Set wShell = WScript.CreateObject("WSCript.shell")
Set wShell = WScript.CreateObject("Shell.Application")


Function printl(txt)
    WScript.StdOut.Write txt
End Function

Function printw(txt)
    WScript.StdOut.WriteLine txt
End Function

Function GetDomainName()
    Dim Info
    Set Info = CreateObject("AdSystemInfo")
    GetDomainName = Info.DomainDNSName
End Function

Function GetFirewallStatus()
    ' Get each Firewall Status : Domain, Public and Private
    Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\localhost\root\default:StdRegProv")

    If err.number = 0 Then
        objReg.GetDWORDValue &H80000002, "SYSTEM\CurrentControlSet\" &_
        "Services\SharedAccess\Parameters\FirewallPolicy\" &_
        "DomainProfile\", "EnableFirewall", DoValue

        objReg.GetDWORDValue &H80000002, "SYSTEM\CurrentControlSet\" &_
        "Services\SharedAccess\Parameters\FirewallPolicy\" &_
        "PublicProfile\", "EnableFirewall", PuValue

        objReg.GetDWORDValue &H80000002, "SYSTEM\CurrentControlSet\" &_
        "Services\SharedAccess\Parameters\FirewallPolicy\" &_
        "StandardProfile\", "EnableFirewall", StValue
    End If

    printw( DoValue & StValue & PuValue)
    
    
End Function

domain = GetDomainName()
res = GetFirewallStatus()

If len(domain) > 5 Then
    printw(domain)
Else
    printw("No Domain found")
    domain = False
End If

If len(domain) > 5 Then
    disableFirewall = False
    For Each objNode in colNodes
        If objNode.Text = """" & domain  & """" Then
            printw "domaine secure"
            disableFirewall = True
        End If
    Next
    printw(disableFirewall)
    If disableFirewall Then
        If DoValue OR StValue OR PuValue Then
            printw( DoValue & StValue & PuValue)
            printw("Firewall enabled. Deactivating...")

            ' PowerShell
            ' Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled False
            For Each profile In Array("Domain", "Public", "Private")
                wShell.ShellExecute "powershell.exe", " Set-NetFirewallProfile -Profile " & profile &" -Enabled False", , "runas", 0
            Next
        Else
            printw("Firewall disabled")
        End If
    End If
End If
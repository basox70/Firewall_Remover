On Error Resume Next

Set wShell = WScript.CreateObject("WSCript.shell")

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
    Dim PuValue, StValue, DoValue
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

    If DoValue OR StValue OR PuValue Then
        printw( DoValue & StValue & PuValue)
        printw("Firewall enabled. Disactivating...")
        ' objReg.SetDWORDValue &H80000002, "SYSTEM\CurrentControlSet\" &_
        ' "Services\SharedAccess\Parameters\FirewallPolicy\" &_
        ' "PublicProfile\", "EnableFirewall", 0

        ' Set WshShell = CreateObject("WScript.Shell")
        ' myKey = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\PublicProfile\EnableFirewall"
        ' WshShell.RegWrite myKey, 0, "REG_DWORD"

        ' PowerShell
        ' Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled False
        wShell.run ("powershell.exe -noexit Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled False")
    Else
        printw("Firewall disabled")
    End If
End Function

domain = GetDomainName()
res = GetFirewallStatus()

If len(domain) > 5 Then
    printw(domain)
Else
    printw("No Domain found")
End If
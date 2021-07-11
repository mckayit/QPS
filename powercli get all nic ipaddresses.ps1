
#PowerCLI C:\> 

Get-VM | Select Name, @{n = "ip address1"; e = { @($_.guest.IPAddress[0]) } }, @{n = "ip address2"; e = { @($_.guest.IPAddress[1]) } }, @{n = "ip address3"; e = { @($_.guest.IPAddress[2]) } }, @{n = "ip address4"; e = { @($_.guest.IPAddress[3]) } } 
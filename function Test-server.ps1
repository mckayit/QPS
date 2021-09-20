function Test-server
{
    #... copies Tests Server to se eif it is up or not and if can be Decomm'd
    <#
    here is how i get the Details for the different domains

            Write-host "Please wait while I get the following Information. `nThis may take a few Min." -ForegroundColor cyan

        write-host '   Getting Servers from DESQLD' -ForegroundColor green
        $DESQLD = Get-ADComputer  -Filter 'operatingsystem -like "*server*" ' -server desqld.internal  -Credential $desqldcreds -Properties description, OperatingSystem , OperatingSystemServicePack, OperatingSystemVersion, CanonicalName, IPv4Address, whenchanged , PasswordExpired , PasswordLastSet , LastLogonDate | select CanonicalName, DistinguishedName, DNSHostName, Enabled, IPv4Address, Name, Description, OperatingSystem, OperatingSystemServicePack, OperatingSystemVersion, whenchanged, PasswordExpired , PasswordLastSet , LastLogonDate
        write-host "found " $desqld.count
        write-host '   Getting Servers from DVDS' -ForegroundColor green
        $DVDS = Get-ADComputer  -Filter 'operatingsystem -like "*server*" '  -Properties description, OperatingSystem , OperatingSystemServicePack, OperatingSystemVersion, CanonicalName, IPv4Address, whenchanged , PasswordExpired , PasswordLastSet , LastLogonDate | select CanonicalName, DistinguishedName, DNSHostName, Enabled, IPv4Address, Name, Description, OperatingSystem, OperatingSystemServicePack, OperatingSystemVersion, whenchanged, PasswordExpired , PasswordLastSet , LastLogonDate
        write-host "found " $dvds.count

        write-host '   Getting Servers from ACDS' -ForegroundColor green
        $acds = Get-ADComputer  -Filter 'operatingsystem -like "*server*" ' -Server edwqpsaddcac01.acds.accpol  -Credential $acdscreds  -Properties description, OperatingSystem , OperatingSystemServicePack, OperatingSystemVersion, CanonicalName, IPv4Address, whenchanged , PasswordExpired , PasswordLastSet , LastLogonDate | select CanonicalName, DistinguishedName, DNSHostName, Enabled, IPv4Address, Name, Description, OperatingSystem, OperatingSystemServicePack, OperatingSystemVersion, whenchanged, PasswordExpired , PasswordLastSet , LastLogonDate
        write-host "found " $acds.count

        write-host '   Getting Server Names from PRDS' -ForegroundColor green
        $names = Get-ADComputer -Filter 'operatingsystem -like "*server*" ' -server prds.qldpol  | Select-Object -ExpandProperty name
        write-host "PRDS A found " $names.count
        $i = 1
        $prds = foreach ($name in $names)
        {
     
            $paramWriteProgress = @{

                
                Activity        = 'Getting PRDs info'
                Status          = "Processing [$i] of [$($names.Count)] Servers"
                PercentComplete = (($i / $names.Count) * 100)
                                
            }
                            
            Write-Progress @paramWriteProgress
                        
            $i++
            Get-ADComputer -Identity $name -server prds.qldpol  -Properties description, OperatingSystem , OperatingSystemServicePack, OperatingSystemVersion, CanonicalName, IPv4Address, whenchanged , PasswordExpired , PasswordLastSet , LastLogonDate | select CanonicalName, DistinguishedName, DNSHostName, Enabled, IPv4Address, Name, Description, OperatingSystem, OperatingSystemServicePack, OperatingSystemVersion, whenchanged, PasswordExpired , PasswordLastSet , LastLogonDate
        }
        write-host "found prds B " $prds.count

    
    #>

    [CmdletBinding()]
    param (
        $Servers,
        [ValidateSet('PRDS', 'DVDS', 'ACDS', 'DESQLD')]
        [String]$DNS_ = 'PRDS',
        [String]$HQRFile2Import
    )
    
    begin
    {
        If ($dns_ -eq 'PRDS') { $dnsserver = '164.112.162.142' } 
        If ($dns_ -eq 'DVDS') { $dnsserver = '164.112.162.250' } 
        If ($dns_ -eq 'ACDS') { $dnsserver = '164.112.162.230' } 
        If ($dns_ -eq 'DESQLD') { $dnsserver = '10.2.109.50' } 
   
    
        try
        {
            $hqr = import-csv $HQRFile2Import
        }
        catch
        {
            Write-host 'Failed to import HQR   `nData will not be fully populated' -forgroundcolor Cyan
        }
    }
    process
    {
        $i = 1
        foreach ($server in $servers)
        {
            $paramWriteProgress = @{

                
                Activity        = 'Getting  info for $($server.name)'
                Status          = "Processing [$i] of [$($servers.Count)] Servers"
                PercentComplete = (($i / $servers.Count) * 100)
                                
            }
                            
            Write-Progress @paramWriteProgress
                        
            $i++
           

            if ($Dns = resolve-dnsname $server.dnshostname -server $dnsserver -ErrorAction SilentlyContinue) 
            {
                $InDNS = "True"  
            }
            else
            {
                $InDNS = "$false"    
            }

            if (test-connection $server.dnshostname -Quiet -count 2) { $PINGname = "True" }
            else
            {
                $PINGname = "False"
            }
            try
            {
                if (test-connection $dns.IPAddress -Quiet -count 2 -ErrorAction SilentlyContinue) { $PINGableIP = "True" }
                else
                {
                    $PINGABLEIP = "False"
                }
            }
            catch { $PINGABLEIP = "False" }

            $decom = ""
            If (-not ($dns.IPAddress) ) { $Decom = "Can be Decommed as does not exist in DNS or Pingable" }

            # checking HQR
            $hqr_ = $hqr | Where-Object { $server.name -contains $_."server name" } 
      

            [PSCustomObject] @{
                "Servername"                                = $server.dnshostname
                "CanonicalName"                             = $server.CanonicalName
                "OperatingSystem"                           = $server.OperatingSystem
                "Description"                               = $server.description
                "Is_Pingable_by_Name"                       = $PINGname
                "Is_Pingable_by_IP"                         = $PINGABLEIP
                "Found_IN_DNS"                              = $inDNS
                "IP_Found_IN_AD"                            = $server.IPv4Address
                "IP_found_IN_DNS"                           = $dns.IPAddress
                "Server_IN_AD"                              = "True"
                'Last_time_server_spoke_2_Domaincontroller' = $server.whenchanged
                'PasswordExpired'                           = $server.PasswordExpired 
                'AD_Object_PasswordLastSet'                 = $server.PasswordLastSet
                'AD_Object_LastLogonDate'                   = $server.LastLogonDate
                'Decom'                                     = $decom
                BusinessOwner                               = $hqr_.'Business Owner - Name'
                Support_Group                               = $hqr_.'Support - Group'
                SystemName                                  = $hqr_.'System - Name'
            }


        }
        
    }
    end
    {
        
    }    
       
}

write-host "Command PS `n`n Test-server -Servers <VAR> -DNS_ prds -HQRFile2Import D:\tmp\hqr.csv  |export-csv d:\tmp\prds_servers_2_b_decommed.csv -NoTypeInformation'" -ForegroundColor green

function get-da
{
    $i = 1
    foreach ($server in $servers)
    {

        $paramWriteProgress = @{

                
            Activity        = 'Getting  info for $($server.appserver)'
            Status          = "Processing [$i] of [$($servers.Count)] Servers"
            PercentComplete = (($i / $servers.Count) * 100)
                                
        }
                            
        Write-Progress @paramWriteProgress
                        
        $i++
           


        $os = $server
        $whenChanged = $server.whenchanged
        $AppServerFQDN = $server.AppServerFQDN
        try
        {
            $OS = get-adcomputer -Identity $server.appserver   -Properties OperatingSystem , OperatingSystemVersion, IPv4Address, whenChanged #-server POLPSBADS01.psba.qld.gov.au -credential $desqldcreds
            $AppServerFQDN = $os.DNSHOSTNAME
            Write-host ':-))' -ForegroundColor green
        }
        catch
        {
            Write-host ':-(' -ForegroundColor Cyan
            $os = $server
            $AppServerFQDN = $server.AppServerFQDN
        }
       
        write-host  " "

        [PSCustomObject] @{
            Agency            = $server.Agency
            Application       = $server.Application
            AppServer         = $server.appserver
            AppServerFQDN     = $AppServerFQDN
            OperatingSystem   = $os.OperatingSystem
            whenLastspokeToAD = $whenChanged
            Appservercopy     = $server.Appservercopy
        }
    }

}


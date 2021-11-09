function get-allSQLServers
{
    <#
 #... gets the SQL servers in Domain and showes SQL Version 
.SYNOPSIS
    gets the SQL servers in Domain and showes SQL Version
.DESCRIPTION
    gets all computer objects in the current Domain that have SQL in the name and then doews a remote Reg connection to see if there is SQL installed and what Version
    This script uses the Add remove Reg keys and then looks for the displayname -like "Microsoft SQL Server * Setup (English)"
    If it is found then it outputs this with the Server name.
    If there is multi versions installed it will display them all.


.PARAMETER none

.
.PARAMETER None
 
.PARAMETER InputObject
    Specifies the object to be processed.  You can also pipe the objects to this command.
.EXAMPLE
    C:\PS>Get-AllSqlServers |export-csv c:\tmp\AllSqlVersion.csv
    C:\PS>Get-AllSQLServers
    Example of how to use this cmdlet

.OUTPUTS
    Output from this cmdlet is an arry of the Results. 
.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    21 Sept 2021
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           21 Sept 2021         Lawrence       Initial Coding

#>
    $servers = Get-ADComputer  -Filter 'operatingsystem -like "*server*" ' 
    $i = 1
    foreach ($server in $servers)
    {
        $paramWriteProgress = @{

                
            Activity        = 'Getting server $($server.name)'
            Status          = "Processing [$i] of [$($servers.Count)] Servers"
            PercentComplete = (($i / $servers.Count) * 100)
                                
        }
                            
        Write-Progress @paramWriteProgress
                        
        $i++
        try
        {

            $installed = invoke-command -computername $server.dnshostname { Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.displayname -like "Microsoft SQL Server * Setup (English)" } | Sort-Object displayname | Select-Object DisplayName, Publisher } -ErrorAction SilentlyContinue

            #sql instance
       

            foreach ($Instal in $installed)
            {
                if ($instal.displayname -like "Microsoft SQL Server * Setup (English)")
                {


                    $pspath = invoke-command -computername $server.dnshostname { (get-itemproperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\*').pschildname }
                    foreach ($ps in $pspath)
                    {

                  
                        $ps1 = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\$ps"
                        


                        $SQLInstances = invoke-command -computername $server.dnshostname -scriptblock { gi ($using:ps1) }
                        [PSCustomObject] @{
                            "servername"        = $server.name
                            "SQLAPP"            = $instal.displayname
                            "Domain"            = $env:USERDNSDOMAIN
                            "Instance name"     = $SQLInstances.pschildname
                            "Instance Property" = $SQLInstances.property 
                        }
                    }


                    
                }
            }

        }
        Catch
        {
            [PSCustomObject] @{
                servername = $server.name
                SQLAPP     = 'Cant not attach to Server'

            }
        }
    }
}

function get-allOracaleServers
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

            $installed = invoke-command -computername $server.dnshostname { Get-ItemProperty HKLM:\SOFTWARE\ORACLE\*  | where { $_.PSChildname -like "KEY_ag*" } } -ErrorAction SilentlyContinue

            #sql instance
       

            foreach ($Instal in $installed)
            {
                if ($instal.PSChildName -match "KEY_ag")
                {

                    [PSCustomObject] @{
                        "servername"        = $server.name
                        "Oracale"           = $instal.PSChildname
                        "App"               = $instal.ORACLE_HOME_NAME
                        "ORACLE_GROUP_NAME" = $instal.ORACLE_GROUP_NAME
                        "Domain"            = $env:USERDNSDOMAIN
                        
 
                    }
                }


                    
            }
        }

    
        Catch
        {
            [PSCustomObject] @{
                servername = $server.name
                APP        = 'Cant not attach to Server'

            }
        }
    }
}

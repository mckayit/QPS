function get-serverup
{
    <#
 #... get-server up or not from csv
.SYNOPSIS
    get-server up or not from csv
.DESCRIPTION
    get-server up or not from csv
.PARAMETER one
    None
.EXAMPLE
    C:\PS>
    Example of how to use this cmdlet
.INPUTS
    Inputs to this cmdlet (if any)
.OUTPUTS
    Output from this cmdlet (if any)
.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    03 Feb 2021
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           3 Feb 2021          Lawrence       Initial Coding

#>
     
    [CmdletBinding()]
    param(
  
    )
    
    
    begin 
    {
             
    }
    
    process 
    {
            
        try 
        {
                    

            $servers = import-csv C:\tmp\list.csv | where { $_.ResourcesAssigned -match 'lawrence' } | select -ExpandProperty name
            $collection = $()
            foreach ($server in $servers)
            {
                $status = @{ "ServerName" = $server; "TimeStamp" = (Get-Date -f s); }
                if (Test-Connection $server -Count 1 -ea 0 )# -Quiet)
                { 
                         
                    $ip = Test-Connection $server -Count 1 -ea 0 | select -ExpandProperty ipv4address
                    $status = @{ "ServerName" = $server; "TimeStamp" = (Get-Date -f s); "ipaddress" = $ip ; "Results" = "UP" }
                    
                } 
                else 
                { 
                    $status["Results"] = "Down" 
                }
                New-Object -TypeName PSObject -Property $status -OutVariable serverStatus
                $collection += $serverStatus
            }
          
        }
        catch 
        {
            Write-Host 'ERROR : $(_.Exception.Message)' -ForegroundColor Magenta
        }
    }
    
    end
    {
        $collection 
    }
}
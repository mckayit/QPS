function start-DomainCleanup_old_Computer_objects
<#
 #... Removes AD computer objects from a Domain
.SYNOPSIS
    Removes AD computer objects from a Domain
.DESCRIPTION
    this will use an arry of AD computer objects and remove them This is used for 
    cleaning up old stale AD computer objects.

.PARAMETER Servers
    List of Server names.
.EXAMPLE
    C:\PS>start-DomainCleanup_old_Computer_objects -servers $serverlist
    C:\PS>start-DomainCleanup_old_Computer_objects -servers server1,server2
    C:\PS>start-DomainCleanup_old_Computer_objects -servers server1,server2 |export-csv d:\temp\servers.csv
    Example of how to use this cmdlet
.OUTPUTS
    Output from this cmdlet is in an arry format.   Servername, Cleanup_Status
.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    14 Sep 2021
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           14 Sept 2021        Lawrence       Initial Coding

#>
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true,
            HelpMessage = 'Server name or array of names.')]
        $Servers
    )

    begin
    {

    }  

    process 
    {
        #Progress bar setup
        $i = 1 
        foreach ($server in $servers)
        {
            $paramWriteProgress = @{

                
                Activity        = 'Removing AD Computer Object from AD'
                Status          = "Processing [$i] of [$($Servers.Count)] computers"
                PercentComplete = (($i / $Servers.Count) * 100)
                                
            }
                            
            Write-Progress @paramWriteProgress
                        
            $i++
            #Progress bar End.
            try
            {
                ## removes the Ad object and leaf object if exist without prompting. 
                get-adcomputer $server | remove-adobject -Recursive -Confirm:$false 

                $done = 'AD Clean up Passed'
            }

            catch { $done = 'AD Clean up Failed.' }

            [PSCustomObject]@{Servername_ = $server
                Cleanup_Status            = $done
            }

        }
    }
    
    end
    {

    }
}

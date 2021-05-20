Function get-serverrolesandFeatures
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true,
            HelpMessage = 'Enter The Serrver array of Name and SID')]
        $servers
    )
    

    $i = 1
    foreach ($server in $servers)
    {
        $servername = $server.servername
        
        #progress bar.
       
        $paramWriteProgress = @{
            Activity         = "getting Server info"
            Status           = "Processing [$i] of [$($Servers.Count)] users"
            PercentComplete  = (($i / $Servers.Count) * 100)
            CurrentOperation = "processing the following Server : [ $($servername)]"
        }
        Write-Progress @paramWriteProgress       
        
        $i++
        
        
        $roles_Feature_Installed = Get-WmiObject -ComputerName $servername  -query 'select * from win32_optionalfeature  where installstate=1' -ErrorAction silentlycontinue

        
         
        $roles_Feature_Installed | add-member -type NoteProperty -name "SID" -value $($server.serversid)
        $roles_Feature_Installed
        
    }
}

function get-serversoftware
{
    <#
 #... Gets the software from Servers
.SYNOPSIS
    Gets the software from Servers
.DESCRIPTION
    gets the Software via Remote Registry from all the servers imputted in via the Array.

    the array must have the following Fields "ServerName" as well as the "SID"
.PARAMETER Servers
    array must have the following Fields "ServerName" as well as the "SID"

.PARAMETER InputObject
    Specifies the object to be processed.  You can also pipe the objects to this command.
.EXAMPLE
    C:\PS>Get-serverSoftware -Servers $arrayname
    C:\PS>Get-serverSoftware -Servers $arrayname |export-csv d:\serversoftware.csv 
    
.INPUTS
    Inputs to this cmdlet (if any)
.OUTPUTS
    Output from this cmdlet (if any)
.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    19 May 2021
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           19 May 2021         Lawrence       Initial Coding

#>
     
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true,
            HelpMessage = 'Enter The Serrver array of Name and SID')]
        $servers
    )
    
    
    begin 
    {
             
    }
    
    process 
    {
        $i = 1
        foreach ($server in $servers)
        {

            #progress bar.
            $paramWriteProgress = @{
                Activity         = "getting Server info"
                Status           = "Processing [$i] of [$($Servers.Count)] users"
                PercentComplete  = (($i / $Servers.Count) * 100)
                CurrentOperation = "processing the following server : [ $($server.Servername)]"
            }
            Write-Progress @paramWriteProgress       
            $i++



            New-pssession -ComputerName $server.serverName -ErrorAction  Silentlycontinue
            $global:apps1 = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object Displayname, Publisher, Displayversion, VersionMajor, VersionMinor, Version, HelpLink, IrlInfoAbout, Comments, installDate 
            $global:apps1 = $global:apps1 += Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object Displayname, Publisher, Displayversion, VersionMajor, VersionMinor, Version, HelpLink, IrlInfoAbout, Comments, installDate
                

            Exit-PSSession
            get-pssession | remove-pssession
            $global:InstalledApps = $null
            $global:InstalledApps = $apps1 | Select-Object | where { $_.Displayname.length -gt 2 }

                    
            [string]$Displayname = $global:InstalledApps.Displayname
            [string]$Publisher = $global:InstalledApps.Publisher
            [string]$Displayversion = $global:InstalledApps.Displayversion
            [string]$VersionMajor = $global:InstalledApps.VersionMajor
            [string]$VersionMinor = $global:InstalledApps.VersionMinor
            [string]$Version = $global:InstalledApps.Version
            [string]$HelpLink = $global:InstalledApps.HelpLink
            [string]$IrlInfoAbout = $global:InstalledApps.IrlInfoAbout
            [string]$Comments = $global:InstalledApps.Comments
            [string]$installDate = $global:InstalledApps.installDate

            <#
               $global:outapps =[PSCustomObject]@{
                    SID            = $Server.ServerSID 
                    Displayname    = $Displayname
                    Publisher      = $Publisher
                    Displayversion = $Displayversion
                    VersionMajor   = $VersionMajor
                    VersionMinor   = $VersionMinor
                    Version        = $Version
                    HelpLink       = $HelpLink
                    IrlInfoAbout   = $IrlInfoAbout
                    Comments       = $Comments
                    installDate    = $installDate
                }
  #>          
  
  
  
            $global:InstalledApps | add-member -type NoteProperty -name "SID" -value $($server.serversid)
            $global:InstalledApps
        }
    }
}




   

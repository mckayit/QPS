function get-serversoftware
{
    <#
 #... Gets the software from Servers
.SYNOPSIS
    Gets the software from Servers
.DESCRIPTION
    gets the Software via Remote Registry from all the servers imputted in via the Array.

    the array must have the following Fields "ServerName"
.PARAMETER Servers
    array must have the following Fields "ServerName"

.PARAMETER InputObject
    Specifies the object to be processed.  You can also pipe the objects to this command.
.EXAMPLE
    C:\PS>Get-serverSoftware -Servers $arrayname
    C:\PS>Get-serverSoftware -Servers SERVERNAME
    
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
                CurrentOperation = "processing the following server : [ $($server)]"
            }
            Write-Progress @paramWriteProgress       
            $i++
            $app = ""
            $apps1 = @{}


            ### Need to put in some error Checking to capture when it cant get onto Server Eg Blocked or can not find it.

            #New-pssession -ComputerName $server -ErrorAction  Silentlycontinue
            $apps1 = Invoke-command -computer $Server { Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object Displayname, Publisher, Displayversion, VersionMajor, VersionMinor, Version, HelpLink, IrlInfoAbout, Comments, installDate }
            $apps1 = $apps1 += Invoke-command -computer $Server { Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object Displayname, Publisher, Displayversion, VersionMajor, VersionMinor, Version, HelpLink, IrlInfoAbout, Comments, installDate }
   
   
          

            foreach ($app in $Apps1 | Select-Object | Where-Object { $_.Displayname.length -gt 2 })
            {
            
               
                #Outputting it as an array.          
                [PSCustomObject]@{
                    ServerName     = $Server
                    Displayname    = $app.Displayname
                    Publisher      = $app.Publisher
                    Displayversion = $app.Displayversion
                    #    VersionMajor   = $app.VersionMajor
                    #    VersionMinor   = $app.VersionMinor
                    #    Version        = $app.Version
                    #    HelpLink       = $app.HelpLink
                    #    IrlInfoAbout   = $app.IrlInfoAbout
                    #    Comments       = $app.Comments
                    #    installDate    = $app.installDate
                }
            }
  
        }
    }
}




   

function Get-OS_from_VMWareDump
{
    <#
 #... gets the OS from AD for guests in VMWare dump file
.SYNOPSIS
    gets the OS from AD for guests in VMWare dump file
.DESCRIPTION
    gets the OS from AD for guests in Vmwarebump dump file.
    Logins into ADCS, PRDS,DVDS,DESQLD domains as admin (Prompted) 
    and gets all the Servers and then matches the OS to the name from the VMware Dump file.

.PARAMETER $VMWarefile
    Path to the VMWare dump file

.PARAMETER PRDSCREDS
    gets Creds for domain
.PARAMETER desqldcreds
    gets Creds for domain
.PARAMETER acdscreds
    gets Creds for domain
.PARAMETER DVDSCreds
    gets Creds for domain

.PARAMETER InputObject
    Specifies the object to be processed.  You can also pipe the objects to this command.
.EXAMPLE
    C:\PS>Get-OS_from_VMWareDump -VMWarefile C:\tmp\vmwExport.csv |export-csv c:\tmp\vmware_WithOS.csv
    C:\PS>Get-OS_from_VMWareDump -VMWarefile C:\tmp\vmwExport.csv
    Example of how to use this cmdlet

.OUTPUTS
    Output from this cmdlet is an arry of the Results. 
.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    15 June 2021
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           16 June 2021         Lawrence       Initial Coding

#>
     
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true,
            HelpMessage = 'Enter The path to Nutanix export file')]
        $VMWarefile,
        [Parameter(Mandatory = $false)]
        $PRDSCREDS = (get-credential -Message "Enter your PRDS Admin Creds"  ),

        [Parameter(Mandatory = $false)]
        $desqldcreds = (get-credential -Message "Enter your DESQLD Admin Creds" ),

        [Parameter(Mandatory = $false)]
        $acdscreds = (get-credential -Message "Enter your ACDS Admin Creds" ),

        [Parameter(Mandatory = $false)]
        $dvdscreds = (get-credential -Message "Enter your DVDS Admin Creds" )
     
    )
    
    
    begin 
    {
        Write-host "Please wait while I get the following Information. `nThis may take a few Min." -ForegroundColor cyan
        write-host '   Getting Servers from PRDS' -ForegroundColor green
        $servers = Get-ADComputer -Filter 'operatingsystem -like "*server*" ' -server prds.qldpol -Credential $PRDSCREDS -Properties OperatingSystem , OperatingSystemServicePack, OperatingSystemVersion, CanonicalName

        write-host '   Getting Servers from DESQLD' -ForegroundColor green
        $servers += Get-ADComputer -Filter 'operatingsystem -like "*server*" ' -server desqld.internal  -Credential $desqldcreds -Properties OperatingSystem , OperatingSystemServicePack, OperatingSystemVersion, CanonicalName

        write-host '   Getting Servers from DVDS' -ForegroundColor green
        $servers += Get-ADComputer -Filter 'operatingsystem -like "*server*" ' -Server dvds.devpol -Credential $dvdscreds -Properties OperatingSystem , OperatingSystemServicePack, OperatingSystemVersion, CanonicalName

        write-host '   Getting Servers from ACDS' -ForegroundColor green
        $servers += Get-ADComputer -Filter 'operatingsystem -like "*server*" ' -Server edwqpsaddcac01.acds.accpol  -Credential $acdscreds  -Properties OperatingSystem , OperatingSystemServicePack, OperatingSystemVersion, CanonicalName

             
    }
    
    process 
    {
            
        try 
        {
                 
            $vmwareFilea = import-csv $VMWarefile | select 'VM uuid', 'VM', 'Power State'

        }
        catch 
        {
                 
            Write-Host 'ERROR : $(_.Exception.Message)' -ForegroundColor Magenta
        }

        foreach ($server in $vmwarefilea)
        {
            #    write-host $server.'vm Name'
        
            foreach ($line in  $servers)
            {
                if ($line.name -match $server.vm)
                {
                    $OS_EOL = ""
                    If ($line.OperatingSystem -match 'Windows Server 2003') { $OS_EOL = "14-07-2015" }
                    If ($line.OperatingSystem -match 'Windows Server 2008') { $OS_EOL = "14-01-2020" }
                    If ($line.OperatingSystem -match 'Windows Server 2012') { $OS_EOL = "10-10-2023" }
                    If ($line.OperatingSystem -match 'Windows Server 2016') { $OS_EOL = "01-12-2027" }
                    If ($line.OperatingSystem -match 'Windows Server 2019') { $OS_EOL = "01-09-2029" }
                    
 
                    [PSCustomObject] @{"VM UUID"     = $server.'VM UUID'
                        "VM"                         = $server.VM
                        "OperatingSystem"            = $line.OperatingSystem
                        "OperatingSystemVersion"     = $Line.OperatingSystemVersion
                        "OperatingSystemServicePack" = $Line.OperatingSystemServicePack
                        "OperatingSystemHotfix"      = $line.OperatingSystemHotfix
                        "OSEndofSupportDate"         = $OS_EOL
                        "CanonicalName"              = $line.CanonicalName
                    }

                }
 
            }
        
        }
    }

}
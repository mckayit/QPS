function Get-AllServerOS_from_Domains
{
    <#
 #... gets the OS from AD for DES,PRDS,ACDS,DVDS
.SYNOPSIS
    gets the OS from AD for DES,PRDS,ACDS,DVDS
.DESCRIPTION
    gets the OS from AD for DES,PRDS,ACDS,DVDS
    Logins into ADCS, PRDS,DVDS,DESQLD domains as admin (Prompted) 
    and gets all the Servers and then matches the OS which have a name of Server.

.PARAMETER $nutanixfile
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
    C:\PS>Get-OS_from_VMWareDump -nutanixfile C:\tmp\vmwExport.csv |export-csv c:\tmp\vmware_WithOS.csv
    C:\PS>Get-OS_from_VMWareDump -nutanixfile C:\tmp\vmwExport.csv
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
        $nutanixfile,
        [Parameter(Mandatory = $false)]
        $PRDSCREDS = (get-credential -Message "Enter your PRDS Admin Creds"  )  ,

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

        write-host '   Getting Servers from DESQLD' -ForegroundColor green
        $DESQLD = Get-ADComputer  -Filter 'operatingsystem -like "*server*" ' -server desqld.internal  -Credential $desqldcreds -Properties OperatingSystem , OperatingSystemVersion, IPv4Address, whenchanged
        write-host "found " $desqld.count
        write-host '   Getting Servers from DVDS' -ForegroundColor green
        $DVDS = Get-ADComputer  -Filter 'operatingsystem -like "*server*" ' -Server dvds.devpol -Credential $dvdscreds -Properties OperatingSystem , OperatingSystemVersion, IPv4Address, whenchanged
        write-host "found " $dvds.count

        write-host '   Getting Servers from ACDS' -ForegroundColor green
        $acds = Get-ADComputer  -Filter 'operatingsystem -like "*server*" ' -Server edwqpsaddcac01.acds.accpol  -Credential $acdscreds  -Properties OperatingSystem , OperatingSystemVersion, IPv4Address, whenchanged
        write-host "found " $acds.count

        write-host '   Getting Server Names from PRDS' -ForegroundColor green
        $names = Get-ADComputer -Filter 'operatingsystem -like "*server*" ' -server prds.qldpol -Credential $PRDSCREDS | Select-Object -ExpandProperty name
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
            Get-ADComputer -Identity $name -server prds.qldpol -Credential $PRDSCREDS -Properties OperatingSystem , OperatingSystemVersion, IPv4Address, whenchanged
        }
        write-host "found prds B " $prds.count
        #        $prds = Get-ADComputer -ResultPageSize 999999 -Filter 'operatingsystem -like "*server*" ' -server prds.qldpol -Credential $PRDSCREDS -Properties OperatingSystem , OperatingSystemServicePack, OperatingSystemVersion, CanonicalName, IPv4Address, whenchanged
        #write-host "found " $prds.Count

        $systems = $prds
        $Systems += $acds
        $Systems += $dvds
        $Systems += $desqld
        $global:allSystems = $systems

    }
}
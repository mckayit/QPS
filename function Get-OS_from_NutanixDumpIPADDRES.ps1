function Get-OS_from_NutanixDumpIPADDRESS
{
    <#
 #... gets the OS from AD for guests in VMWare dump file
.SYNOPSIS
    gets the OS from AD for guests in VMWare dump file
.DESCRIPTION
    gets the OS from AD for guests in Vmwarebump dump file.
    Logins into ADCS, PRDS,DVDS,DESQLD domains as admin (Prompted) 
    and gets all the Servers and then matches the OS to the name from the VMware Dump file.

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
        $DESQLD = Get-ADComputer  -Filter 'operatingsystem -like "*server*" ' -server desqld.internal  -Credential $desqldcreds -Properties OperatingSystem , OperatingSystemServicePack, OperatingSystemVersion, CanonicalName, IPv4Address, whenchanged
        write-host "found " $desqld.count
        write-host '   Getting Servers from DVDS' -ForegroundColor green
        $DVDS = Get-ADComputer  -Filter 'operatingsystem -like "*server*" ' -Server dvds.devpol -Credential $dvdscreds -Properties OperatingSystem , OperatingSystemServicePack, OperatingSystemVersion, CanonicalName, IPv4Address, whenchanged
        write-host "found " $dvds.count

        write-host '   Getting Servers from ACDS' -ForegroundColor green
        $acds = Get-ADComputer  -Filter 'operatingsystem -like "*server*" ' -Server edwqpsaddcac01.acds.accpol  -Credential $acdscreds  -Properties OperatingSystem , OperatingSystemServicePack, OperatingSystemVersion, CanonicalName, IPv4Address, whenchanged
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
            Get-ADComputer -Identity $name -server prds.qldpol -Credential $PRDSCREDS -Properties OperatingSystem , OperatingSystemServicePack, OperatingSystemVersion, CanonicalName, IPv4Address, whenchanged
        }
        write-host "found prds B " $prds.count
        #        $prds = Get-ADComputer -ResultPageSize 999999 -Filter 'operatingsystem -like "*server*" ' -server prds.qldpol -Credential $PRDSCREDS -Properties OperatingSystem , OperatingSystemServicePack, OperatingSystemVersion, CanonicalName, IPv4Address, whenchanged
        #write-host "found " $prds.Count

        $systems = $prds
        $Systems += $acds
        $Systems += $dvds
        $Systems += $desqld
        $global:allSystems = $systems

        $line = $null
        $vmserver = $null
             
    }
    
    process
    {
            
        try 
        {
                 
            $nutanixlist = import-csv $nutanixfile 

        }
        catch 
        {
                 
            Write-Host 'ERROR : $(_.Exception.Message)' -ForegroundColor Magenta
        }
        $systems = $global:allSystems
        foreach ($vmserver in $nutanixlist)
        {
            #  write-host $vmserver.vm
        
            foreach ($system in  $systems)
            {
            
                $VMIPall = ($($vmserver)."IP Address(es)") 
                $VMIPS = ($($vmserver)."IP Address(es)") -split (",")
       
                $systemip = $($system).IPv4Address

                # WRITE-HOST  $systemip -ForegroundColor CYAN
                foreach ($vmip1 in $VMIPS)
                {        
                    if ($vmip1 -eq $systemip)
                    {
                        $vmservername = $($nutanixlist)."vm name"
             
                        $OS_EOL = ""
                        If ($system.OperatingSystem -match '2003') { $OS_EOL = "14-07-2015" }
                        If ($system.OperatingSystem -match '2008') { $OS_EOL = "14-01-2020" }
                        If ($system.OperatingSystem -match '2012') { $OS_EOL = "10-10-2023" }
                        If ($system.OperatingSystem -match '2016') { $OS_EOL = "01-12-2027" }
                        If ($system.OperatingSystem -match '2019') { $OS_EOL = "01-09-2029" }
                    
 
                        [PSCustomObject] @{
                            "VM UUID"                    = $vmserver.'VM UUID'
                            "namefromVMList"             = $vmservername
                            "VM"                         = $system.name
                            'ADIPv4Address'              = $systemip
                            "VMIPAddressMatch"           = $vmip1
                            VMIPADdress_all              = $VMIPall
                            "OperatingSystem"            = $system.OperatingSystem
                            "OperatingSystemVersion"     = $system.OperatingSystemVersion
                            "OperatingSystemServicePack" = $system.OperatingSystemServicePack
                            "OperatingSystemHotfix"      = $system.OperatingSystemHotfix
                            "OSEndofSupportDate"         = $OS_EOL
                            "OSWhenChanged"              = $system.whenchanged
                            "CanonicalName"              = $system.CanonicalName
                        }

             
 
                    }
                }
                
            }
                
        }
    }
   
}

measure-command { Get-OS_from_VMWareDumpIPADDRESS -nutanixfile D:\tmp\VMExport.csv | export-csv d:\tmp\vmwareexport_withOS_Basedon_IP.csv -NoTypeInformation }


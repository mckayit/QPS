<# 
    ********************************************************************************  
    *                                                                              *  
    *        This script Loads all my common f u n c t i o n s                     *
    *                                                                              *  
    ********************************************************************************    
    Note.
    All the standard Functions I use loaded.
      
    *******************
    Copyright Notice.
    *******************
    This Program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.

    test-exchangeonlineconnected

    Author:
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty Ltd
    
    Date:   26  March 2018
   

    ******* Update Version number below when a change is done.*******

    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           15 Sept 20121       Lawrence       Initial Coding



    #>

$version = "Version 0.0.31"
Write-Host $version -ForegroundColor Green


#$global:Functisloaded = "YES"

#Write-host "`nLoading $PSCommandPath"  -ForegroundColor green -BackgroundColor Red
$FormatEnumerationLimit = -1

Function get-help1 
{
    <#
    #... Displays the Functions in QHSupport.
    Syntax   Get-help1

    This will generate the report for all of the batchname starting with Batch14*
#>

    param
    (
	
    )
   	
    BEGIN
    {
        Write-Host $version -ForegroundColor Green
    }

    PROCESS
    {
        #... Display Functions with the comments
        $DDSPLAY = ""
        #Reads in the Current Powershell script file
        $DDSPLAY = get-content $PSCommandPath 

        foreach ($line in $DDSPLAY)
        {
            if ($line.Trim().StartsWith('Function', "CurrentCultureIgnoreCase") -or $line.Trim().Startswith('#...', "CurrentCultureIgnoreCase"))
            {
                $1 = $line

                if ($1.Trim().StartsWith('#...', "CurrentCultureIgnoreCase"))
                {
                
                    #Removes the first 4 char's Eg "#... "
                    $linedes = $line.trim().substring(4)
                    Write-Host $linedes -f Gray -NoNewline
                }
                Elseif (!($1.Trim().Startswith('#...')))
                {
                    Write-host ''
                }
                if ($1.Trim().StartsWith('Function', "CurrentCultureIgnoreCase"))
                {
                    $linelong = $line + "                                               "
                
                    #makes the Line length to be 50 so the comments all line up. Fills it up with a space.
                    $line = $linelong.substring(0, 50)
                    Write-host "  $line" -f green -NoNewline
                }
            }
        }
        
    }
    END
    {
        Write-Output  ""
    }
}
Function sync-QPSmodule
{
    #... copies QHO365MigrationOps.psm1 module to 
    [string]$sourcefiles = "C:\Users\904223\OneDrive - Queensland Police Service\Github\QPS\QPSsupport.psm1"
    [string]$destinationDir = 'c:\Windows\System32\WindowsPowerShell\v1.0\Modules\QPSSupport\'
    copy-item -force -Recurse $sourcefiles -Destination $destinationDir
    
    Remove-Module QPSSupport -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 10
    import-module QPSSupport
}
function start-DomainCleanup_old_Computer_objects
{
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
}
function Test-server
{
    #... copies Tests Server to se eif it is up or not and if can be Decomm'd
    <#
    here is how i get the Details for the different domains

            Write-host "Please wait while I get the following Information. `nThis may take a few Min." -ForegroundColor cyan

        write-host '   Getting Servers from DESQLD' -ForegroundColor green
        $DESQLD = Get-ADComputer  -Filter 'operatingsystem -like "*server*" ' -server desqld.internal  -Credential $desqldcreds -Properties description, OperatingSystem , OperatingSystemServicePack, OperatingSystemVersion, CanonicalName, IPv4Address, whenchanged , PasswordExpired , PasswordLastSet , LastLogonDate | select CanonicalName, DistinguishedName, DNSHostName, Enabled, IPv4Address, Name, Description, OperatingSystem, OperatingSystemServicePack, OperatingSystemVersion, whenchanged, PasswordExpired , PasswordLastSet , LastLogonDate
        write-host "found " $desqld.count
        write-host '   Getting Servers from DVDS' -ForegroundColor green
        $DVDS = Get-ADComputer  -Filter 'operatingsystem -like "*server*" ' -Server dvds.devpol -Credential $dvdscreds -Properties description, OperatingSystem , OperatingSystemServicePack, OperatingSystemVersion, CanonicalName, IPv4Address, whenchanged , PasswordExpired , PasswordLastSet , LastLogonDate | select CanonicalName, DistinguishedName, DNSHostName, Enabled, IPv4Address, Name, Description, OperatingSystem, OperatingSystemServicePack, OperatingSystemVersion, whenchanged, PasswordExpired , PasswordLastSet , LastLogonDate
        write-host "found " $dvds.count

        write-host '   Getting Servers from ACDS' -ForegroundColor green
        $acds = Get-ADComputer  -Filter 'operatingsystem -like "*server*" ' -Server edwqpsaddcac01.acds.accpol  -Credential $acdscreds  -Properties description, OperatingSystem , OperatingSystemServicePack, OperatingSystemVersion, CanonicalName, IPv4Address, whenchanged , PasswordExpired , PasswordLastSet , LastLogonDate | select CanonicalName, DistinguishedName, DNSHostName, Enabled, IPv4Address, Name, Description, OperatingSystem, OperatingSystemServicePack, OperatingSystemVersion, whenchanged, PasswordExpired , PasswordLastSet , LastLogonDate
        write-host "found " $acds.count

        write-host '   Getting Server Names from PRDS' -ForegroundColor green
        $names = Get-ADComputer -Filter 'operatingsystem -like "*server*" ' -server prds.qldpol  | Select-Object -ExpandProperty name
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
            Get-ADComputer -Identity $name -server prds.qldpol  -Properties description, OperatingSystem , OperatingSystemServicePack, OperatingSystemVersion, CanonicalName, IPv4Address, whenchanged , PasswordExpired , PasswordLastSet , LastLogonDate | select CanonicalName, DistinguishedName, DNSHostName, Enabled, IPv4Address, Name, Description, OperatingSystem, OperatingSystemServicePack, OperatingSystemVersion, whenchanged, PasswordExpired , PasswordLastSet , LastLogonDate
        }
        write-host "found prds B " $prds.count

    
    #>

    [CmdletBinding()]
    param (
        $Servers,
        [ValidateSet('PRDS', 'DVDS', 'ACDS', 'DESQLD')]
        [String]$DNS_ = 'PRDS',
        [String]$HQRFile2Import
    )
    
    begin
    {
        If ($dns_ -eq 'PRDS') { $dnsserver = '164.112.162.142' } 
        If ($dns_ -eq 'DVDS') { $dnsserver = '164.112.162.250' } 
        If ($dns_ -eq 'ACDS') { $dnsserver = '164.112.162.230' } 
        If ($dns_ -eq 'DESQLD') { $dnsserver = '10.2.109.50' } 
   
    
        try
        {
            $hqr = import-csv $HQRFile2Import
        }
        catch
        {
            Write-host 'Failed to import HQR   `nData will not be fully populated' -forgroundcolor Cyan
        }
    }
    process
    {
        $i = 1
        foreach ($server in $servers)
        {
            $paramWriteProgress = @{

                
                Activity        = 'Getting  info for $($server.name)'
                Status          = "Processing [$i] of [$($servers.Count)] Servers"
                PercentComplete = (($i / $servers.Count) * 100)
                                
            }
                            
            Write-Progress @paramWriteProgress
                        
            $i++
           
            $InDNS = "$false"   

            if ($Dns = resolve-dnsname $server.dnshostname -server $dnsserver -ErrorAction SilentlyContinue | select-string -pattern 'Answer') 
            {
                $InDNS = "True"  
            }
         

            if (test-connection $server.dnshostname -Quiet -count 2) { $PINGname = "True" }
            else
            {
                $PINGname = "False"
            }
            try
            {
                if (test-connection $dns.IPAddress -Quiet -count 2 -ErrorAction SilentlyContinue) { $PINGableIP = "True" }
                else
                {
                    $PINGABLEIP = "False"
                }
            }
            catch { $PINGABLEIP = "False" }

            $decom = ""
            If (-not ($dns.IPAddress) ) { $Decom = "Can be Decommed as does not exist in DNS or Pingable" }

            # checking HQR
            $hqr_ = $hqr | Where-Object { $server.name -contains $_."server name" } 
      

            [PSCustomObject] @{
                "Servername"                                = $server.dnshostname
                "CanonicalName"                             = $server.CanonicalName
                "OperatingSystem"                           = $server.OperatingSystem
                "Description"                               = $server.description
                "Is_Pingable_by_Name"                       = $PINGname
                "Is_Pingable_by_IP"                         = $PINGABLEIP
                "Found_IN_DNS"                              = $inDNS
                "IP_Found_IN_AD"                            = $server.IPv4Address
                "IP_found_IN_DNS"                           = $dns.IPAddress
                "Server_IN_AD"                              = "True"
                'Last_time_server_spoke_2_Domaincontroller' = $server.whenchanged
                'PasswordExpired'                           = $server.PasswordExpired 
                'AD_Object_PasswordLastSet'                 = $server.PasswordLastSet
                'AD_Object_LastLogonDate'                   = $server.LastLogonDate
                'Decom'                                     = $decom
                BusinessOwner                               = $hqr_.'Business Owner - Name'
                Support_Group                               = $hqr_.'Support - Group'
                SystemName                                  = $hqr_.'System - Name'
            }


        }
        
    }
    end
    {
        
    }    
    #    write-host "Command PS `n`n Test-server -Servers <VAR> -DNS_ prds -HQRFile2Import D:\tmp\hqr.csv  |export-csv d:\tmp\prds_servers_2_b_decommed.csv -NoTypeInformation'" -ForegroundColor green
       
}

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
    end
    {
        #Outputs to screen
        Write-host $global:allSystems
    }
}

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
    $servers = get-adcomputer -Filter { Name -like "*sql*" } 
    $i
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


            foreach ($Instal in $installed)
            {
                if ($instal.displayname -like "Microsoft SQL Server * Setup (English)")
                {

                    [PSCustomObject] @{
                        servername = $server.name
                        SQLAPP     = $instal.displayname

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

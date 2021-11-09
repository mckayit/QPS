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
            Get-ADComputer -Identity $name -server prds.qldpol  -Properties OperatingSystem , OperatingSystemVersion, IPv4Address, whenchanged
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

            $installed = invoke-command -computername $server.dnshostname { Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.displayname -like "Microsoft SQL Server * Setup (English)" } | Sort-Object displayname | Select-Object DisplayName, Publisher } -ErrorAction SilentlyContinue

            #sql instance
       

            foreach ($Instal in $installed)
            {
                if ($instal.displayname -like "Microsoft SQL Server * Setup (English)")
                {


                    $pspath = invoke-command -computername $server.dnshostname { (get-itemproperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\*').pschildname }
                    foreach ($ps in $pspath)
                    {

                  
                        $ps1 = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\$ps"
                        


                        $SQLInstances = invoke-command -computername $server.dnshostname -scriptblock { gi ($using:ps1) }
                        [PSCustomObject] @{
                            "servername"        = $server.name
                            "SQLAPP"            = $instal.displayname
                            "Domain"            = $env:USERDNSDOMAIN
                            "Instance name"     = $SQLInstances.pschildname
                            "Instance Property" = $SQLInstances.property 
                        }
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


function get-dcinfo
{
    <#
     #... Get all the DC's and os and functional Level.
    .SYNOPSIS
   Get all the DC's and os and functional Level.
    
    .DESCRIPTION
    Get all the DC's and os and functional Level.
    
    .EXAMPLE
    get-dcinfo
    
    .NOTES
    General notes
    #>
    
    $DCs = (get-addomain).ReplicaDirectoryServers | ForEach { Get-ADDomainController -Identity $_ | Select-Object name, OperatingSystem }
    #$DCs = Get-ADDomainController -filter * | Select-Object name, OperatingSystem
    $domainname = $env:USERDNSDOMAIN
    $DomainMode = Get-ADDomain | select -ExpandProperty DomainMode
    $forestLevel = Get-ADForest | select -ExpandProperty ForestMode

    foreach ($DC in $DCs)
    {

        [PSCustomObject] @{
            servername      = $dc.name
            OS              = $dc.OperatingSystem
            DomainName      = $env:USERDNSDOMAIN
            FunctionalLevel = $DomainMode 
            ForestLevel     = $forestLevel

        }
    }

}


function get-disksizeWithMountpoints
{
    <#
 #... get All Moiuntpoints on a system
.SYNOPSIS
    get all mount points of disks on a system
.DESCRIPTION
    displays all disks sizings that are Fixed disk type. (3)  
    this will also show the disks that have mountpoints and where thewy are mounted.. 
    
.PARAMETER one
    Specifies Pram details.
.PARAMETER two
    Specifies Pram details
.PARAMETER InputObject
    Specifies the object to be processed.  You can also pipe the objects to this command.
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
     
    Date:    9 aug 2021
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           9 Aug 2021         Lawrence       Initial Coding

#>
     
  

    begin 
    {
        #Setting the info on how the sizing are done.
                  
    }
    
    process 
    {

        try
        {
            $volumes = Get-WmiObject win32_volume -Filter "DriveType='3'" | Sort-Object DriveLetter # -Descending


            foreach ($volume in $volumes)
            {
                $TotalGB = [math]::round(($volume.Capacity / 1gb), 2) 
            
                $FreeGB = [math]::round(($volume.FreeSpace / 1Gb), 2) 
        
                $FreePerc = [math]::round(((($volume.FreeSpace / 1GB) / ($volume.Capacity / 1GB)) * 100), 0) 
    
                [PSCustomObject] @{
                    Name            = $volume.name
                    Label           = $volume.label
                    DriveLetter     = $volume.driveletter
                    FileSystem      = $volume.filesystem
                    "Capacity(GB)"  = $TotalGB
                    "FreeSpace(GB)" = $FreeGB
                    "Free(%)"       = $FreePerc
                }
            }
        }
    
  

        catch 
        {
            Write-Host 'ERROR : $(_.Exception.Message)' -ForegroundColor Magenta
        }

    }

    end 
    {
            
    }

}


function Get-RandomPassword
{
    <#
    Random password creator.
    Does not use the following Char.   1 I l O 0   Due to to hard to read and mis-typing

creates a 16 char password by default unless you enter more. 
Will not create Password less than 8 Chars'

    #>


    Param(
        [Parameter(mandatory = $false)]
        [int]$Length = 16
    )
    Begin
    {
        if ($Length -lt 8)
        {
            Write-host  "Password is to short   must be over 8 char." -ForegroundColor white -BackgroundColor Red
            break
        }
        $Numbers = 2..9
        $LettersLower = 'abcdefghijkmnopqrstuvwxyz'.ToCharArray()
        $LettersUpper = 'ABCEDEFHJKLMNPQRSTUVWXYZ'.ToCharArray()
        $Special = '!@#$%^&*()=+[{}]/?<>'.ToCharArray()

        # For the 4 character types (upper, lower, numerical, and special)
        $N_Count = [math]::Round($Length * .2)
        $L_Count = [math]::Round($Length * .4)
        $U_Count = [math]::Round($Length * .2)
        $S_Count = [math]::Round($Length * .2)
    }
    Process
    {
        $Pswrd = $LettersLower | Get-Random -Count $L_Count
        $Pswrd += $Numbers | Get-Random -Count $N_Count
        $Pswrd += $LettersUpper | Get-Random -Count $U_Count
        $Pswrd += $Special | Get-Random -Count $S_Count

        # If the password length isn't long enough (due to rounding), add X special characters
        # Where X is the difference between the desired length and the current length.
        if ($Pswrd.length -lt $Length)
        {
            $Pswrd += $Special | Get-Random -Count ($Length - $Pswrd.length)
        }

        # Lastly, grab the $Pswrd string and randomize the order
        $Pswrd = ($Pswrd | Get-Random -Count $Length) -join ""
    }
    End
    {
        $Pswrd
    }
}

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

function get-shareinfo
{
    #get all Shares
    $shares = Get-WmiObject -Class Win32_Share 
    $shareList = New-Object -TypeName System.Collections.ArrayList

    foreach ($share in $shares)
    {
  
        #excluding default shares   
        if (($share.Name -notmatch '(?im)^[a-z]{1,1}\$') -and ($share.Name -notmatch '(?im)^[admin]{5,5}\$') -and ($share.Name -notmatch '(?im)^[ipc]{3,3}\$') -and ($share.Name -notmatch '(?im)^[print]{5,5}\$') )
        {      
    
            $shareAccessInfo = ''
            $ntfsAccessInfo = ''    
    
            #extract permissions from the current share
            $fileAccessControlList = Get-Acl -Path $($share.Path) | Select-Object -ExpandProperty Access | Select-Object -Property FileSystemRights, AccessControlType, IdentityReference    
    
            #excluding uncritical information as Builtin Accounts as Administratrators, System, NT Service and Trusted installer
            foreach ($fileAccessControlEntry in $fileAccessControlList)
            {
                if (($fileAccessControlEntry.FileSystemRights -notmatch '\d') -and ($fileAccessControlEntry.IdentityReference -notmatch '(?i)Builtin\\Administrators|NT\sAUTHORITY\\SYSTEM|NT\sSERVICE\\TrustedInstaller'))
                {      
                    $ntfsAccessInfo += "$($fileAccessControlEntry.IdentityReference); $($fileAccessControlEntry.AccessControlType); $($fileAccessControlEntry.FileSystemRights)" + ' | '  
                }
            } #END foreach ($fileAccessControlEntry in $fileAccessControlList)

            $ntfsAccessInfo = $ntfsAccessInfo.Substring(0, $ntfsAccessInfo.Length - 3)
            $ntfsAccessInfo = $ntfsAccessInfo -replace ',\s?Synchronize', ''   
    
            #getting share permissions   
            $shareSecuritySetting = Get-WmiObject -Class Win32_LogicalShareSecuritySetting -Filter "Name='$($share.Name)'"               
            $shareSecurityDescriptor = $shareSecuritySetting.GetSecurityDescriptor()
            $shareAcccessControlList = $shareSecurityDescriptor.Descriptor.DACL          
    
            #converting share permissions to be human readable
            foreach ($shareAccessControlEntry in $shareAcccessControlList)
            {
    
                $trustee = $($shareAccessControlEntry.Trustee).Name      
                $accessMask = $shareAccessControlEntry.AccessMask
      
                if ($shareAccessControlEntry.AceType -eq 0)
                {
                    $accessType = 'Allow'
                }
                else
                {
                    $accessType = 'Deny'
                }
        
                if ($accessMask -match '2032127|1245631|1179817')
                {          
                    if ($accessMask -eq 2032127)
                    {
                        $accessMaskInfo = 'FullControl'
                    }
                    elseif ($accessMask -eq 1179817)
                    {
                        $accessMaskInfo = 'Read'
                    }
                    elseif ($accessMask -eq 1245631)
                    {
                        $accessMaskInfo = 'Change'
                    }
                    else
                    {
                        $accessMaskInfo = 'unknown'
                    }
                    $shareAccessInfo += "$trustee; $accessType; $accessMaskInfo" + ' | '
                }            
    
            } #END foreach($shareAccessControlEntry in $shareAcccessControlList)
    
       
            if ($shareAccessInfo -match '|')
            {
                $shareAccessInfo = $shareAccessInfo.Substring(0, $shareAccessInfo.Length - 3)
            }               
    
            #putting extracted information together into a custom object    
            $myShareHash = @{'Name' = $share.Name }
            $myShareHash.Add('FileSystemSPath', $share.Path )       
            $myShareHash.Add('Description', $share.Description)        
            $myShareHash.Add('NTFSPermissions', $ntfsAccessInfo)
            $myShareHash.Add('SharePermissions', $shareAccessInfo)
            $myShareObject = New-Object -TypeName PSObject -Property $myShareHash
            $myShareObject.PSObject.TypeNames.Insert(0, 'MyShareObject')  
    
            #store the custom object in a list    
            $null = $shareList.Add($myShareObject)
  
        } #END if (($share.Name -notmatch '(?im)^[a-z]{1,1}\$') -and ($share.Name -notmatch '(?im)^[admin]{5,5}\$') -and ($share.Name -notmatch '(?im)^[ipc]{3,3}\$') )

    } #END foreach ($share in $shares)
    $shareList
} }


Function remove-vmwaretoolsmanually
{
    <#
 #... removes VMware tools and informational error reporting
.SYNOPSIS
    removes VMware tools and informational error reporting
.DESCRIPTION
    Manually removes VMware tools and informational error reporting
    in the application Event log on 2019 servers you can get lots of informational events relating to windows reporting (Errors)
    The reason this happens is due to an application has crashed and logged data to C:\ProgramData\Microsoft\Windows\WER\ReportQueue
    the eventlog seems to poole this loaction regularry and contuniously report the errors and fill the eventlog.
    Removing all the sub folders will stop these errors from contuniously reporting.
.PARAMETER one
    None

    .EXAMPLE
    C:\PS>
    
.OUTPUTS
    Output from this cmdlet (if any)
.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    11 Nov 2021
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           10 Nov 2021         Lawrence       Initial Coding

#>

    function remove-QPSFolders
    {
        <#
     #... Removes DIR and REG keys recursively 
    .SYNOPSIS
    Removes DIR and REG keys recursively 
    
    .DESCRIPTION
    this function will remove all dir's or Regkey's recursivly.  
    There is ** NO ** Confirmation and will just do it.
    So be very carefull with what you add as a Var "Foldername"

  
    .EXAMPLE
    remove-QPSFolders -foldername  '<Folder path>'
    remove-QPSFolders -foldername  'C:\Program Files\VMware'

    remove-QPSFolders -foldername  '<Reg key>'
    remove-QPSFolders -foldername  "HKLM:\SOFTWARE\VMware, Inc."
    
    .NOTES
    Note 
        this script needs to be run twice (2x) to cleanup correctly.

        

    #>


        [CmdletBinding()]
        param( 
            [Parameter(Mandatory = $true)]
            $FolderName 
        
        )
        Begin
        {
     

        }

        Process
        {    
            if (Test-Path $FolderName)
            {
                Try
                {
                
                    Remove-Item $FolderName -Force -Recurse
                    Write-Host "$foldername has been removed." -ForegroundColor Green
                }
                catch
                {
                    Write-Host "$FolderName Failed to be removed" -ForegroundColor  red 
                }
            }
            else
            {
                Write-Host "$FolderName Doesn't Exists"  -ForegroundColor  Cyan 

            }
        }
        end
        {

        }

    }

    <# 
    Removes all the Windows Error Reporting folder contents.
    This removes the error from appearing in the Eventlog.
    this is required as the evenlog reports this as informational.
#>
    function remove-informationaleventlogforerrores
    {
        <#
 #... removes informationalError from eventlog 
.SYNOPSIS
    removes informationalError from eventlog
.DESCRIPTION
    in the application Event log on 2019 servers you can get lots of informational events relating to windows reporting (Errors)
    The reason this happens is due to an application has crashed and logged data to C:\ProgramData\Microsoft\Windows\WER\ReportQueue
    the eventlog seems to poole this loaction regularry and contuniously report the errors and fill the eventlog.
    Removing all the sub folders will stop these errors from contuniously reporting.
.PARAMETER one
    None

    .EXAMPLE
    C:\PS>remove-informationaleventlogforerrores
    
    
.OUTPUTS
    Output from this cmdlet (if any)
.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    11 Nov 2021
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           10 Nov 2021         Lawrence       Initial Coding

#>
     
        begin 
        {
            $WERDirs = (Get-ChildItem -Directory C:\ProgramData\Microsoft\Windows\WER\ReportQueue).FullName      
        }
    
        process 
        {
            foreach ($WER in $WERDIRs)
            { 
                try 
                {
                    remove-QPSFolders -foldername $wer 
                }
    
                catch 
                {

                }
            }
        }

    }
    
    end
    {
            
    }

    # Removes the Informational reporting from eventlog
    remove-informationaleventlogforerrores

    # removes the VMware tools DIR's
    remove-QPSFolders -foldername  'C:\Program Files\VMware'
    remove-QPSFolders -foldername  'C:\ProgramData\VMware' 

    # removes VMware Tools reg keys. 
    remove-QPSFolders -foldername  "HKLM:\SOFTWARE\VMware, Inc."
    remove-QPSFolders -foldername  "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{1D060220-2A64-4153-A6F5-C43B95C3BFC7}"
    remove-QPSFolders -foldername  "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Compatibility Assistant\Store" 
    remove-QPSFolders -foldername  "HKLM:\SOFTWARE\Classes\Installer\Products\022060D146A235146A5F4CB3593CFB7C"
    remove-QPSFolders -foldername  "HKLM:\SOFTWARE\Classes\TypeLib\{6B8C0665-86D9-4DC9-8D58-FABE31A495E3}"


    #reboots the system within 2 Sec
    shutdown /r /t 2
} 

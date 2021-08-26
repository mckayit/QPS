<# 
  ********************************************************************************  
  *                                                                              *  
  *  This script gets the info required to complete the INI Server QA report     *
  *                                                                              *  
  ********************************************************************************    

    Note:
            Past in ISE and run   Gives all the info needed to complete the INI Server QA Sheet


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




    Author:
            Lawrence McKay
            Lawrence@mckayit.com
            McKayIT Solutions Pty Ltd
    
    Date:   20  Nov 2015


   ******* Update Version number below when a change is done.*******

    
    History
            Version    Date               Name           Detail
            ------------------------------------------------------------------------------------
            1.0      20  Nov 2015     Lawrence       Initial Coding
            1.1      23  Nov 2015     Lawrence       Fixed up layup issues
                                                     Fixed up Features and Functions
            1.2      24  Nov 2015     Lawrence       Add number of features and functions installed to functionsinstalled.  
            1.3      25  Nov 2015     Lawrence       Added Check to see id VMware Tools are installed and running 'Function  VMwarecheck'
            1.4       2  Dec 2015     Lawrence       Added Block Size options to the Disk Report.
                                                     Added Service running as user check.  
            1.5       3  Dec 2015     Lawrence       Fixed Software results layout.
                                                     Cleaned up NLB Error on screen when NLB not installed.
            1.6       3  December     Lawrence       Fixed up Funsctiosinstalled Function.   Cleaned up lots of repatition.
            1.7       5  January      Lawrence       Fixed up bug on share permissions displayed.
            1.8      12  January      Lawrence       Added QCAD VAS folder permissions check.
            1.9      28  January      Lawrence       Added Windows Version to the Report.
            1.10      2  Feburary     Lawrence       fixed Trusted for Delagation Function
            1.11      8  Feburary     Lawrence       Fixed CPU reporting to now show more CPU info.
            1.12     18  Feburary     Lawrence       fixed issue where c:\temp does not exist by creating it at the start.
            1.13      8  April        Lawrence       Re added 2 windows features to be shown in the report.                                                                                                                  
            1.14     13  April        Lawrence       Fixed issue with inetpup logs folder getting created.        
            1.15     18  April        Lawrence       Added option to show User Domain as well as Server Domain  (To show boxes in the SGS) 
            1.16     29  April        Lawrence       Displaying Installed Apps                                                                                              
            1.17     04  May          Lawrence       Fixed bug in showing Scheduled tasks as user.  Was showing author not user.                                                                                                               
            1.17     05  May          Lawrence       Fixed display error for getting Drive permissions for A:\
            1.18     10  May          Lawrence       Updated Schduled Tasks Detection.
            1.19     16  May          Lawrence       Fixed issue with Scheduled task showing up when it should have not.
            1.20     20  MAy          Lawrence       Now showing If the Oracle Client is installed.  Eg x32 or x64.  (Not that nice  parsing the Reg keys.)
            1.21     20  MAy          Lawrence       Checking to make sure script is running in an esculated Powershell window. Eg as Administrator.
            1.22     20  MAy          Lawrence       Cleans up any remaining Installation DIR's that are left as part of the Scripted Builds used. (Only known DIR's) See Function "Cleanup-install-folder"
            1.23     20  MAy          Lawrence       Cleaned up script for outputing to the report file.
            1.24     20  MAy          Lawrence       Fixed issue re displaying Pagefile Size.
            1.25     02  June         Lawrence       Added Network Teaming results to be displayed. (added to function networksettings)
            1.26     02  June         Lawrence       Configured Display to not show NLB settings if nlb not installed.
            1.27     03  June         Lawrence       Fixed display issue with reporting of Windows Share.   
            1.28     02  June         Lawrence       Added Network Teaming results to include MAC Address and Link Speed
            1.29     09  June         Lawrence       Fixed Display issue for SMB Shares
            1.30     09  June         Lawrence       Code cleanup..
            1.31     10  June         Lawrence       Fixed Display issue with Script Version not showing in the Report
            1.32     10  June         Lawrence       Fixed Issue with Eventlog error Display. Function "eventlogerrors" 
            1.33     13  June         Lawrence       Inproved layout of report
            1.34     13  June         Lawrence       More Inproved layout of report
            1.35     16  June         Lawrence       Added Number of Logical Processors to CPU report
            1.36     17  June         Lawrence       IIS options will now not show in report if IIS is not installed
            1.37     17  June         Lawrence       Fixed Scheduled task report logic.
            1.38     17  June         Lawrence       Fixed VMware tools display output
            1.39     17  June         Lawrence       Fixed Pagefile Layout 
            1.40     17  June         Lawrence       Fixed Serives-check Layout 
            1.41     17  June         Lawrence       Fixed Applications installed Layout.
            1.42     17  June         Lawrence       Fixed / Cleaned up code Global Variables  as they were no longer used.
            1.43     20  June         Lawrence       Added option to display the OU this system is in.
            1.44     20  June         Lawrence       Fixed Display issue for Pagefile size and location
            1.45     21  June         Lawrence       Fixed Display issue for vmWARE TOOLS INSTALLED.   NEEDED TO MAKE $cpu1 vAIRABLE A GLOBAL ONE.
            1.46     23  June         Lawrence       Added the QA report now gets email to user that ran it.   Can be commented out.
            1.47     24  June         Lawrence       Fixed bug re emailing report.  now it lets u know to email report to yourself you need to run 'email-report'
            1.48     28  June         Lawrence       Fixed how the date for uptime is shown
            1.50     30  March 2021   Lawrence       Cleaned up Code and fixed a few issues.

#> 


$Global:ver = "1.50"





#checks to see if running as Admin
Clear-Host
Write-host  "`n`n  *********************************" -ForegroundColor Green
Write-host  "  *   QA Script Version is $ver   *" -ForegroundColor GREEN
Write-host  "  *********************************`n`n`n" -ForegroundColor Green
     

get-process notepad | stop-process
clear

Function Scriptver
{
    $Global:Scriptversion = " QA Script Ver $ver"
    Write-host  "`n`n       Script Version is $ScriptVersion `n`n`n" -ForegroundColor Yellow
}

function cpu
{
    Write-host "Processing..CPU settings."-ForegroundColor green
    #CPU info
    $processors = get-wmiobject -computername localhost win32_processor
    [int]$cores = 0
    [int]$sockets = 0
    [string]$test = $null
    foreach ($proc in $processors)
    {
        if ($proc.numberofcores -eq $null)
        {
            If (-not $Test.contains($proc.SocketDesignation))
            {
                $Test = $Test + $proc.SocketDesignation
                $sockets++
            }
            $cores++
        }
        else
        {
            $sockets++
            $cores = $cores + $proc.numberofcores
            $LOGProc = $logProc + $proc.numberoflogicalProcessors
        }
    }
    #“Cores: $cores, Sockets: $sockets”
    $cpu = [char]0x2551 + "    CPU info.                                                               " + [char]0x2551
    
    $procManufacturer = get-wmiobject -computername localhost win32_computersystem | select Model, Manufacturer
    $procManufacturer1 = get-wmiobject -computername localhost win32_processor | Out-String
  
    $1 = "                 Manufacturer: "
    $1 = $1 + $procManufacturer.Manufacturer | Out-String
    $2 = "                 Machine Type: "
    $2 = $2 + $procManufacturer.Model | Out-String
    $3 = "              Number of Cores: "
    $3 = $3 + $cores | Out-String

    $3a = " Number of Logical Processors: "
    $3a = $3a + $LOGProc | Out-String


    $4 = "            Number of Sockets: "
    $4 = $4 + $sockets | Out-String

    $5 = " CPU Model / Type. "
    $5 = $5 + $procManufacturer1  | Out-String
    $global:cpu1 = $1 + $2 + $3 + $3a + $4 + $5


    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $cpu  | Out-file  C:\temp\$servername.txt -append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $cpu1 | Out-file  C:\temp\$servername.txt -Append
        
}

##<#
.SYNOPSIS
#

.DESCRIPTION
Long description

.EXAMPLE
An example

.NOTES
General notes
#>
function memory
{
    # Display memory 
    Write-host "Processing..Memory settings."-ForegroundColor green
    $mem = Get-WmiObject -Class Win32_ComputerSystem 
    #$mem1 = $mem.TotalPhysicalMemory 
    $mem1 = [math]::Ceiling($mem.TotalPhysicalMemory / 1024 / 1024 / 1024)
    $memory1 = [char]0x2551 + "    Server RAM in GB                                                        " + [char]0x2551
    $memory2 = "This system has $mem1 GB RAM " 

    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $Memory1 | Out-file  C:\temp\$servername.txt -Append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $Memory2 | Out-file  C:\temp\$servername.txt -Append


}

function pagefile
{
    Write-host "Processing..Pagefile Settings."-ForegroundColor green
    #$Pagefile

    $Pagesize = (Get-WmiObject win32_pagefileusage | select @{name = "PageFile Location"; Expression = { $_.Name } },
        @{name = "Base Size(MB)"; Expression = { $_.AllocatedBaseSize } }) | ft  | Out-String -width 50
         
    $PGSize = [char]0x2551 + "    Page file settings                                                      " + [char]0x2551 
    $PGSize1 = $pagesize 
    $PGSize2 = "Pagefile set as Auto managed."
    $PGSize3 = (gwmi Win32_ComputerSystem).AutomaticManagedPagefile
    $PGSize2 = $PGSize2 + $PGSize3
    $PGSize3a = "       NOTE: It should be false"
    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $PGSize | Out-file  C:\temp\$servername.txt -Append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $PGSize1 | Out-file  C:\temp\$servername.txt -Append
    $PGSize2 | Out-file  C:\temp\$servername.txt -Append
    $PGSize3a | Out-file  C:\temp\$servername.txt -Append
}

function networksettings
{
    #network settings
    Write-host "Processing..Network settings."-ForegroundColor green
    $netwrk = [char]0x2551 + "    Network settings.                                                       " + [char]0x2551

    #$netwrk1 = Get-NetIPConfiguration -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Out-String
    $netwrk1 = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter IPEnabled=TRUE -ComputerName . | Select-Object -Property [a-z]* | fl IPAddress, IPSubnet, DefaultIPGateway, Description*, DHCPEn*, DNSDomainSuffixSearchOrder
    
    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $netwrk | Out-file  C:\temp\$servername.txt -Append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $netwrk1 | Out-file  C:\temp\$servername.txt -Append

    <#
    #network Teaming 
    if (Get-NetLbfoTeam | where name -ne "")
    {
        $teamname = (Get-NetLbfoTeam).name | Out-String

        foreach ($i in((get-netlbfoteam).name))
        {
            $Teamstatus = Get-NetAdapter (Get-NetLbfoTeamMember -team $i).name | ft -AutoSize | Out-String
        }

       
        $NicTeaming1 = [char]0x2551 + "    Network Teaming Info if set.                                            " + [char]0x2551
        $Nicteaming2 = "Network Team name: $teamname"
        $nicteaming3 = "Network Adapters in the Team:`n$teamstatus"
    

        $linetop | Out-file  C:\temp\$servername.txt -append
        $NicTeaming1 | Out-file  C:\temp\$servername.txt -append
        $linebottom | Out-file  C:\temp\$servername.txt -append
        $Nicteaming2 | Out-file  C:\temp\$servername.txt -append
        $nicteaming3 | Out-file  C:\temp\$servername.txt -append
    }
    else {}
#>
}

function domain
{
    Write-host "Processing..Domain settings."-ForegroundColor green
    #Server Domain
        
    $dom = [char]0x2551 + "    User Domain          Server Domain                                      " + [char]0x2551
    #  commented out due to ps2 $DOM2 = Get-NetIPConfiguration -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
    #  commented out due to ps2 $dom1 = "    " + $env:USERDNSDOMAIN + "            " + $DOM2.netprofile.name
    $dom1 = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter IPEnabled=TRUE -ComputerName . | Select-Object -ExpandProperty DNSDomainSuffixSearchOrder
    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $dom | Out-file  C:\temp\$servername.txt -Append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $dom1 | Out-file  C:\temp\$servername.txt -Append

}

function descript
{
    Write-host "Processing..Computer Description settings."-ForegroundColor green
    #server Description
    $pcdesc = Get-WmiObject Win32_operatingSystem 
    $desc = [char]0x2551 + "    Server Description                                                      " + [char]0x2551
    $desc1 = $pcdesc.description

    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $desc | Out-file  C:\temp\$servername.txt -Append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $desc1 | Out-file  C:\temp\$servername.txt -Append


}

function NLBsettings
{
    Write-host "Processing..NLB settings settings."-ForegroundColor green
    #NLB settings
    $nlb = Get-WindowsFeature nlb | where { $_.Installed -eq 'true' }
    
    if ($nlb)
    {
        Import-Module  NetworkLoadBalancingClusters
        # 'Installed'
        $nlbcluster = Get-NlbCluster  -HostName $env:COMPUTERNAME -WarningAction SilentlyContinue -erroraction 'silentlycontinue'   | Out-String 
    
        if ($nlbcluster)
        {
                
            $clst = [char]0x2551 + "    Cluster Info                                                            " + [char]0x2551
            $clst1 = Get-NlbCluster  | Out-String
            $clstnode = " Cluster nodes"
            $clstnode1 = Get-NlbClusterNode  | Out-String
            $clstnode1 = Get-NlbClusterNode  | Out-String
            $affinity = (get-NLBClusterPortRule).Affinity
            $clstnode4 = "Cluster Affinity mode is set to: $affinity"  

    
            $blank | Out-file  C:\temp\$servername.txt -append
            $linetop | Out-file  C:\temp\$servername.txt -append
            $clst | Out-file  C:\temp\$servername.txt -Append
            $linebottom | Out-file  C:\temp\$servername.txt -append
            $clst1 | Out-file  C:\temp\$servername.txt -Append
            $clstnode3 | Out-file  C:\temp\$servername.txt -Append
            $clstnode | Out-file  C:\temp\$servername.txt -Append
            $clstnode1 | Out-file  C:\temp\$servername.txt -Append
            $clstnode3 | Out-file  C:\temp\$servername.txt -Append
            $clstnode4 | Out-file  C:\temp\$servername.txt -Append

        }
        else
        {
            $clst = [char]0x2551 + "    Cluster Info                                                            " + [char]0x2551
            $clst1 = 'Clustering Role Installed but Not in a Cluster '

            $blank | Out-file  C:\temp\$servername.txt -append
            $linetop | Out-file  C:\temp\$servername.txt -append
            $clst | Out-file  C:\temp\$servername.txt -Append
            $linebottom | Out-file  C:\temp\$servername.txt -append
            $clst1 | Out-file  C:\temp\$servername.txt -Append

        }

    }
    else {}

        
}

function filepermission
{
    Write-host "Processing..File system Permissions settings."-ForegroundColor green
    #file system Permissions
    $FilePerd = ""
    $FilePer = [char]0x2551 + "    Check File Permissions                                                  " + [char]0x2551
    $1 = get-psdrive -PSProvider FileSystem | where { $_.name -le "w" } | where { $_.name -ne "c" -and $_.name -ne "a" } 
        
    foreach ($root in $1)
    {
        $r = $root.root | Out-String
        $FilePerd = $FilePerd + (get-acl $root.root | fl | out-string)
    }
    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $fileper | Out-file  C:\temp\$servername.txt -Append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $fileperd | Out-file  C:\temp\$servername.txt -Append
}

function iisfolder
{
    Write-host "Processing..IIS log folder settings."-ForegroundColor green
      
    #if IIS not installed then hide this from report 
    if (!(Get-Command get-webconfiguration -ErrorAction SilentlyContinue))
    {
        write-host ""
    }
    else 
    {
        #test IIS log folder
        $iisfldr = [char]0x2551 + "    Test folder exists for the IIS Logs                                     " + [char]0x2551
        IF (!(TEST-PATH E:\inetpub\logs\LogFiles))
        { 
            $iisfldr1 = "E:\inetpub\logs\LogFiles DOES NOT exist"
        }
        else
        {
            $iisfldr1 = "E:\inetpub\logs\LogFiles DOES exist"
        }

        $blank | Out-file  C:\temp\$servername.txt -append
        $linetop | Out-file  C:\temp\$servername.txt -append
        $iisfldr | Out-file  C:\temp\$servername.txt -Append
        $linebottom | Out-file  C:\temp\$servername.txt -append
        $iisfldr1 | Out-file  C:\temp\$servername.txt -Append
    }
}

function iisdefaultlocal
{
    if (get-command get-webconfiguration -ErrorAction SilentlyContinue)
    {
        Write-host "Processing..IIS log folder Location settings."-ForegroundColor green
        #set iis Default Log location
        $iisdef = [char]0x2551 + "    IIS Default log location                                                " + [char]0x2551
        $iisdef1 = get-webconfiguration /System.Applicationhost/Sites/SiteDefaults/logfile | select dir* | Out-String
    }

}

Function adminpassword
{
    Write-host "Processing..Local Admin Password settings."-ForegroundColor green
    #Admin account password set never to expire
        
    $adminpwd = [char]0x2551 + "    Local Administrator account password set never to expire                " + [char]0x2551
    $ADS_UF_PASSWD_CANT_CHANGE = 64        # 0x40
    $ADS_UF_DONT_EXPIRE_PASSWD = 65536     # 0x10000

    $computer = $null
    $users = $null
    $computer = [ADSI]"WinNT://$env:computerName,computer"
    $Users = $computer.psbase.Children | Where-Object { $_.psbase.schemaclassname -eq 'user' }
    
    foreach ($user in $Users.psbase.syncroot)
    {
        try
        {
            If ( $user.name -eq "Administrator")
            {
                $user.userflags = $user.userflags[0] -bor $ADS_UF_DONT_EXPIRE_PASSWD
                #  $user.SetInfo()
                $name = $user.FullName | out-string
                # $name
                $adminpwd1 = "True   Password set to never Expire" 
            }
            else
            {
                $adminpwd1 = "False  Should be set to never Expire" 
            }
        }
        catch {}
    }
    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $adminpwd | Out-file  C:\temp\$servername.txt -Append
    $lineBottom | Out-file  C:\temp\$servername.txt -append
    $adminpwd1 | Out-file  C:\temp\$servername.txt -Append  
}

function Windozupdates
{
    Write-host "Processing..Windozs Update settings."-ForegroundColor green
    #Number of windows updates
             
    $winhotfix = [char]0x2551 + "    Number of windows updates should be Approx 15                           " + [char]0x2551
    $winhotfix1 = (Get-HotFix).count# | measure | fl Count | Out-String

    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $winhotfix | Out-file  C:\temp\$servername.txt -Append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    Write-output "$winhotfix1 updates Installed" | Out-file  C:\temp\$servername.txt -Append  

}

FUNCTION DRIVES
{
    Write-host "Processing..Checking size of Disks settings."-ForegroundColor green
          
    $drvs = [char]0x2551 + "    Checking size of Disks                                                  " + [char]0x2551 
    $drvs1 = get-wmiobject win32_volume | where { $_.driveletter -ne 'X:' -and $_.label -ne 'System Reserved' } | sort driveletter | Ft -AutoSize  Driveletter, label, 
    @{name = "  Disk Size(GB) "; Expression = { "{0,8:N0}" -f ($_.Capacity / 1gb) } } ,
    @{name = "  Free Disk Size(GB) "; Expression = { "{0,8:N0}" -f ($_.FreeSpace / 1gb) } } ,
    @{name = "  Block Size (KB) "; Expression = { "{0,8:N0}" -f ($_.Blocksize / 1kb) } } 

    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $drvs | Out-file  C:\temp\$servername.txt -Append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $drvs1 | Out-file  C:\temp\$servername.txt -Append 
}

function functionsinstalled
{
    Write-host "Processing..Installed Functions and Roles settings."-ForegroundColor green
         
    $funt = [char]0x2551 + "    Checking to see these roles and functions exist.                        " + [char]0x2551  
    
    $FUNTA = Get-WmiObject -query 'select * from win32_optionalfeature where installstate=1' | foreach { $_.Name }
    <#$FUNTA = Get-WindowsFeature | where-object { $_.installed -eq $true } | where name -ne "FileAndStorage-Services" `
    | where name -ne "File-Services" `
    | where name -ne "FS-FileServer"`
    | where name -ne "storage-Services" `
    | where name -ne "RDC"  `
    | where name -ne "RSAT"  `
    | where name -ne "RSAT-Feature-Tools" `
    | where name -ne "RSAT-Role-Tools" `
    | where name -ne "RSAT-AD-Tools" `
    | where name -ne "RSAT-AD-PowerShell" `
    | where name -ne "FS-SMB1" `
    | where name -ne "User-Interfaces-Infra" `
    | where name -ne "Server-Gui-Mgmt-Infra" `
    | where name -ne "Server-Gui-Shell" `
    | where name -ne "PowerShellRoot" `
    | where name -ne "PowerShell" `
    | where name -ne "PowerShell-ISE" `
    | where name -ne "WoW64-Support" 
#>
    $funt1num = ($funta  | measure ).count
    $funt1 = $funta | out-string
    $funt1 = $funt1 + $funt1num

    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $funt | Out-file  C:\temp\$servername.txt -Append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $funt1 | Out-file  C:\temp\$servername.txt -Append 

}

function bginfo
{
    Write-host "Processing..Checking BG Info settings."-ForegroundColor green
            
    $bginf = [char]0x2551 + "    Checking BG Info build date key set                                     " + [char]0x2551  
    $bgin = get-ItemProperty  -path HKLM:\SOFTWARE\Wow6432Node\QPS\SysInfo -name Build_Date
    $bginf1 = $bgin.Build_date

    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $bginf | Out-file  C:\temp\$servername.txt -Append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $bginf1 | Out-file  C:\temp\$servername.txt -Append 

}

function security
{
    Write-host "Processing..Checking Security Template settings."-ForegroundColor green

    $sec = [char]0x2551 + "    Checking to see if the 'Windows Server 2012 R2' Security folder exists. " + [char]0x2551 
    
    IF (!(TEST-PATH "C:\Windows\security\Windows Server 2012 R2"))
    {
        $sec1 = "Security Template Did not apply"
    }
        
    else
    {

        if (!(TEST-PATH C:\Windows\security\logs\ISSDefaultTemplate.log))
        {
            $sec1 = "Security Template Did not apply"
        }
        else
        {
            $s1 = (get-content C:\Windows\security\logs\ISSDefaultTemplate.log | select -last 3 -skip 1 ) 
            $s1 = $s1 + "" + (get-content C:\Windows\security\logs\ISSDefaultTemplate.log | select -last 1) 
            $sec1 = $s1
        }
    }
    #$sec1
 
    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $sec | Out-file  C:\temp\$servername.txt -Append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $sec1 | Out-file  C:\temp\$servername.txt -Append 
}

function openview
{
    Write-host "Processing..Checking hp oPENVIEW settings."-ForegroundColor green
                   
    $opview = [char]0x2551 + "    Checking to see if HPPA service exist.                                  " + [char]0x2551 
    $arrService = Get-Service -Name HPOvTrcSvc
    if ($arrService.Status -eq "Running")
    {
        $opview1 = "HP Software Shared Trace Service Exists and is started"
    }
    else
    { 
        $opview1 = "HP Software Shared Trace Service  ****Missing****"
    }

    $arrService = Get-Service -Name OvCtrl 
    if ($arrService.Status -eq "Running")
    {
        $opview2 = "HP OpenView Ctrl Service Exists and is started"
    }
    else
    { 
        $opview2 = "HP OpenView Ctrl  ****Service Missing****"
    }

    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $opview | Out-file  C:\temp\$servername.txt -Append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $opview1 | Out-file  C:\temp\$servername.txt -Append   
    $opview2 | Out-file  C:\temp\$servername.txt -Append     

}

function softwwareinstalled
{
    Write-host "Processing..Checking sOFTWARE Instalations settings."-ForegroundColor green
    ######   works out if Oracle Client is installed"
    new-item c:\temp\apps.txt -force -type file | out-null
    $smallline = repeat-string $ln20 12
    if (!(test-path -path HKLM:\SOFTWARE\ORACLE\KEY_OraClient11g_home1 -PathType Container ))
    {}
    else 
    {
        $oracleversion64 = "x64 Oracle Client Installed"
    }
    

    if (!(test-path -path HKLM:\SOFTWARE\Wow6432Node\ORACLE\KEY_OraClient11g_home1_32bit -PathType Container ))
    {}
    else 
    {
        $oracleversion32 = "x32 Oracle Client Installed"
     
    }


    Write $oracleversion

            
    $appsinst = [char]0x2551 + "    List of APPS Installed.                                                 " + [char]0x2551 
    $apps1 = "`n64 BIT Apps`n"
    $apps1 = $apps1 += $smallline + "`n"

    $apps1 = $apps1 += $oracleversion64
    $apps1 = $apps1 += Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | sort displayname | Select-Object DisplayName, Publisher | ft  | Out-String
    $apps1 = $apps1 + "`n32 BIT Apps`n"
    $apps1 = $apps1 += $smallline + "`n"
    $apps1 = $apps1 += $oracleversion32
    $apps1 = $apps1 += (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, Publisher | ft   | Out-String) 
    $apps1 >> c:\temp\apps.txt 
    


    $appsinst1 = gc c:\temp\apps.txt | where { $_ -notlike " *" }
    $appsinst1 = $appsinst1 + (gc c:\temp\apps.txt | where { $_ -notlike " *" } | measure).count


    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $appsinst | Out-file  C:\temp\$servername.txt -Append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $appsinst1 | Out-file  C:\temp\$servername.txt -Append   
    $appsinst2 | Out-file  C:\temp\$servername.txt -Append     

}

function Eventlogerrors
{
    Write-host "Processing..Checking Eventviewer Errors settings."-ForegroundColor green
            
    $evt = [char]0x2551 + "    Windows Event Log Errors                                                " + [char]0x2551 
    <#$evt0 = Get-WinEvent -ListLog 'Application', 'System', 'Security' -ErrorAction SilentlyContinue | Where RecordCount -gt 0 |`
        ForEach-Object -Process { Get-WinEvent -FilterHashtable @{LogName = $_.LogName; StartTime = (Get-Date).AddDays(-1); Level = 1, 2, 3 } -ErrorAction SilentlyContinue }
        
    $evt1 = Get-WinEvent -ListLog 'Application', 'System', 'Security' -ErrorAction SilentlyContinue | Where RecordCount -gt 0 |`
        ForEach-Object -Process { Get-WinEvent -FilterHashtable @{LogName = $_.LogName; StartTime = (Get-Date).AddDays(-1); Level = 1, 2, 3 } -ErrorAction SilentlyContinue } | `
        Select-Object LogName, LevelDisplayName, Id, Level, TimeCreated, Message | sort LogName | ft -auto | Out-String
#>
    $evt0a = Get-WinEvent -ListLog 'Application', 'System', 'Security' -ErrorAction SilentlyContinue | Where { $_.RecordCount -gt 0 } 
    $evt0b = $evt0a | select -ExpandProperty logname
    $evt1 = @()
    $evt1 = ForEach ($EVENTLOG in $evt0b)
    {
        Write-Output "Eventlog Name:  $eventlog "
        Get-WinEvent -FilterHashtable @{LogName = $eventlog; StartTime = (Get-Date).AddDays(-1); Level = 1, 2, 3 } -ErrorAction SilentlyContinue
   
    }

    if ($evt1.Count -gt 0)
    {
            
    }
        
    else
    {
        $evt1 = "No Errors Detected  "
    }

    

    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $evt | Out-file  C:\temp\$servername.txt -Append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $evt1 | Out-file  C:\temp\$servername.txt -Append    


}

function Sharesinfo
{
    Write-host "Processing..Checking File system Shares settings."-ForegroundColor green
    
    $smbshr = [char]0x2551 + "    File System Shares settings                                             " + [char]0x2551 
             
    if (!($smbshares = Get-WmiObject Win32_share ))
    #Get-SmbShare | where { $_.Special -notlike "*true*" }))
    {
        $sharenotexist = "No Shares Defined"  
           
        $smbshr = [char]0x2551 + "    File System Shares settings                                             " + [char]0x2551 
        $blank | Out-file  C:\temp\$servername.txt -Append
        $linetop | Out-file  C:\temp\$servername.txt -append
        $smbshr | Out-file  C:\temp\$servername.txt -Append
        $linebottom | Out-file  C:\temp\$servername.txt -append
        $sharenotexist | Out-file  C:\temp\$servername.txt -append
    }
    else
    {
        $smbshraa = $smbshares
        <#
        foreach ($smb in $smbshares)
        {
            $aa = ($smb.name)
            $smbshraa = $smbshraa + "Share Name: $aa`r`n"
            $smbshraa = $smbshraa + "Share Permissions:`r`n"
            $smbshraa = $smbshraa + "------------------------------------------------`r`n"
            $aaa = Get-SmbShareAccess ($smb).name  | ft Name, Accountname, AccessControlType, AccessRight -AutoSize | out-string 
            $smbshraa = $smbshraa + $aaa
            $smbshraa = $smbshraa + $acclper
            $bbb = (get-acl ($smb).Path).Access | select FileSystemRights , AccessControlType , IdentityReference | ft -auto | out-string
            $smbshraa = $smbshraa + $bbb
            $smbshraa = $smbshraa + $aaaa
        }
 #>

        $blank | Out-file  C:\temp\$servername.txt -Append
        $linetop | Out-file  C:\temp\$servername.txt -append
        $smbshr | Out-file  C:\temp\$servername.txt -Append
        $linebottom | Out-file  C:\temp\$servername.txt -append
        $smbshraa | Out-file  C:\temp\$servername.txt -Append   
    }

}

function TrustedForDelegation
{
    Write-host "Processing..Checking Trusted for Delegation settings."-ForegroundColor green
              
    $svrdel = [char]0x2551 + "    Checking to see if the Server Account is set for Trusted for Delegation " + [char]0x2551 
    install-windowsfeature RSAT-AD-PowerShell 2>&1 | Out-Null
        
    if (Get-ADComputer $env:COMPUTERNAMe -Properties *  | where { $_.TrustedForDelegation -like "*true*" }) { $svrdel1 = "Computer Account TrustedForDelegation equals true `r`r`n  Note: this should be True  if The Computer Account is Trusted For Delegation" }
        
    else
    {
        $svrdel1 = "Computer Account TrustedForDelegation equals False `r`n This Computer Account is Not Trusted for Delegation   "
    }
 
    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $svrdel | Out-file  C:\temp\$servername.txt -Append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $svrdel1 | Out-file  C:\temp\$servername.txt -Append   
   
}

function tasksched
{

    Write-host "Prosessing..Task Scheduler" -ForegroundColor green 
    $tasksched1 = [char]0x2551 + "    Checking Scheduled Tasks                                                " + [char]0x2551 
    $SchedTasks = ""   
               
    $SchedTasks = @();
    ForEach ($Task in (Get-ChildItem -Path 'C:\Windows\System32\Tasks'))
    {
        $TaskPrincipal = ([XML](Get-Content -Path $Task.FullName)).Task.Principals.Principal;
        $TaskRegistrationInfo = ([XML](Get-Content -Path $Task.FullName)).Task.RegistrationInfo;
        $SchedTasks += [PSCustomObject]@{'TaskName' = "$($Task.Name)"; 'Author' = $TaskRegistrationInfo.Author; 'User' = $TaskPrincipal.UserId; 'RunLevel' = $TaskPrincipal.RunLevel; }
    }
    #return the list of tasks back to outside the script block
    
    $SchedTasks = @()
    $SchedTasks = $SchedTasks | Where  User -notlike "" | where { $_.User -ne "system" } | where { $_.User -notlike "s-1-5-21*" } | where { $_.Taskname -notlike "Optimize Start Menu Cache Files-S-1-5-21-*-500" } | Out-String
         
    IF (!($SchedTasks))
    {
        $tasksched2 = "There are no Scheduled Tasks that are running as a prohibited User."
    }
    else
    {
        $tasksched2 = "There are Scheduled Tasks that are running as a prohibited User."
    }

    $tasksched3 = @()         
    $tasksched3 = $tasksched2 + $SchedTasks

    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $tasksched1 | Out-file  C:\temp\$servername.txt -append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $tasksched3 | Out-file  C:\temp\$servername.txt -append

}

function Eventlogsettingscheck
{
    Write-host "Checking EventLog Settings" -ForegroundColor green
                 
    $evlsc = [char]0x2551 + "    Event Log Config Check                                                  " + [char]0x2551 
    $Result = Get-Eventlog -list | `
        Where-Object { ($_.LogDisplayName -eq 'Application') -or `
        ($_.LogDisplayName -eq 'Security') -or `
        ($_.LogDisplayName -eq 'System') }# | `
    $Eventlog = $Result | Select-Object LogDisplayName, MaximumKilobytes, OverflowAction | ft -AutoSize | Out-String
    $evlsc2 = $Eventlog    
    if ($Result | Where-Object { $_.MaximumKilobytes -lt 50Mb / 1Kb })
    {
        $evlsc1 = "Eventlog Size setting issue.`n`n  Please check settings."
    }
        
    else
    { 
        $evlsc1 = " "
    }

    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $evlsc | Out-file  C:\temp\$servername.txt -append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $evlsc1 | Out-file  C:\temp\$servername.txt -append
    $evlsc2 | Out-file  C:\temp\$servername.txt -append


}
 
function CheckNTP
{
    Write-host "Processing..Time Sync" -ForegroundColor green

    $Result = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters" -Name NtpServer, Type | Select NtpServer, Type
    $ntpserver = $Result.NtpServer
    $Type = $Result.Type
    $pdc = (Get-ADDomain).pdcemulator
    $dt = gwmi win32_operatingsystem -computer $pdc
    $dt_str = $dt.converttodatetime($dt.localdatetime)
    $localtime = (get-date -Format "dd-MM-yyyy HH:mm:ss")
    $timezone = (Get-WMIObject -class Win32_TimeZone -ComputerName $env:computername).caption
                 
    $ntptitle = [char]0x2551 + "    Checking System Time Synced (NTP).                                      " + [char]0x2551 
    $ntp = " NTP Server              Type      PDC                           Local Time              Time Zone"
    $ntp1 = "-------------------------------------------------------------------------------------------------------"
    $ntp2 = " $ntpserver    $Type    $PDC   $localtime   $timezone "
    $timedif = New-TimeSpan -End $localtime -Start $dt_str
    if ($timedif -gt "0")
    {
        $ntp3 = "Time is not sync'ed "
        $ntp4 = "    Local time is: $localtime     Remote time is: $dt_str"
    }
    else
    {
        $ntp3 = " Time is Synced"
    }

    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $ntptitle | Out-file  C:\temp\$servername.txt -append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $ntp | Out-file  C:\temp\$servername.txt -Append
    $ntp1 | Out-file  C:\temp\$servername.txt -Append   
    $ntp2 | Out-file  C:\temp\$servername.txt -Append     
    $ntp3 | Out-file  C:\temp\$servername.txt -Append     
    $ntp4 | Out-file  C:\temp\$servername.txt -Append     

}

function cleareventlogs
{
    write-host  "`n Clearing all Eventlogs." -ForegroundColor green
    Get-WinEvent -ListLog * -Force | % { Wevtutil.exe cl $_.logname }
    New-EventLog –LogName Application –Source “INI Build”  2>&1 | Out-Null
    Write-EventLog –LogName Application –Source “INI Build” –EntryType Information –EventID 1  –Message “Log Event log cleaned by  :$env:username”

}

Function VMwarecheck
{
    Write-host "Processing..VMWare Tools" -ForegroundColor green
              
    $VMTools = [char]0x2551 + "    Checking to see if VMware Tools are running                             " + [char]0x2551

    if ($cpu1 -like '*VMware*')
    {
        if (!(Get-Process | where { $_.name -like 'vmtool*' }))
        {
            $VMTools1 = "VMware Tools are NOT running"
        }
        else 
        {
            $VMTools1 = "VMware Tools are running"
        }
    }
    else
    {
        $VMTools1 = "This is a Physical Server."
        write-host " This is a Physical Server."  
    }

    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $VMTools | Out-file  C:\temp\$servername.txt -Append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $VMTools1 | Out-file  C:\temp\$servername.txt -Append


}

function Services-check
{
    Write-host "Processing..Services are running as a User" -ForegroundColor green
                      
    $servicesrunning = [char]0x2551 + "    Checking to See If any Services are running as a User                   " + [char]0x2551
    $string = Get-WmiObject win32_service | where { $_.Startname -notlike "NT AUTHORITY\*" -and $_.Startname -ne "LocalSystem" -and $_.Startname -notlike "NT Service\*" } | ft name, Displayname, startname  -AutoSize
    if ($string) { $servicesrunning1 = $string } 
    else 
    { $servicesrunning1 = 'There are No Services running as a user.' }
 
    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $servicesrunning | Out-file  C:\temp\$servername.txt -Append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $servicesrunning1 | Out-file  C:\temp\$servername.txt -Append  
    $blank | Out-file  C:\temp\$servername.txt -append
                
}

function iisdefaultlocal
{
    # Moved it to below. Write-host "Processing..IIS Default Log Location." -ForegroundColor green
    
    if (!(Get-Command get-webconfiguration -ErrorAction SilentlyContinue))
    {
        # write-host ""
    }
    else 
    {
        # Showing it is Processing IIS
        Write-host "Processing..IIS Default Log Location." -ForegroundColor green
        #set iis Default Log location
                  
        $iisdef = [char]0x2551 + "    IIS Default log location                                                " + [char]0x2551
        $iisdef1 = get-webconfiguration /System.Applicationhost/Sites/SiteDefaults/logfile | select dir* | Out-String
        $blank | Out-file  C:\temp\$servername.txt -append
        $linetop | Out-file  C:\temp\$servername.txt -append
        $iisdef | Out-file  C:\temp\$servername.txt -Append
        $linebottom | Out-file  C:\temp\$servername.txt -append
        $iisdef1 | Out-file  C:\temp\$servername.txt -Append   
    }
}

Function QCAD-VAS-folder-permissions
{
    Write-host "Checking QCAD-VAS-folder-permissions" -ForegroundColor green
    $FilePerdvas = ""
                   
    $FilePervas = [char]0x2551 + "    Check File Permissions for QCAD VAS servers.                            " + [char]0x2551
    $FilePerdvas = $FilePerdvas + "Checking permissions for D:\Oracle\Admin\prcaddg\adump"
    $FilePerdvas = $FilePerdvas + (get-acl D:\Oracle\Admin\prcaddg\adump | fl | out-string)

    $FilePerdvas = $FilePerdvas + "Checking permissions for D:\Oracle\Admin\prcad\adump"
    $FilePerdvas = $FilePerdvas + (get-acl D:\Oracle\Admin\prcad\adump | fl | out-string)

}

Function Windowsver
{
    Write-host "Processing.. Windows Version" -ForegroundColor green
                      
    $windowsversion = [char]0x2551 + "    Checking Windows Version                                                " + [char]0x2551
    $windowsversion1 = (Get-WmiObject -class Win32_OperatingSystem).Caption


    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $windowsversion | Out-file  C:\temp\$servername.txt -append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $windowsversion1 | Out-file  C:\temp\$servername.txt -append


}

Function Cleanup-install-folders
{
    param(
        [string]$cleanpath
    )

    if ( $(Try { Test-Path $cleanpath } Catch { $false }) )
    {
        Write-host "Removing $cleanpath" -ForegroundColor Yellow
        rmdir $cleanpath -force -Recurse
    }

    if ( $(Try { Test-Path c:\temp } Catch { $false }) )
    {
        ""
    }
}



function Uptime
{

    Write-host "Last time System was rebooted Check" -ForegroundColor Green
 
    $Uptimetitle = [char]0x2551 + "    Last time System was rebooted Check                                     " + [char]0x2551
    $wmi = Get-WmiObject -Class Win32_OperatingSystem
    $lastreboot = ($wmi.ConvertToDateTime($wmi.LastBootUpTime)) | Out-String
    $uptimecheck = "This Server was last rebooted $lastreboot  "  
    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $Uptimetitle | Out-file  C:\temp\$servername.txt -append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $uptimecheck | Out-file  C:\temp\$servername.txt -append
     
    #Write to Screen
    write-host "`n`n $uptimecheck"  -ForegroundColor yellow

}


function cleanupinstallscripts
{
    Write-host "Cleaning up old install folders left over from the Scripted installations if they exist." -ForegroundColor Green
    Cleanup-install-folders "C:\temp\hppa"
    Cleanup-install-folders "C:\temp\sccm"
    Cleanup-install-folders "C:\temp\ODAC110203_x64"
    Cleanup-install-folders "c:\temp\OracleClient_11.2.0.4_x86"
    Cleanup-install-folders "C:\temp\SQLEXPRWT_x64_ENU"
    Cleanup-install-folders "C:\temp\networker"

}


function repeat-string([string]$str, [int]$repeat) { $str * $repeat }

Function get-OU
{
    Write-host "Processing..OU Settings." -ForegroundColor Green
 
    #getting the Info for the server. 
    $oulocal = Get-ADComputer ($env:COMPUTERNAME) -Properties *                   
                                                                                                                                            
    $OUTitle = [char]0x2551 + "    Organisational Unit for this system.                                    " + [char]0x2551
    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $oUTitle | Out-file  C:\temp\$servername.txt -append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $oulocal.CanonicalName | Out-file  C:\temp\$servername.txt -append
}

Function Send-report
{
    $findusername = $env:username
    $pattern = ’[^1234567890]’
    $emailusername = ($findusername –replace $pattern, ’’) -as [string]
    $emailusername
    $emailcreds = Get-Credential   -Message "Enter your PRDS standard user code. Do not enter Domain.  This is required to send email only."
    $global:emailcreds = $emailcreds
    $emailaddress1 = get-aduser $emailcreds.username -Properties *
    $Global:emailaddress = $emailaddress1.mail
    Send-MailMessage -From $emailaddress -Subject "QA Report for: $env:computername" -To $emailaddress -Body "QA Report for `n`n   Computername  $env:computername" -Attachments C:\temp\$servername.txt -Credential $emailcreds -Port 25 -SmtpServer smtp.police.qld.gov.au


}

Function Logoff
{
    Shutdown /l /f 
}


      
<#
Start of the MAin part to the Script
#>
#sets up what is needed to create the the output enviroment for QA report.

cls
$servername = $env:COMPUTERNAME
$Scriptver = "QA Script Ver $ver"
$name = "                                                    $servername"
$blank = " " 
import-module ServerManager  # crap 2008 does not have this modle loaded by default and not auto loading.
Write-host "Installing Powershell AD Tools.  This will be cleaned up at end of process." -ForegroundColor Cyan
install-windowsfeature RSAT-AD-PowerShell 2>&1 | Out-Null
add-windowsfeature RSAT-AD-PowerShell 2>&1 | Out-Null
Import-Module ActiveDirectory
#Install-WindowsFeature RSAT-AD-Tools

#check to see if c:\temp Exists for Report.
if ( !$( Test-Path c:\temp )) { mkdir 'c:\temp' }

#Setting up the lines for around titles.
$ln1 = [char]0x2554 #Left top
$ln2 = [char]0x2557 #Right Top
$ln20 = [char]0x2550 #=
$ln5 = [char]0x2588 
$ln30 = repeat-string $ln20 76
$ln40 = repeat-string $ln5 110
$linetop = $ln1 + $ln30 + $ln2
$ln3 = [char]0x255A #Left bottom
$ln4 = [char]0x255D #Right Bottom

$linebottom = $ln3 + $ln30 + $ln4
$line = repeat-string $ln20 110


#setting up the Title box.
$Title1 = $ln1 + $line + $ln2
$Title2 = $ln3 + $line + $ln4

$Title10 = [char]0x2551 + "                                               " + $servername + "                                                  " + [char]0x2551

$Title1  | Out-file  C:\temp\$servername.txt 
$Title10 | Out-file  C:\temp\$servername.txt -append 
$Title2 | Out-file  C:\temp\$servername.txt -append
$blank | Out-file  C:\temp\$servername.txt -append
$Scriptver | Out-file  C:\temp\$servername.txt -append

#cls
cleanupinstallscripts
Scriptver
cpu
Windowsver
VMwarecheck
memory
pagefile
networksettings
NLBsettings
Domain
get-OU
descript
iisfolder
iisdefaultlocal
Drives
filepermission     #No longer changed
bginfo
security           #do not use the /Security Profile now.  2019 is secure by detault.
adminpassword
Services-check
Windozupdates
functionsinstalled
tasksched
CheckNTP
#openview           # Not used now.
softwwareinstalled
eventlogerrors
Eventlogsettingscheck
sharesinfo
TrustedForDelegation
Uptime

#Display the Report in Notepad.
notepad C:\temp\$servername.txt
#Showing on screen where file Report is located.
Write-host "`n`n$env:computername.txt file is located in c:\temp" -BackgroundColor Yellow -ForegroundColor Red 


#Comment this out as well as the line above if you do not want to email the report to you.
Write-host ="`n`n To send this report to yourself enter 'send-report' " -BackgroundColor green -ForegroundColor Red 

#cleaning up the Powershell RSAT module.
Write-host "Removing Powershell AD Tools.  As installed at start of Script." -ForegroundColor Cyan
uninstall-windowsfeature RSAT-AD-PowerShell 2>&1 | Out-Null

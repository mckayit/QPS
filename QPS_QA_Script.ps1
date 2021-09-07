<# 
.SYNOPSIS
  ********************************************************************************  
  *                                                                              *  
  *  This script gets the info required to complete the Server QA report         *
  *                                                                              *  
  ********************************************************************************    

 #... <Short Description>

    Short description
.DESCRIPTION
      This script gets the info required to complete the Server QA report    
      It uses a range of WMI as well as Powershell CMDLets and outputs all the results to c:\temp\<Servername>.txt
      It also has a switch that allows you to send the Report to your email address by using the Following Switch -SendEmail at the end of the QPS_QA_Script.ps1 file.
      Eg QPS-QA_Script -EendEmail
      You will be prompted to enter the Email Address you want to send the report to.    Or at the end of the check you can amso enter Send-Report
       and you will be prompted for your email address and a copy will be sent to you.
.PARAMETER SendEmail
    This Prarameter is the switch used to prompt for the Email Address and send the Final report to.
    

.PARAMETER DontDisplayReport
    This Switch disables the Displaying of the Final Report in Notepad.


.EXAMPLE
# This will not send email with copy of QA Report.
    C:\PS>QPS_QA_Script   
## This will send email with copy of QA Report.
    C:\PS>QPS_QA_Script -sendemail

## This will disable notpad displaying the QA Report.
    C:\PS>QPS_QA_Script -DontDisplayReport


## This will disable notpad displaying the QA Report and Email the Report.
    C:\PS>QPS_QA_Script -DontDisplayReport -sendemail

         
    

.NOTES
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
            1.51     09  July  2021   Lawrence       Added BGINNFO Reg Key
            1.52     12  July  2021   Lawrence       Cleaned up code abit.   and also using get-computerinfo to
            1.54     06  aug   2021   Lawrence       fixed up issue with Clustering as well as PS AD tools
            1.55     06  Aug   2021   Lawrence       Fixed up how CPU displayed   big Thanks to Rana..
            1.56     11  Aug   2021   Lawrence       Fixed up AD and Local Description Format issue
            1.57     11  Aug   2021   Lawrence       Added IPv6 Check to see if enabled. 
                                                     Added Check for BGInfo Reg Keys
                                                     Fixed Clustering Output. 
                                                     Added Check for Performance settings.
                                                     Shows if Nutanix tools installed for a Nutanis System as well as VMWare tools
                                                     Fixed Title (Server name) display issue.
                                                     Added Commandline switch to auto prompt for Creds and send email
            1.58     11  Aug   2021   Lawrence       Now checking Performance settings in the .Default Reg key where it is now set.
            1.59     16  Aug   2012   Lawrence       Added the DontDisplayReport Switch as per Gary Chow Request.
                                                     Removed some Char's that Ansiable can not deal with when it copies a file.
                                                     It was rewriting them as something else.   EG BUG in Ansiable (’’, ’ and  – )
                                            **Note** this may comeback as it depends on the Editor used.
            1.60     06 Sept   2021   Lawrence       Fixed Nutanix detection                                   
            1.61     06 Sept   2021   Lawrence       Fixed issue with Time Sync         
            1.62                             





#> 

[CmdletBinding()]
param
(
    [Switch]$Sendemail,
    [Switch]$DontDisplayReport
        
)
	

if ($Sendemail) # Sendemail switch used
{
    Add-Type -AssemblyName Microsoft.VisualBasic
    $emailaddress = [Microsoft.VisualBasic.Interaction]::InputBox('Enter Your Email Address to send Repot to.:', 'Enter Email Address')

}

$Global:ver = "1.61"     





#checks to see if running as Admin
clear
Write-host  "`n`n"
write-host  "  *********************************" -ForegroundColor Green
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
    <#
 #... gets the CPU info on current system
.SYNOPSIS
   gets the CPU info on current system
.DESCRIPTION
    gets the CPU info on current system
.PARAMETER one
    Specifies Pram details.
.PARAMETER two
    Specifies Pram details
.PARAMETER InputObject
    Specifies the object to be processed.  You can also pipe the objects to this command.
.EXAMPLE
    C:\PS>get-cpuinfo 
    Example of how to use this cmdlet
.INPUTS
    Inputs to this cmdlet (if any)
.OUTPUTS
    Output from this cmdlet (if any)
.NOTES
    
    Lawrence McKay
    Lawrence@mckayit.com
    McKayIT Solutions Pty 
     
    Date:    15 Feb 2021
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           9 July  2021         Lawrence       Initial Coding

#>
     
   
    
    begin 
    {
        Write-host "Processing..CPU settings."-ForegroundColor green
    }
    
    process 
    {
            
        try 
        {
            $global:CPUINfO_ = Get-ComputerInfo 

            $prop = [ordered]@{
                "Manufacturer"                  = $global:CPUINfO_.CsManufacturer
                "Machine Type"                  = $global:CPUINfO_.CsModel
                "Number of Logicial Processors" = $global:CPUINfO_.CsNumberOfLogicalProcessors   
                "Number of Sockets"             = $global:CPUINfO_.CsNumberOfProcessors
            }
            # Adding the Cockets to a Seperate Line
            $i = 0
            foreach ($cpu in $($global:CPUINfO_.CsProcessors.Name))
            {
                $prop.add("Socket_$($i)", "$cpu")
                $i++
            }

        }
        catch 
        {
            Write-Host 'ERROR : $(_.Exception.Message)' -ForegroundColor Magenta
        }
    }
    
    end
    {
        #output
        $cpu = [char]0x2551 + "    CPU info.                                                               " + [char]0x2551

        $blank | Out-file  C:\temp\$servername.txt -append
        $linetop | Out-file  C:\temp\$servername.txt -append
        $cpu  | Out-file  C:\temp\$servername.txt -append
        $linebottom | Out-file  C:\temp\$servername.txt -append
        [PScustomobject]$prop | Out-file  C:\temp\$servername.txt -Append
            
    }
    
}

function memory
{
    # Display memory 
    Write-host "Processing..Memory settings."-ForegroundColor green
    # get info from get-computerinfo
    $mem1 = ([math]::Round(($global:CPUINfO_.CsTotalPhysicalMemory / 1GB), 2))
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
    $PGSize3a = " "#      NOTE: It should be false"
    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $PGSize | Out-file  C:\temp\$servername.txt -Append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $PGSize1 | Out-file  C:\temp\$servername.txt -Append
    $PGSize2 | Out-file  C:\temp\$servername.txt -Append
    $PGSize3a | Out-file  C:\temp\$servername.txt -Append
}

function networksettingsipv6
{
    #network settings
    Write-host "Processing..Network settings IPv6."-ForegroundColor green
    $netwrkv6 = [char]0x2551 + "    IPv6 Network settings.                                                  " + [char]0x2551
    $netwrk1v6 = "IPv6 is Disabled. (Expected)"
    #if IPv6 exists 
    if (Get-NetAdapterBinding | where { $_.DisplayName -match "IPv6" -and $_.enabled -eq 'true' } )
    {
        $netwrk1v6 = 'IPv6 Enabled  ---> *** Please Disable ***'
    }
    
    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $netwrkv6 | Out-file  C:\temp\$servername.txt -Append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $netwrk1v6 | Out-file  C:\temp\$servername.txt -Append
}


function networksettings
{
    #network settings
    Write-host "Processing..Network settings."-ForegroundColor green
    $netwrk = [char]0x2551 + "    Network settings.                                                       " + [char]0x2551
    $netwrk1 = Get-NetIPConfiguration -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Out-String

    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $netwrk | Out-file  C:\temp\$servername.txt -Append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $netwrk1 | Out-file  C:\temp\$servername.txt -Append

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

}

function domain
{
    Write-host "Processing..Domain settings."-ForegroundColor green
    #Server Domain
        
    $dom = [char]0x2551 + "    User Domain          Server Domain                                      " + [char]0x2551
    # $DOM2 = Get-NetIPConfiguration -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
    $CSDomain = $global:CPUINfO_.CsDomain.tostring()
    $dom1 = "    " + $env:USERDNSDOMAIN + "            " + $global:CPUINfO_.CsDomain.tostring()

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
    #getting INFO and Outputting it.
    $LocalDesc = $($pcdesc.description.tostring())
    $ADDESC = $(get-adcomputer $servername -Properties  Description | select -ExpandProperty Description).tostring()
     
    $propDesc = [PSCustomObject]@{
        "Local Description" = $LocalDesc
        "AD Description"    = $ADDESC
    }


    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $desc | Out-file  C:\temp\$servername.txt -Append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $propDesc | fl | Out-file  C:\temp\$servername.txt -Append

}
function IsActivated
{
    $activ = Get-CIMInstance -query "select Name, Description, LicenseStatus from SoftwareLicensingProduct where LicenseStatus=1" | select -ExpandProperty LicenseStatus
    $ActivationStatus = "Not Activated"
    if ($activ -eq "1") { $ActivationStatus = "This Windows Operating System Version is Activated" }



    $OUTitle = [char]0x2551 + "    Server Activation Status                                                " + [char]0x2551
    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $oUTitle | Out-file  C:\temp\$servername.txt -append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $ActivationStatus | Out-file  C:\temp\$servername.txt -append

}

function NLBsettings
{
    Write-host "Processing..NLB settings settings."-ForegroundColor green
    #NLB settings
    $nlb = Get-WindowsFeature nlb | where InstallState -eq Installed 
    
    if ($nlb)
    {
        #'Installed'
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
            $CNotINSTALLED = "Network LoadBalanceing installed but not configured."   
            $blank | Out-file  C:\temp\$servername.txt -append
            $linetop | Out-file  C:\temp\$servername.txt -append
            $clst | Out-file  C:\temp\$servername.txt -Append
            $linebottom | Out-file  C:\temp\$servername.txt -append
            $CNotINSTALLED | Out-file  C:\temp\$servername.txt -append
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

Function adminpassword
{
    Write-host "Processing..Local Admin Password settings."-ForegroundColor green
    #Admin account password set never to expire
        
    $adminpwd = [char]0x2551 + "    Local Administrator account password set never to expire                " + [char]0x2551

    $ex = get-localuser -Name Administrator | select -ExpandProperty passwordexpires
    $adminpwd1 = "False   Password set to EXPIRE" 
    if (!($ex) )
    { 
        $adminpwd1 = "True   Password set to never Expire" 
    }
    #$adminpwd1


    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $adminpwd | Out-file  C:\temp\$servername.txt -Append
    $lineBottom | Out-file  C:\temp\$servername.txt -append
    $adminpwd1 | Out-file  C:\temp\$servername.txt -Append  
}


FUNCTION DRIVES
{
    Write-host "Processing..Checking size of Disks settings."-ForegroundColor green
          
    $drvs = [char]0x2551 + "    Checking size of Disks                                                  " + [char]0x2551 
    $drvs1 = Get-Volume | sort DriveLetter
    <# $drvs1 = get-wmiobject win32_volume | where { $_.driveletter -ne 'X:' -and $_.label -ne 'System Reserved' } | sort driveletter | Ft -AutoSize  Driveletter, label, 
    @{name = "  Disk Size(GB) "; Expression = { "{0,8:N0}" -f ($_.Capacity / 1gb) } } ,
    @{name = "  Free Disk Size(GB) "; Expression = { "{0,8:N0}" -f ($_.FreeSpace / 1gb) } } ,
    @{name = "  Block Size (KB) "; Expression = { "{0,8:N0}" -f ($_.Blocksize / 1kb) } } 
#>
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
    $FUNTA = Get-WindowsFeature | where-object { $_.installed -eq $true } | where name -ne "FileAndStorage-Services" `
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
            
    $bginf = [char]0x2551 + "    Checking BG Info REG Keys set                                           " + [char]0x2551  
    $bgin = get-ItemProperty  -path HKLM:\SOFTWARE\BGinfo
    $bginf1 = $bgin.Build_date

    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $bginf | Out-file  C:\temp\$servername.txt -Append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $bgin | Out-file  C:\temp\$servername.txt -Append 

}

function Windozupdates
{
    Write-host "Processing..Windozs Update settings."-ForegroundColor green
    #Number of windows updates
              
    $winhotfix = [char]0x2551 + "    Number of windows updates should be Approx 8 for 2019                   " + [char]0x2551
    $winhotfix1 = (Get-HotFix)# | measure | fl Count | Out-String

    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $winhotfix | Out-file  C:\temp\$servername.txt -Append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    Write-output "$(($winhotfix1).count) updates Installed"  | Out-file  C:\temp\$servername.txt -Append  
    $winhotfix1 | select "HotfixID", "InstalledON"  | Out-file  C:\temp\$servername.txt -Append 

}
function WindowsUpdatesSettings
{
     
    $NotificationLevels = @{ 0 = "0 - Not configured"; 1 = "1 - Disabled"; 2 = "2 - Notify before download"; 3 = "3 - Notify before installation"; 4 = "4 - Scheduled installation"; 5 = "5 - Users configure" }
    $ScheduledInstallationDays = @{ 0 = "0 - Every Day"; 1 = "1 - Every Sunday"; 2 = "2 - Every Monday"; 3 = "3 - Every Tuesday"; 4 = "4 - Every Wednesday"; 5 = "5 - Every Thursday"; 6 = "6 - Every Friday"; 7 = "7 - EverySaturday" }
    $RegistryKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]"LocalMachine", $servername) 
    $RegistrySubKey1 = $RegistryKey.OpenSubKey("SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\") 
    $RegistrySubKey2 = $RegistryKey.OpenSubKey("SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\")                
       
    #Titlebox -SubjectTitle "Windows Update"
    
    $Result = New-Object -TypeName PSObject
    Try
    {
        Foreach ($RegName in $RegistrySubKey1.GetValueNames())
        { 
            $Value = $RegistrySubKey1.GetValue($RegName) 
            $Result | Add-Member -MemberType NoteProperty -Name $RegName -Value $Value
        }
        Foreach ($RegName in $RegistrySubKey2.GetValueNames())
        { 
            $Value = $RegistrySubKey2.GetValue($RegName) 
            Switch ($RegName)
            {
                "AUOptions" { $Value = $NotificationLevels[$Value] }
                "ScheduledInstallDay" { $Value = $ScheduledInstallationDays[$Value] }
            }
            $Result | Add-Member -MemberType NoteProperty -Name $RegName -Value $Value
        }
    }
    Catch
    {
        Write-Error "Could not find registry subkey: HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate. Probably you are not using Group Policy for Windows Update settings." -ErrorAction Stop
    }


    $winhotfixtitle = [char]0x2551 + "    Windows Update Settings.                                                " + [char]0x2551 
    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $winhotfixtitle | Out-file  C:\temp\$servername.txt -Append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    
    
    $result | Out-file  C:\temp\$servername.txt -Append
                         
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

function Sharesinfo
{
    Write-host "Processing..Checking File system Shares settings."-ForegroundColor green
    
    $smbshr = [char]0x2551 + "    File System Shares settings                                             " + [char]0x2551 
             
    if (!($smbshares = Get-SmbShare | where { $_.Special -notlike "*true*" }))
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
    ForEach ($Task in (Get-ChildItem -Path 'C:\Windows\System32\Tasks' -File))
    {
        $TaskPrincipal = ([XML](Get-Content -Path $Task.FullName)).Task.Principals.Principal;
        $TaskRegistrationInfo = ([XML](Get-Content -Path $Task.FullName)).Task.RegistrationInfo;
        $SchedTasks += [PSCustomObject]@{'TaskName' = "$($Task.Name)"; 'Author' = $TaskRegistrationInfo.Author; 'User' = $TaskPrincipal.UserId; 'RunLevel' = $TaskPrincipal.RunLevel; }
    }
    #return the list of tasks back to outside the script block
    $SchedTasks = $SchedTasks | Where  User -notlike "" | where { $_.User -ne "system" } | where { $_.User -notlike "s-1-5-21*" } | where { $_.Taskname -notlike "Optimize Start Menu Cache Files-S-1-5-21-*-500" } | Out-String
         
    IF (!($SchedTasks))
    {
        $tasksched2 = "There are no Scheduled Tasks that are running as a prohibited User."
    }
    else
    {
        $tasksched2 = "There are Scheduled Tasks that are running as a prohibited User."
    }

               
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

function Eventlogerrors
{
    Write-host "Processing..Checking Eventviewer Errors settings."-ForegroundColor green
            
    $evt = [char]0x2551 + "    Windows Event Log Errors                                                " + [char]0x2551 
    $evt0 = Get-WinEvent -ListLog 'Application', 'System', 'Security' -ErrorAction SilentlyContinue | Where RecordCount -gt 0 |`
        ForEach-Object -Process { Get-WinEvent -FilterHashtable @{LogName = $_.LogName; StartTime = (Get-Date).AddDays(-1); Level = 1, 2, 3 } -ErrorAction SilentlyContinue }
        
    $evt1 = Get-WinEvent -ListLog 'Application', 'System', 'Security' -ErrorAction SilentlyContinue | Where RecordCount -gt 0 |`
        ForEach-Object -Process { Get-WinEvent -FilterHashtable @{LogName = $_.LogName; StartTime = (Get-Date).AddDays(-1); Level = 1, 2, 3 } -ErrorAction SilentlyContinue } | `
        Select-Object LogName, LevelDisplayName, Id, Level, TimeCreated, Message | sort LogName | ft -auto | Out-String

    if ($evt0.Count -gt 0)
    {
        $evt1 = "Eventlog Errors Found $evt1"
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

function cleareventlogs
{
    write-host  "`n Clearing all Eventlogs." -ForegroundColor green
    Get-WinEvent -ListLog * -Force | % { Wevtutil.exe cl $_.logname }
    New-EventLog -LogName Application -Source "INI Build"  2>&1 | Out-Null
    Write-EventLog -LogName Application -Source "INI Build" -EntryType Information -EventID 1  -Message "Log Event log cleaned by  :$env:username"

}

Function VMwarecheck
{
    Write-host "Processing..VMWare Tools" -ForegroundColor green
              
    $VMTools = [char]0x2551 + "    Checking to see if VMware / Nutanix Tools are running.                  " + [char]0x2551

    $VMTools1 = "This is a Physical Server."
    #write-host " This is a Physical Server." 

    if ($global:CPUINfO_.CsManufacturer -like '*VMware*')
    {
        if (!(Get-Process | where name -like "*vmtool*"))
        {
            $VMTools1 = "VMware Tools are NOT running"
        }
        else 
        {
            $VMTools1 = "VMware Tools are running"
        }
    }
    
    if ($global:CPUINfO_.CsManufacturer -like '*nutan*')
    {
        if (!(get-service | where name -like "Nutanix Guest*"))
        {
            $VMTools1 = "Nutanix Tools are NOT running"
        }
        else 
        {
            $VMTools1 = "Nutanix Tools are running"
        }
    
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
    $string = Get-WmiObject win32_service | where Startname -notlike "NT AUTHORITY\*"  | where Startname -ne  "LocalSystem" | where Startname -notlike  "NT Service\*" | ft name, Displayname, startname  -AutoSize
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

Function Windowsver
{
    Write-host "Processing..Windows Version" -ForegroundColor green
                      
    $windowsversion = [char]0x2551 + "    Checking Windows Version                                                " + [char]0x2551
    $windowsversion1 = $global:CPUINfO_.OsName


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
    <#  $findusername = $env:username
    $pattern = ’[^1234567890]’
    $emailusername = ($findusername -replace $pattern, ’’) -as [string]
    $emailusername
    $emailcreds = Get-Credential   -Message "Enter your PRDS standard user code. Do not enter Domain.  This is required to send email only."
    $global:emailcreds = $emailcreds
    $emailaddress1 = get-aduser $emailcreds.username -Properties *
    $Global:emailaddress = $emailaddress1.mail
    #>

    # prompt for Email Address
    Add-Type -AssemblyName Microsoft.VisualBasic
    $emailaddress = [Microsoft.VisualBasic.Interaction]::InputBox('Enter Your Email Address to send Repot to.:', 'Enter Email Address')

                
    Send-MailMessage -From "Server_QA-Reports@PRDS.QLDPOL" -Subject "QA Report for: $env:computername" -To $emailaddress -Body "QA Report for `n`n   Computername  $env:computername" -Attachments C:\temp\$servername.txt  -Port 25 -SmtpServer smtp.police.qld.gov.au


}

Function Logoff
{
    Shutdown /l /f 
}

Function get-Performance
{
    Write-host "Processing..getting Performance settings." -ForegroundColor Green
 
    #getting the Info for the server. 
    # user based  $perf = Get-ItemProperty -path HKU:\.default\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects 
    #Checks the .Default settings.
    $perf = Get-ItemProperty -path registry::HKEY_USERS\.default\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects
    $Perfsettings = 'Not set correctly"'             
    if ( $perf.VisualFXSetting -eq '2')
    {
        $Perfsettings = "Server set for Best Performance"
    }                                                             
    

    $OUTitle = [char]0x2551 + "    Server Performance Settings                                             " + [char]0x2551
    $blank | Out-file  C:\temp\$servername.txt -append
    $linetop | Out-file  C:\temp\$servername.txt -append
    $oUTitle | Out-file  C:\temp\$servername.txt -append
    $linebottom | Out-file  C:\temp\$servername.txt -append
    $Perfsettings | Out-file  C:\temp\$servername.txt -append
}

      
<#
Start of the MAin part to the Script
#>
#sets up what is needed to create the the output enviroment for QA report.


$servername = $env:COMPUTERNAME
$Scriptver = "QA Script Ver $ver"
$name = "                                                    $servername"
$blank = " " 
Write-host "Installing Powershell AD Tools.  This will be cleaned up at end of process." -ForegroundColor Cyan

#gets status of PS AD tools installed or not. 
$rsat_ad = Get-WindowsFeature RSAT-AD-PowerShell | select -ExpandProperty installstate
if ($rsat_ad -notmatch 'installed')
{
    install-windowsfeature RSAT-AD-PowerShell 2>&1 | Out-Null
    #Install-WindowsFeature RSAT-AD-Tools
}
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
$Title1 = $line 
$Title2 = $line 

$Title10 = "                                               " + $servername  

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
IsActivated
VMwarecheck
memory
pagefile
networksettings
networksettingsipv6
NLBsettings
Domain
get-OU
descript
BGINFO
get-Performance
iisfolder
iisdefaultlocal
Drives
#filepermission     #No longer changed
#security           #do not use the /Security Profile now.  2019 is secure by detault.
adminpassword
Services-check
Windozupdates
WindowsUpdatesSettings
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


if (!($DontDisplayReport)) #dont sent report switch no used
{
    #Display the Report in Notepad.
    notepad C:\temp\$servername.txt
}
#Display the Report in Notepad.
notepad C:\temp\$servername.txt
#Showing on screen where file Report is located.
Write-host "`n`n$env:computername.txt file is located in c:\temp" -BackgroundColor Yellow -ForegroundColor Red 


#Comment this out as well as the line above if you do not want to email the report to you.
Write-host ="`n`n To send this report to yourself enter 'send-report' " -BackgroundColor green -ForegroundColor Red 
if ($Sendemail) # Sendemail switch used
{
    Send-report 
}
#cleaning up the Powershell RSAT module.
Write-host "Removing Powershell AD Tools.  As installed at start of Script." -ForegroundColor Cyan
if ($rsat_ad -notmatch 'installed')
{
    uninstall-windowsfeature RSAT-AD-PowerShell 2>&1 | Out-Null
}

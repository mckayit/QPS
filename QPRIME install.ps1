<# 
  ********************************************************************************  
  *                                                                              *  
  *  This script Configures D and E drive, installs Apps and fixes up the Build  *
  *                          for Windows Server 2012R2                           *
  *                                                                              *  
  ********************************************************************************    

    Note:
            All that nneds to be Done to the Server before running this is Add the IP address's and Join the Domain.
            Reboot and then run in an evelated POWERSHELL
        
            THIS SCRIPT REQUIRES to be run in an evelated POWERSHELL (Prefered ISE) or 
            eg "powershell.exe <path to script>\Config-server-vx.x.ps1 -Verb runAs"

    Author:
            Lawrence McKay
            Lawrence@mckayit.com
            McKayIT Solutions Pty Ltd
    
    Date:   30 Aug 2015

   ******* Update Version number below when a change is done.*******

    
    History
            Version    Date         Name           Detail
            ------------------------------------------------------------------------------------
            0.0.1      28 Aug 2015  Lawrence       Initial Coding
            1.1        29 Aug 2015  Lawrence       Added Backup Request
            1.2        2 Sept 2015  Lawrence       Converted Drives creation to Pure Powershell
                       2 Sept 2015  Lawrence       Converted Password never expires to Pure Powershell
            1.3       14 September  Lawrence       Added full name to the Backup Request and changed format to RTF.
                      14 September  Lawrence       Added NLB Feature to function "qprime-NDS_requirements"
            1.4       
            1.5       17 September  Lawrence       Now setting DNS to be the 4 main servers (“10.47.1.201","164.112.162.142","10.46.253.23","164.112.162.141")  
            1.6       21 September  Lawrence       added IIS Redir Log files to e:\inetpub\logs\LogFiles
            1.7       23 September  Lawrence       Added new Sophos Installation commandline.
            1.8       23 September  Lawrence       Added password check against PRDS Domain   (Stop account lockout)
            1.9       23 September  Lawrence       Added option for Setting DVDS DNS entries. Need to comment / uncomment out.
            1.10      28 September  Lawrence       Added option to select backup day  
            1.11      29 September  Lawrence       Check CD Rom is on X: if not move to to x:  Added in the function (disk-d_and_e)
                                                   Added logic to Check that d Drive is the Small drive (disk-d_and_e).
            1.12      30 September  Lawrence       Cleaned up code    Made a Lock Point in time for this script..
            1.13       1 October    Lawrence       Added  Web-Request-Monitor and Web-http-logging to (AIP function)
            1.14       2 October    Lawrence       Converted the HPA install batchfile to P/Shell in the function  (hp-openview)
            1.15       6 October    Lawrence       Added option to update the DNS addresses (set-DNS) for all domains 
                                                   Added Current Std username from there A account as default to the Email Dialog. (emailcreds)
            1.16      7 October     Lawrence       Added function "set-pagefilesize" to set Pagefile tobe 4096MB as recommended by M$  
            1.17      7 October     Lawrence       Added function "set-firewall-enabled" to make sure firewall is enabled and started.
            1.18     13 October     LAwrence       Added Global Variables for HP OpenView so i could copy the Niche Openview parm.mwc file over.  
            1.19     14 October     Lawrence       Add Web-Net-Ext45 & Web-Asp-Net45 to the AIP function
            1.20     19 October     Lawrence       Added " install-windowsFeature Web-ASP" to AIP Function
            1.21     20 October     Lawrence       aDDED acds nds SHARE FUNCTION
            1.22     27 October     Lawrence       Added TrustedForDelegation rights to Server for AIP Servers only.
            1.23      5 November    Lawrence       Added SQL format 64k block size for d and E     ** format64k-de   
                                                   Added Setting the Eventlog Sizing and override options.  ** setEventlogsettings
                                                   Added NET-WCF-HTTP-Activation45 feature install to AIP  install 
#>

$ScriptVersion = "Version 1.23"

#checks to see if running as Admin
clear
Write-host  "`n`n       Script Version is $ScriptVersion `n`n`n" -ForegroundColor Yellow
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{    
    write-host  "`n************************************************************`
  `n* This script needs to be run As Admin                     * `
  `n* e.g. powershell.exe d:\scripts\apps.ps1 -Verb runAs or   *`
  `n* Run PowerShell as in an escalated window                 *`
  `n************************************************************`n`n" -ForegroundColor yellow
    Break
}

function Promptforbackupday
{
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
    $x = ""
    $objForm = New-Object System.Windows.Forms.Form 
    $objForm.Text = "Server Description"
    $objForm.Size = New-Object System.Drawing.Size(300, 200) 
    $objForm.StartPosition = "CenterScreen"

    $objForm.KeyPreview = $True
    $objForm.Add_KeyDown( { if ($_.KeyCode -eq "Enter") 
            { $x = $objTextBox.Text; $objForm.Close() } })
    $objForm.Add_KeyDown( { if ($_.KeyCode -eq "Escape") 
            { $objForm.Close() } })

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size(75, 120)
    $OKButton.Size = New-Object System.Drawing.Size(75, 23)
    $OKButton.Text = "OK"
    $OKButton.Add_Click( { $x = $objTextBox.Text; $objForm.Close() })
    $objForm.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(150, 120)
    $CancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $CancelButton.Text = "Cancel"
    $CancelButton.Add_Click( { $objForm.Close() })
    $objForm.Controls.Add($CancelButton)

    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.Location = New-Object System.Drawing.Size(10, 20) 
    $objLabel.Size = New-Object System.Drawing.Size(280, 30) 
    $objLabel.Text = "Please enter the Day this server needs to be backup up. Monday, tuesday etc :"
    $objForm.Controls.Add($objLabel) 
    $x = $objTextBox.Text

    $objTextBox = New-Object System.Windows.Forms.TextBox 
    $objTextBox.Location = New-Object System.Drawing.Size(10, 60) 
    $objTextBox.Size = New-Object System.Drawing.Size(260, 100) 
    $objForm.Controls.Add($objTextBox) 

    $objForm.Topmost = $True

    $objForm.Add_Shown( { $objForm.Activate(); $objTextBox.focus() })
    [void] $objForm.ShowDialog()
    $global:Backupday = $objTextBox.Text

}

function PromptandsetDescription
{
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
    $x = ""
    $objForm = New-Object System.Windows.Forms.Form 
    $objForm.Text = "Server Description"
    $objForm.Size = New-Object System.Drawing.Size(300, 200) 
    $objForm.StartPosition = "CenterScreen"
    $objForm.KeyPreview = $True
    $objForm.Add_KeyDown( { if ($_.KeyCode -eq "Enter") 
            { $x = $objTextBox.Text; $objForm.Close() } })
    $objForm.Add_KeyDown( { if ($_.KeyCode -eq "Escape") 
            { $objForm.Close() } })

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size(75, 120)
    $OKButton.Size = New-Object System.Drawing.Size(75, 23)
    $OKButton.Text = "OK"
    $OKButton.Add_Click( { $x = $objTextBox.Text; $objForm.Close() })
    $objForm.Controls.Add($OKButton)
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(150, 120)
    $CancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $CancelButton.Text = "Cancel"
    $CancelButton.Add_Click( { $objForm.Close() })
    $objForm.Controls.Add($CancelButton)
    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.Location = New-Object System.Drawing.Size(10, 20) 
    $objLabel.Size = New-Object System.Drawing.Size(280, 30) 
    $objLabel.Text = "Please enter the Server Description / Type. `n Eg QPRIME N in the space below:"
    $objForm.Controls.Add($objLabel) 
    $x = $objTextBox.Text
    $objTextBox = New-Object System.Windows.Forms.TextBox 
    $objTextBox.Location = New-Object System.Drawing.Size(10, 60) 
    $objTextBox.Size = New-Object System.Drawing.Size(260, 100) 
    $objForm.Controls.Add($objTextBox) 
    $objForm.Topmost = $True
    $objForm.Add_Shown( { $objForm.Activate(); $objTextBox.focus() })
    [void] $objForm.ShowDialog()
    $Descriptionx = $objTextBox.Text
    $pcdesc = Get-WmiObject Win32_operatingSystem 
    $pcdesc.description = $Descriptionx
    $pcdesc.put()  2>&1 | Out-Null
    $Global:description = $Descriptionx
}

Function emailcreds
{
    $findusername = $env:username
    $pattern = ’[^1234567890]’
    $emailusername = ($findusername –replace $pattern, ’’) -as [string]
    $emailusername
    $global:emailcreds = Get-Credential $emailusername  -Message "Enter your PRDS standard user code. Do not enter Domain.  This is required to send email only."
    $Global:emailaddress = $Emailcreds.username + "@intra.police.qld.gov.au"

}

Function renamenetworkadapter
{
    write-host  "`n Renaming the Default network adapter to be CORE" -ForegroundColor green
    Get-NetAdapter | Rename-NetAdapter -NewName "Core"
}

Function admincreds
{
    $prompt = "prds\" + ($($env:USERNAME)) 
    $Global:creds = Get-Credential $prompt -Message "Enter your PRDS Admin credentials to connet to the servers in PRDS to install the APPS."

    $username = $creds.username
    $password = $creds.GetNetworkCredential().password

    # Get current domain using logged-on user's credentials
    $PRDSDomain = "LDAP://DC=prds,DC=qldpol"
    $domain = New-Object System.DirectoryServices.DirectoryEntry($PRDSDomain, $UserName, $Password)

    if ($domain.name -eq $null)
    {
        write-host "Authentication failed - please verify your username and password."
        admincreds
    }
    else
    {
        write-host "Successfully authenticated with domain $PRDSDomain"
    }
}

function attachenetdrives
{
    write-host  "`n Attaching to network drive for Apps to install" -ForegroundColor green
    New-PSDrive -name deploy -PSProvider FileSystem -Root \\deploy.prds.qldpol\Install -Credential $creds 2>&1 | Out-Null
    New-PSDrive -name networker -PSProvider FileSystem -Root \\QPS-NTW-PR-01.prds.qldpol\Networker -Credential $creds 2>&1 | Out-Null
    New-PSDrive -name Sophos -PSProvider FileSystem -Root \\cit-sav-pr-01.prds.qldpol\SophosUpdate -Credential $creds 2>&1 | Out-Null
    New-PSDrive -name sccm-PSProvider FileSystem -Root \\qps-mgt-pr-41.prds.qldpol\Client$ -Credential $creds 2>&1 | Out-Null
}

function disablehibernation
{
    write-host  "`n Disabling hibernation" -ForegroundColor green
    powercfg -h off 2>&1 | Out-Null
}

function adminpassnotexpire
{
    write-host  "`n Setting Password never expires on local Administrator" -ForegroundColor green
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
                $user.SetInfo()
                $name = $user.FullName | out-string
                $name
                Write-Host "The Local Administrator account has been set with Password Never Expried"  $name   -ForegroundColor "black" -BackgroundColor "yellow"
            }
        }
        catch
        {
            Write-Host "FAILED to set the Local Administrator account to Password Never Expries" -ForegroundColor "Red" -BackgroundColor "Black"

        }
    }
}
    
function iehomepage
{   
    write-host  "`n Setting IE homepage" -ForegroundColor green
    REG ADD   "HKEY_USERS\.DEFAULT\Software\Microsoft\Internet Explorer\Main" /V "START PAGE" /D "About:blank" /F  2>&1 | Out-Null
    #  set-ItemProperty  -path 'HKEY_USER\.DEFAULT\Software\Microsoft\Internet Explorer\Main' -name Build_Date -value $date
}

function emptyrecyclebin
{
    write-host  "`n Empting the Recycle Bin" -ForegroundColor green
    $Shell = New-Object -ComObject Shell.Application 
    $RecBin = $Shell.Namespace(0xA) 
    $RecBin.Items() | % { Remove-Item $_.Path -Recurse -Confirm:$false }
}
    
function cleareventlogs
{
    write-host  "`n Clearing all Eventlogs." -ForegroundColor green
    Get-WinEvent -ListLog * -Force | % { Wevtutil.exe cl $_.logname }

}

function networker
{
    write-host  "`n Installing Networker" -ForegroundColor green
    cp \\QPS-NTW-PR-01.prds.qldpol\Networker\NetWorker 8.1.1.6\nw811_win_x64\win_x64\networkr\ c:\temp\networker\ -recurse -force
    c:
    cd\
    cd .\temp\networker
    .\setup.exe /S /v" /qn /l*v filename.log INSTALLLEVEL=100 NW_INSTALLLEVEL=100l NW_FIREWALL_CONFIG=1 STARTSVC=1 setuptype=Install"  2>&1 | Out-Null
    start-sleep 180
    cd\
    rmdir c:\temp\networker -ErrorAction SilentlyContinue -Confirm:$false  -force -Recurse  2>&1 | Out-Null
}

function hpopenview
{
    write-host  "`n Installing HP OpenView" -ForegroundColor green
   
    $OVPAVERSION = "5.0"
    $OVPALOCALDIR = "C:\Program Files\HP\HP BTO Software"
    $OVPADATADIR = "C:\ProgramData\HP\HP BTO Software\Data"
    $OVPASETUPDIR = "c:\TEMP\hppa"

    #Checking to see if old Software existed.
    if ( $(Try { Test-Path "C:\Program Files (x86)\hp openview\bin\ovpacmd.exe" } Catch { $false }) )
    {
        write-host "C:\Program Files (x86)\hp openview\bin\ovpacmd.exe   exists..  "
        write-host "`n`n Stopping execution since HPPA is already installed.`n Please uninstall the old version and re run this script."   -ForegroundColor blue  -BackgroundColor DarkYellow
        break 
    }
    Else
    {
        write-host  "`n       QPS - Install or Upgrade HP Open View...`n`n       Installing HP Openview $OVPAVERSION ...." -ForegroundColor green
    }

    #coping the Source file to %temp%\hppa
    cp \\deploy.prds.qldpol\install\Installation Media - By Vendor\HP\HP Openview\HPPA 5.0 $OVPASETUPDIR -recurse -force
    c:
    cd $OVPASETUPDIR
    .\setup.exe /v /s /qn
    Start-Sleep 180
    Write-host "Replacing Standard settings with the with new settings"  -ForegroundColor green
    cP $OVPASETUPDIR\alarmdef-parm\*.mwc $OVPALOCALDIR\newconfig\ -force  2>&1 | Out-Null
    RM $OVPALOCALDIR\newconfig\alarmdef.mwc -force  2>&1 | Out-Null
    RM $OVPALOCALDIR\newconfig\Parm.mwc -force  2>&1 | Out-Null
    rm $OVPADATADIR\alarmdef.mwc -force  2>&1 | Out-Null
    rm $OVPADATADIR\Parm.mwc -force  2>&1 | Out-Null
    cp $OVPASETUPDIR\alarmdef-parm\*.mwc $OVPADATADIR -force  2>&1 | Out-Null
    #NDS  cp $OVPASETUPDIR\alarmdef-parm\*.mwc $OVPADATADIR\parm.mwc -force  2>&1 | Out-Null
    $global:hpOVPASETUPDIR = $OVPASETUPDIR
    $global:hpOVPADATADIR = $OVPADATADIR
    write-host  "Stopping new services"
    cd $OVPALOCALDIR\bin\
    .\ovpacmd.exe STOP all
    Write-host  "Cleanup settings"
    diskperf.exe -Y
    .\agsysdb.exe -ito off
    Write-host "Starting  services"
    .\ovpacmd.exe START all
    .\perfstat.exe  2>&1 | Out-Null
    cd \
    #commented out as i need file to do the Niche counters. 
    #rmdir $OVPASETUPDIR -ErrorAction SilentlyContinue -Confirm:$false  -force -Recurse  2>&1 | Out-Null
    Write-host "Finished installing/upgrading OVPA $OVPAVERSION ... "  -ForegroundColor green

}

function Sophos
{
    write-host  "`n Installing Sophos" -ForegroundColor green
    #\\cit-sav-pr-01.prds.qldpol\SophosUpdate\CIDs\S000\SAVSCFXP\setup.exe -mng yes -crt R -updp '\\cit-sav-pr-01.prds.qldpol\SophosUpdate\CIDs\S000\SAVSCFXP' -G '\citec-sav-pr-01\SESC95\Corporate Servers\Application Servers' -ni  2>&1 | Out-Null
    \\cit-sav-pr-01.prds.qldpol\SophosUpdate\CIDs\S147\SAVSCFXP\setup.exe -mng yes -crt R -updp 'http://cit-sav-pr-01.prds.qldpol/SophosUpdate/CIDs/S147/SAVSCFXP/' -G '\cit-sav-pr-01\SESC10\QPS Servers' -ni  2>&1 | Out-Null
    $global:sophosiconnfile = "C:\ProgramData\Sophos\AutoUpdate\Config\iconn.cfg"
    If (!(Test-Path -Path $sophosiconnfile))
    {
        write-host "$sophosiconnfile does not exist `n Waiting 30sec and will retry" 
        Start-Sleep 30
        sophos
    }
    Else
    {
        write-host "Updating $sophosiconnfile"
        cp \\deploy.prds.qldpol\Install\Installation Media - By Vendor\sophos\iconn.cfg c:\temp
        cp c:\temp\iconn.cfg "C:\ProgramData\Sophos\AutoUpdate\Config\"
    }
}

function drivepermissions
{
    write-host  "`n Updating the file permissions on d and e drive" -ForegroundColor green
    cacls d: /e /r everyone  2>&1 | Out-Null
    cacls d: /e /r "Creator Owner"  2>&1 | Out-Null
    cacls e: /e /r everyone 2>&1 | Out-Null
    cacls e: /e /r "Creator Owner" 2>&1 | Out-Null

}

Function backuprequestform
{

    write-host  "`n Processing and emailing the Backup Form to be submitted via Infra" -ForegroundColor green
    $computername = $env:COMPUTERNAME
    $bios = gwmi win32_Computersystem | select-object Model
    if ($bios.Model -like "*VMware*") { $Machinetype = "Virtual" }
    else { $Machinetype = "Physical" }
    $email = $env:tmp + "\" + $computername + "_backup_request.rtf"
    #write-host $Machinetype

    $pcdesc = Get-WmiObject Win32_operatingSystem
    $description = $pcdesc.Description

    $global:date = Get-Date -Format "dd-MMM-yyy"

    $drives = get-PSDrive -PSProvider FileSystem c, d, e | ft -auto @{Name = "Drives"; expression = { ($_.name) } },
    @{name = "Disk Size(GB) "; Expression = { "{0,8:N0}" -f ($_.free / 1gb + $_.used / 1gb) } },
    @{name = "Disk Size(GB)  "; Expression = { "{0,8:N0}" -f ($_.used / 1gb) } }

    "                  ******************************************"  | out-file $email
    "                  *         BACKUP REQUEST FORM            *"  | out-file $email -append
    "                  ******************************************"  | out-file $email -append
    " "  | out-file $email -append
    " "  | out-file $email -append
    "Date:    $date                    Infra:________________"  | out-file $email -append
    "  "  | out-file $email -append
    "Server name:               $computername   " | out-file $email -append
    "Note: One server per form. (unless they have the same backup requirements)"  | out-file $email -append
    " "  | out-file $email -append
    "Server Function/Application:$description " | out-file $email -append
    "System Owner:"  | out-file $email -append
    " "   | out-file $email -append
    "Server file system type:    Windows NTFS "  | out-file $email -append
    " "  | out-file $email -append
    "Server Type:                $Machinetype "  | out-file $email -append
    ""  | out-file $email -append
    "Server physical location:   _______________________________"  | out-file $email -append
    ""  | out-file $email -append
    "Microsoft Clustered:        NO      Cluster Name ___"  | out-file $email -append
    ""  | out-file $email -append
    "Backup Agents Required:     Windows Agent      SQL Agent: __ OTHER: __ "  | out-file $email -append
    ""  | out-file $email -append
    "Full Backup Schedule day:   $Backupday between Mon to Sun "  | out-file $email -append
    " "  | out-file $email -append
    "Backup Time:                Any time between 5PM and 7AM "  | out-file $email -append
    " "  | out-file $email -append

    "Important note:"  | out-file $email -append
    "1. The default backup level offered by PSG is Weekly Full / Incremental Other Days', with a 28 days retention period."  | out-file $email -append 
    "2. For Windows Server 2008, system state recovery (Recovery of the system state without the need to reinstall the OS) is supported by PSG"  | out-file $email -append
    " "  | out-file $email -append
    "Bare metal recovery is not supported by PSG."  | out-file $email -append
    " "  | out-file $email -append
    " "  | out-file $email -append
    " "  | out-file $email -append
    "Additional backup information: _____________________________"  | out-file $email -append
    " "  | out-file $email -append
    " "  | out-file $email -append
    $drives   | out-file $email -append
    " "  | out-file $email -append
    " "  | out-file $email -append
    $requester = $emailcreds.username
    $fullname = (NET USER $requester /DOMAIN | FIND /I "Full name ")
    "Requester $fullname"  | out-file $email -append
    "Date Requested: $date"  | out-file $email -append
    Send-MailMessage -From $emailaddress -Subject "Backup $computername" -To $emailaddress -Body "Backup request to forward to PSG" -Attachments $email -Credential $emailcreds -Port 25 -SmtpServer qps-xch-pr-03.prds.qldpol

}

function sccmclient
{
    write-host  "`n Installing SCCM Client to the system" -ForegroundColor green
    cp \\qps-mgt-pr-41.prds.qldpol\Client$\CustomDeployment\ "c:\temp\sccm"  -recurse -force
    c: 
    cd\
    cd .\temp\sccm 
    & '.\Deploy ConfigMgr 2012 R2 Client.exe' 2>&1 | Out-Null
    write-host  "`n Wainting for the SCCM installation to finish" -ForegroundColor green
    start-sleep 60
    rmdir c:\temp\sccm  -ErrorAction SilentlyContinue -Confirm:$false  -force -Recurse
}

function SETBGInfodate
{
    write-host  "`n Setting the Build Date for the BGI info" -ForegroundColor green
    $date = Get-Date -Format "dd-MMM-yyy"
    set-ItemProperty  -path HKLM:\SOFTWARE\Wow6432Node\QPS\SysInfo -name Build_Date -value $date
}

function qprimeNDSrequirements
{
    write-host  "`n Installing the Qprime NDS Windows Features" -ForegroundColor green
    Install-WindowsFeature NET-Framework-45-Core
    Install-windowsfeature NET-Framework-Core
    Install-Windowsfeature NLB 
    Install-Windowsfeature RSAT-NLB  

    #copy the Niche Openview Counter over.
    cp $hpOVPASETUPDIR\alarmdef-parm\nds*.mwc $hpOVPADATADIR\parm.mwc -force  2>&1 | Out-Null

}

function PRcreatendssharefolderpermissions
{ 
    write-host  "`n Creating NDS Share and folder with permissions for PR" -ForegroundColor green
    $PATH_2_folder = 'E:\Program Files\Niche\NicheRMS\ServerLog'
    $sharename = 'NicheServerLog'
    IF (!(TEST-PATH $PATH_2_folder))
    { 
        NEW-ITEM $PATH_2_folder -type Directory 
    }
    New-SmbShare -name $sharename -path $PATH_2_folder -ChangeAccess 'prds\QPS-P CIT-AUD-CLUPR01 LogLoaders'
    $acl = Get-Acl $PATH_2_folder
    $permission = "prds\QPS-P CIT-AUD-CLUPR01 LogLoaders", "Modify", 'ContainerInherit, ObjectInherit', 'None', 'Allow' 
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
    $acl.SetAccessRule($accessRule)
    $acl | Set-Acl $PATH_2_folder

    $acl = Get-Acl $PATH_2_folder
    $permission = "prds\QPS-P-Manage-IPS NDS Servers", "FullControl", 'ContainerInherit, ObjectInherit', 'None', 'Allow' 
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
    $acl.SetAccessRule($accessRule)
    $acl | Set-Acl $PATH_2_folder
}

function VTcreatendssharefolderpermissions
{ 
    write-host  "`n Creating NDS Share and folder with permissions for VT" -ForegroundColor green
    $PATH_2_folder = 'E:\Program Files\Niche\NicheRMS\ServerLog'
    $sharename = 'NicheServerLog'

    IF (!(TEST-PATH $PATH_2_folder))
    { 
        NEW-ITEM $PATH_2_folder -type Directory 
    }

    New-SmbShare -name $sharename -path $PATH_2_folder -ChangeAccess 'prds\CAPAdmin-CSV-PR', 'prds\QPS-P CIT-AUD-CLUPR01 LogLoaders'

    $acl = Get-Acl $PATH_2_folder
    $permission = "prds\CAPAdmin-PR", "FullControl", 'ContainerInherit, ObjectInherit', 'None', 'Allow' 
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
    $acl.SetAccessRule($accessRule)
    $acl | Set-Acl $PATH_2_folder

    $acl = Get-Acl $PATH_2_folder
    $permission = "prds\CAPAdmin-VT", "FullControl", 'ContainerInherit, ObjectInherit', 'None', 'Allow' 
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
    $acl.SetAccessRule($accessRule)
    $acl | Set-Acl $PATH_2_folder

    $acl = Get-Acl $PATH_2_folder
    $permission = "prds\QPS-P-Manage-IPS NDS Servers", "FullControl", 'ContainerInherit, ObjectInherit', 'None', 'Allow' 
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
    $acl.SetAccessRule($accessRule)
    $acl | Set-Acl $PATH_2_folder
}

function ACcreatendssharefolderpermissions
{ 
    write-host  "`n Creating NDS Share and folder with permissions for VT" -ForegroundColor green
    $PATH_2_folder = 'E:\Program Files\Niche\NicheRMS\ServerLog'
    $sharename = 'NicheServerLog'

    IF (!(TEST-PATH $PATH_2_folder))
    { 
        NEW-ITEM $PATH_2_folder -type Directory 
    }

    New-SmbShare -name $sharename -path $PATH_2_folder -ChangeAccess 'everyone'

    $acl = Get-Acl $PATH_2_folder
    $permission = "ACDS\CAPAdmin-AC", "FullControl", 'ContainerInherit, ObjectInherit', 'None', 'Allow' 
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
    $acl.SetAccessRule($accessRule)
    $acl | Set-Acl $PATH_2_folder

    $acl = Get-Acl $PATH_2_folder
    $permission = "prds\QPS-P-Manage-IPS Servers", "FullControl", 'ContainerInherit, ObjectInherit', 'None', 'Allow' 
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
    $acl.SetAccessRule($accessRule)
    $acl | Set-Acl $PATH_2_folder
}

function TrustedForDelegation
{
    write-host  "`n Adding TrustedForDelegation rights to the Computer Account." -ForegroundColor green
    $Svrname = $env:COMPUTERNAME 
    set-ADcomputer $svrname -TrustedForDelegation 1 

}

function qprimeAIPrequirements
{
    write-host  "`n Installing the Qprime AIP Windows Features" -ForegroundColor green
    Install-WindowsFeature AS-TCP-Activation
    Install-WindowsFeature AS-web-support
    Install-WindowsFeature AS-Named-Pipes
    Install-WindowsFeature Web-Static-Content
    Install-WindowsFeature Web-Default-Doc
    Install-WindowsFeature Web-Asp-Net
    Install-WindowsFeature Web-Net-Ext
    Install-WindowsFeature Web-ISAPI-Ext
    Install-WindowsFeature Web-ISAPI-Filter
    Install-WindowsFeature Web-Basic-Auth
    Install-WindowsFeature Web-Windows-Auth
    Install-WindowsFeature Web-Filtering
    install-windowsfeature Web-Net-Ext45 
    install-windowsfeature Web-Asp-Net45
    Install-WindowsFeature NET-Framework-Core
    Install-WindowsFeature NET-Framework-45-Core
    Install-WindowsFeature NET-HTTP-Activation
    install-windowsfeature NET-WCF-HTTP-Activation45
    Install-WindowsFeature NET-Non-HTTP-Activ
    Install-WindowsFeature WAS-Process-Model
    Install-WindowsFeature WAS-NET-Environment
    Install-WindowsFeature WAS-Config-APIs
    Install-Windowsfeature RSAT-NLB
    Install-Windowsfeature NLB
    Install-WindowsFeature Web-Mgmt-Console
    Install-WindowsFeature Web-Http-Logging
    Install-WindowsFeature Web-Request-Monitor
    install-windowsFeature Web-ASP
    
    #Calls function to redir IIS Log Files 
    redirIISlogfiles
    TrustedForDelegation
}

Function SetDNS
{
    write-host  "`n Updating DNS Settings to the 4 main entries" -ForegroundColor green
    $Dom_name = $ENV:USERDOMAIN

    if ($Dom_name -like "PRDS")
    {
        Set-DNSClientServerAddress –interfaceIndex 12 –ServerAddresses (“10.47.1.201", "164.112.162.142", "10.46.253.23", "164.112.162.141")
    }
    if ($Dom_name -like "DVDS")
    {
        Set-DNSClientServerAddress -interfaceIndex 12 -ServerAddresses ("164.112.62.250", "164.112.162.249")
    }
    if ($Dom_name -like "DTDS")
    {
        Set-DNSClientServerAddress -interfaceIndex 12 -ServerAddresses ("10.46.193.164", "10.47.2.89")
    }
    if ($Dom_name -like "ACDS")
    {
        Set-DNSClientServerAddress -interfaceIndex 12 -ServerAddresses ("164.112.162.99", "164.112.162.230")
    }

}

function redirIISlogfiles
{
    write-host  "`n Relocating IIS SiteDefault log files to Drive E:" -ForegroundColor green
    NEW-ITEM 'E:\inetpub\logs\LogFiles' -type Directory 
    Set-WebConfigurationProperty -Filter System.Applicationhost/Sites/SiteDefaults/logfile -name Directory -value 'E:\inetpub\logs\LogFiles'

}

Function removelegalnotice
{
    remove-itemproperty -path HKLM:\Software\Microsoft\Windows\CurrentVersion\policies\system\ -name legalnoticecaption
    remove-itemproperty -path HKLM:\Software\Microsoft\Windows\CurrentVersion\policies\system\ -name legalnoticetext

}

function setpagefilesize
{
    write-host  "`n Making sure Page file is set to 4096MB" -ForegroundColor green
    Set-CimInstance -Query "Select * from win32_computersystem" -Property @{automaticmanagedpagefile = "False" }
    Set-CimInstance -Query "Select * from win32_PageFileSetting" -Property @{InitialSize = 4096; MaximumSize = 4096 }

}

function setfirewallenabled
{

    write-host  "`n Making sure Windows Firewall is enabled and started" -ForegroundColor green
    set-service -Name MpsSvc -StartupType Automatic 
    Start-Service -Name MpsSvc
}
    
function securityinf
{
    Start-Sleep 10
    cp \\deploy.prds.qldpol\install\Installation Media - By Vendor\InI\Security\Windows Server 2012 R2\ "C:\Windows\security\Windows Server 2012 R2"  -recurse -force
    c:
    Cd\
    cd '.\Windows\security\Windows Server 2012 R2'
    .\ISS-Sec-Impl-Quiet.cmd
    start-sleep 120
    Restart-Computer -force
}

function DISkCD
{   
    write-host  "`n making sure CD Rom is set to drive X:" -ForegroundColor green
    #making sure the CD Rom Drive is set as Drive X:
    Get-CimInstance -Query "SELECT * FROM Win32_Volume WHERE Drivetype ='5'" | Set-CimInstance -Arguments @{DriveLetter = "x:" }
}

function diskdande
{
    write-host  "`n Creating D and E drives. Setting Volume Labels" -ForegroundColor green
    #gets the RAW disks attached and sets them up
    $VolumeNumber = (Get-Disk | where { $_.partitionstyle -eq 'raw' }).Number
    Clear-Disk -Number $VolumeNumber -RemoveData -confirm:$false  2>&1 | Out-Null
    Initialize-Disk $VolumeNumber -PartitionStyle GPT -PassThru 2>&1 | Out-Null
    Clear-Disk -Number $VolumeNumber -RemoveData -confirm:$false 2>&1 | Out-Null
    Initialize-Disk $VolumeNumber -PartitionStyle GPT -PassThru 2>&1 | Out-Null
    #This stops the Prompting for format drive Dialog
    Stop-Service ShellHWDetection
    
    foreach ($number in $VolumeNumber)
    {
        New-Partition -DiskNumber $Number -AssignDriveLetter -UseMaximumSize | Format-Volume -FileSystem NTFS -Confirm:$false -Force  2>&1 | Out-Null
    }
    Start-Service ShellHWDetection
  
  
    ###   making small drive d

    $D = Get-WmiObject Win32_LogicalDisk -filter "DeviceID='d:'" | select-object size
    $DriveD_size = $D.size
    $driveD_size

    $D = Get-WmiObject Win32_LogicalDisk -filter "DeviceID='E:'" | select-object size
    $DriveE_size = $E.size
    $driveE_size

    if ($DriveE_size -le $DriveD_size)
    {

        Get-CimInstance -Query "SELECT * FROM Win32_Volume WHERE driveletter='d:'" | Set-CimInstance -Arguments @{DriveLetter = "h:" }
        Get-CimInstance -Query "SELECT * FROM Win32_Volume WHERE driveletter='e:'" | Set-CimInstance -Arguments @{DriveLetter = "i:" }
        Get-CimInstance -Query "SELECT * FROM Win32_Volume WHERE driveletter='h:'" | Set-CimInstance -Arguments @{DriveLetter = "e:" }
        Get-CimInstance -Query "SELECT * FROM Win32_Volume WHERE driveletter='i:'" | Set-CimInstance -Arguments @{DriveLetter = "d:" }
    }

    Set-Volume -DriveLetter d -NewFileSystemLabel 'Applications'
    Set-Volume -DriveLetter e -NewFileSystemLabel 'Data'
}

Function format64k-de
{
    Format-Volume -DriveLetter D -AllocationUnitSize 65536 -FileSystem NTFS  -NewFileSystemLabel Applications -Confirm:$false -Force
    Format-Volume -DriveLetter E -AllocationUnitSize 65536 -FileSystem NTFS  -NewFileSystemLabel Data -Confirm:$false -Force

}

function setEventlogsettings
{
    write-host "`nSetting the Event Log Sizing and override settings." -ForegroundColor green

    limit-eventLog -logname "Application" -MaximumSize 51200KB -OverflowAction OverwriteAsNeeded
    limit-eventLog -logname "Security" -MaximumSize 81920KB -OverflowAction OverwriteAsNeeded
    limit-eventLog -logname "System" -MaximumSize 51200KB -OverflowAction OverwriteAsNeeded

    $evtsettings = Get-Eventlog -list | Where-Object { ($_.LogDisplayName -eq 'Application') -or `
        ($_.LogDisplayName -eq 'Security') -or `
        ($_.LogDisplayName -eq 'System') } | ft  log, MaximumKilobytes, OverflowAction -AutoSize | Out-String
    Write-host $evtsettings

    write-host  "`n Clearing all Eventlogs as part of the Build" -ForegroundColor green
    Get-WinEvent -ListLog * -Force | % { Wevtutil.exe cl $_.logname }
    New-EventLog –LogName Application –Source “INI Build”  2>&1 | Out-Null
    Write-EventLog –LogName Application –Source “INI Build” –EntryType Information –EventID 1  –Message “Log Event log cleaned by  :$env:username”

}

<#
Functions that are run to configure the server
Comment any out below if you don't want to run them
by using a hash
#>

removelegalnotice
emailcreds
admincreds
PromptandsetDescription
Promptforbackupday
attachenetdrives
DISKCD
diskdande
#uncomment below for SQL Servers
#format64k-de

sccmclient
backuprequestform
renamenetworkadapter
SetDNS
disablehibernation
hpopenview
adminpassnotexpire

#Uncomment line below if the server is a NDS Server
#qprimeNDSrequirements
#PRcreatendssharefolderpermissions
#VTcreatendssharefolderpermissions
#ACcreatendssharefolderpermissions

#Uncomment line below if the server is a AIP QPRIME Integration Server
#qprimeAIPrequirements

iehomepage
SetBGInfoDate
emptyrecyclebin
networker
sophos
drivepermissions
setpagefilesize
setfirewallenabled
setEventlogsettings
securityinf

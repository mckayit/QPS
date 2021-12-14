function softwwareinstalled
{
    $reportfile = "c:\temp\apps.txt"
    Write-host "Processing..Checking sOFTWARE Instalations settings."-ForegroundColor green
    ######   works out if Oracle Client is installed"
    #new-item c:\temp\apps.txt -force -type file | out-null
   
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

            
    $appsinst = "List of APPS Installed. " 
    $appsinst1 = "------------------------------------------" 
    
    
    $apps1 = "`n64 BIT Apps`n"
    $apps1 = $apps1 += $smallline + "`n"

    $apps1 = $apps1 += $oracleversion64
    $apps1 = $apps1 += Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | sort displayname | Select-Object DisplayName, Publisher | ft  | Out-String
    $apps1 = $apps1 + "`n32 BIT Apps`n"
    $apps1 = $apps1 += $oracleversion32
    $apps1 = $apps1 += (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, Publisher | ft   | Out-String) 
    #$apps1 >> c:\temp\apps.txt 
    


   
    $blank | Out-file  $reportfile -append
    $linetop | Out-file  $reportfile -append
    $appsinst | Out-file  $reportfile -Append
    $appsinst1 | Out-file  $reportfile -Append
    $linebottom | Out-file  $reportfile -append
    $apps1 | Out-file  $reportfile -Append   
   

}
function repeat-string([string]$str, [int]$repeat) { $str * $repeat }



function get-logonhistory
{
    Param (
        [string]$Computer = ($env:CLIENTNAME),
        [int]$Days = 360
    )
    cls
    $Result = @()
    Write-Host "Gathering Event Logs, this can take awhile..."
    $ELogs = Get-EventLog System -Source Microsoft-Windows-WinLogon -After (Get-Date).AddDays(-$Days)# -ComputerName $Computer
    If ($ELogs)
    {
        Write-Host "Processing..."
        ForEach ($Log in $ELogs)
        {
            If ($Log.InstanceId -eq 7001)
            {
                $ET = "Logon"
            }
            <#
            ElseIf ($Log.InstanceId -eq 7002)
            {
                $ET = "Logoff"
            }#>
            Else
            {
                Continue
            }
            $Result += New-Object PSObject -Property @{
                Time         = $Log.TimeWritten
                'Event Type' = $ET
                User         = (New-Object System.Security.Principal.SecurityIdentifier $Log.ReplacementStrings[1]).Translate([System.Security.Principal.NTAccount])
            }
        }
        $global:1 = $Result | Select Time, "Event Type", User |  Sort Time -Descending  #| Out-GridView
        Write-Host "Done."
    }
    Else
    {
        Write-Host "Problem with $Computer."
        Write-Host "If you see a 'Network Path not found' error, try starting the Remote Registry service on that computer."
        Write-Host "Or there are no logon/logoff events (XP requires auditing be turned on)"
    }
}


get-logonhistory -Computer "computername" -Days 360
Write-output 'Users last logged in (Last 300 days if available in Security eventlog)' | out-file c:\temp\apps.txt 
$global:1 | ft -auto | out-file c:\temp\apps.txt -Append
softwwareinstalled
c:\temp\apps.txt

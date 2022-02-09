function get-logonhistory
{
    Param (
        #endregion[string]$Computer = (Read-Host Remote computer name),
        [int]$Days = 3600
    )
    cls
    $Result = @()
    Write-Host "Gathering Event Logs, this can take awhile..." -ForegroundColor green
    $ELogs = Get-EventLog System -Source Microsoft-Windows-WinLogon -After (Get-Date).AddDays(-$Days) #-ComputerName $Computer
    If ($ELogs)
    {
        Write-Host "Processing..."
        ForEach ($Log in $ELogs)
        {
            If ($Log.InstanceId -eq 7001)
            {
                $ET = "Logon"
            }
            ElseIf ($Log.InstanceId -eq 7002)
            {
                $ET = "Logoff"
            }
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
        $Result | Select Time, "Event Type", User |  Sort Time -Descending #| Out-GridView
        Write-Host "Done."
        $Result | Select Time, "Event Type", User |  Sort Time -Descending >c:\temp\info.txt 
    }
    Else
    {
        Write-Host "Problem with $Computer."
        Write-Host "If you see a 'Network Path not found' error, try starting the Remote Registry service on that computer."
        Write-Host "Or there are no logon/logoff events (XP requires auditing be turned on)"
    }


}

function get-apps_on_server
{
    $ln20 = [char]0x2550
    $smallline = repeat-string $ln20 12
    $Apps1 = "------------ Software installed ------------ "
    $apps1 = $apps1 += $smallline + "`n"
    $apps1 = $apps1 + "64 BIT Apps`n"
    $apps1 = $apps1 += $smallline + "`n"
    $apps1 = $apps1 += Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | sort displayname | Select-Object DisplayName, Publisher | ft  | Out-String
    $apps1 = $apps1 += $smallline + "`n"
    $apps1 = $apps1 + "`n32 BIT Apps`n"
    $apps1 = $apps1 += $smallline + "`n"
    $apps1 = $apps1 += $oracleversion32
    $apps1 = $apps1 += (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, Publisher | ft   | Out-String) 
    $apps1 >> c:\temp\info.txt 
    Write-host $apps1

}
Function get-Uptime
{
    $reportfile = 'c:\temp\info.txt'
    Write-host "Last time System was rebooted Check" -ForegroundColor Green

    $Uptimetitle = "    Last time System was rebooted Check "
    $wmi = Get-WmiObject -Class Win32_OperatingSystem
    $lastreboot = ($wmi.ConvertToDateTime($wmi.LastBootUpTime)) | Out-String
    $uptimecheck = "This Server was last rebooted $lastreboot  "  
    $blank | Out-file  $reportfile -append
    $linetop | Out-file  $reportfile -append
    $Uptimetitle | Out-file  $reportfile -append

    $linebottom | Out-file  $reportfile -append
    Test-PendingReboot
    $uptimecheck | Out-file  $reportfile -append
}    
get-uptime
get-logonhistory -Days 360
get-apps_on_server
notepad c:\temp\info.txt 
Param(
[Parameter(Mandatory=$true)]
[string]$Configuration,
[Parameter(Mandatory=$true)]
[string]$VMOutputCSV,
[Parameter(Mandatory=$true)]
[string]$VMDiskOutputCSV,
[Parameter(Mandatory=$true)]
[string]$VMNetOutputCSV,
[string]$LoggingLocation = "E:\Logs\DCVM-Reporting"

)


if (!(Test-Path -Path $LoggingLocation)) { throw "The logging location '$LoggingLocation' is not available."}

$LogFile = Join-Path -Path $LoggingLocation -ChildPath $("DCVM-DataCollector-" + $(Get-Date -Format 'yyyy-MM-dd') + ".log")

Start-Transcript -Path $LogFile

Write-Host "Importing vCenter QA Credential"
$credentialsStore = Get-Content -Path $Configuration -ErrorAction Stop | ConvertFrom-Csv -ErrorAction Stop


$credentialsStore | %{
    $vcenterCred = New-Object System.Management.Automation.PSCredential($_.name,($_.encpassword | ConvertTo-SecureString))
    
    Connect-VIServer -Server $_.vCenterTarget -Credential $vcenterCred -Force -ErrorAction Continue


}
$Counter = 0
$DateStarted = Get-Date

$allVMObject = Get-VM

$vmReportingTable = @()
$vmDiskReportingTable = @()
$vmNetworkReportingTable = @()

foreach ($v in $allVMObject) {

    
    $Counter = $Counter + 1;  if ( $($Counter % 1) -eq 0 ) { Write-Progress -Activity "Generating VM records..." -PercentComplete $(($Counter / $allVMObject.Count) * 100) -CurrentOperation "$Counter of $($allVMObject.Count)" }

    
    $vmReportingItem = [PSCustomObject]@{}


    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'VM UUID' -Value $($v.PersistentId)
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'VM' -Value $($v.Name)
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'Powerstate' -Value $($v.PowerState.ToString())
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'Template' -Value $($v.ExtensionData.Guest.GuestFullName)
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'DNS Name' -Value $($v.ExtensionData.Guest.HostName)
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'Connection state' -Value $($v.ExtensionData.Runtime.ConnectionState)
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'Guest state' -Value $($v.ExtensionData.Guest.GuestState )
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'Heartbeat' -Value $($v.ExtensionData.GuestHeartbeatStatus)
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'PowerOn' -Value $($v.PowerState.ToString())
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'CPUs' -Value $(($v.NumCpu * $v.CoresPerSocket))
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'Memory' -Value $($v.MemoryGB)
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'NICs' -Value $(@($v | Get-NetworkAdapter).Count)
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'Disks' -Value $null # $(($v.Guest.Disks | ConvertTo-Json))
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'Primary IP Address' -Value $($v.ExtensionData.Guest.IpAddress)
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'Network #1' -Value $(($v | Get-NetworkAdapter | select -First 1).NetworkName)
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'Resource pool' -Value $(($v | Get-ResourcePool).Name)
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'Path' -Value $($v.ExtensionData.Config.Files.VmPathName)
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'Log directory' -Value $($v.ExtensionData.Config.Files.LogDirectory)
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'Annotation' -Value $($v.Notes -replace "\n", " " )
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'Datacenter' -Value $(($v | Get-Datacenter).Name)
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'Cluster' -Value $(($v | Get-Cluster).Name)
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'Host' -Value $(($v | Get-VMHost).Name)
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'VI SDK Server' -Value $($v.Uid.Split('/')[1].Split('@')[1])
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'OS config file' -Value $($null)
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'DateLastSeen' -Value $DateStarted
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'IsVM' -Value $true
    $vmReportingItem | Add-member -MemberType NoteProperty -Name 'WorkLoadType' -Value "VMWare"  # Added this So the Column is there for import and helps with the import for CMDB



    Write-Verbose -Message $vmReportingItem
    $vmReportingTable += $vmReportingItem

    

    $disks = $v.Guest.Disks
    

    foreach ($d in $disks)
    {
        $vmDiskReportingItem = [PSCustomObject]@{}
        $vmDiskReportingItem | Add-member -MemberType NoteProperty -Name 'VM UUID' -Value $($v.PersistentId)
        $vmDiskReportingItem | Add-member -MemberType NoteProperty -Name 'CapacityGB' -Value $d.CapacityGB
        $vmDiskReportingItem | Add-member -MemberType NoteProperty -Name 'FreeSpaceGB' -Value $d.FreeSpaceGB
        $vmDiskReportingItem | Add-member -MemberType NoteProperty -Name 'Path' -Value $d.Path
        
        $vmDiskReportingTable += $vmDiskReportingItem
    }

    
    #vmNetworkReportingTable

    $nets = @($v | Get-NetworkAdapter)
    
    foreach ($n in $nets)
    {
        $vmNetReportingItem = [PSCustomObject]@{}
        $vmNetReportingItem | Add-member -MemberType NoteProperty -Name 'VM UUID' -Value $($v.PersistentId)
        $vmNetReportingItem | Add-member -MemberType NoteProperty -Name 'Name' -Value $n.Name
        $vmNetReportingItem | Add-member -MemberType NoteProperty -Name 'MacAddress' -Value $n.MacAddress
        $vmNetReportingItem | Add-member -MemberType NoteProperty -Name 'Type' -Value $n.Type
        $vmNetReportingItem | Add-member -MemberType NoteProperty -Name 'Connected' -Value $n.ConnectionState.Connected
        
        $vmNetworkReportingTable += $vmNetReportingItem
    }




}

Stop-Transcript

Write-Progress -Activity "Generating VM records..." -Completed

if ($vmReportingTable.Count -ge 1) { $vmReportingTable | ConvertTo-CSV -NoTypeInformation | Out-File -FilePath $VMOutputCSV -Force } else { Write-Warning -Message "No new records to add to VM reporting table." }

if ($vmDiskReportingTable.Count -ge 1) { $vmDiskReportingTable | ConvertTo-CSV -NoTypeInformation | Out-File -FilePath $VMDiskOutputCSV -Force } else { Write-Warning -Message "No new records to add to VM disk reporting table." }

if ($vmNetworkReportingTable.Count -ge 1) { $vmNetworkReportingTable | ConvertTo-CSV -NoTypeInformation | Out-File -FilePath $VMNetOutputCSV -Force } else { Write-Warning -Message "No new records to add to VM network reporting table." }

Disconnect-VIServer -Server * -Confirm:$false -ErrorAction SilentlyContinue
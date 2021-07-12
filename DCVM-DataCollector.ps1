<#
 .SYNOPSIS
    Exports all VMware guest info and outputs to Array
.DESCRIPTION
    Reades in Configuration file that contains list of Hosts as well as Creds
    From this list it gets all the VM guests info and Outputs to array | Exported to file.

.PARAMETER Configuration
    This is the Location of the Configuration file that contains the List of V-Centers ans well as Encrypted Creds.
    This is a Mandatory Value
    Contents of file is in the following CSV format
        vCenterTarget,domain,name,encpassword
    
.PARAMETER VMOutputCSV
    Specifies the location for the Output file for the VMGuest as a CSV
    This is a Mandatory Value

.PARAMETER VMDiskOutputCSV
    Specifies the location for the Output file for the VMDisks as a CSV
    This is a Mandatory Value

.PARAMETER VMNetOutputCSV
    Specifies the location for the Output file for the VMNetwork as a CSV
    This is a Mandatory Value


.EXAMPLE
    C:\PS>DCVM-DataCollector.ps1 -Configuration -VMOutputCSV  -VMDiskOutputCSV -VMNetOutputCSV 
    C:\PS>DCVM-DataCollector.ps1 -Configuration D:\Scripts\DC-VM-Reporting\ServAccCredsEncrypted.csv -VMOutputCSV "E:\VMExportData\VMExport.csv" -VMDiskOutputCSV "E:\VMExportData\VMDiskExport.csv" -VMNetOutputCSV "E:\VMExportData\VMNetExport.csv"
    Example of how to use this cmdlet
    
.OUTPUTS
    Outputs are all exported as a CSV File.  These are using the Above Var's
.NOTES
    This Script is dumping all the Info from the V-Centers and putting them in a location for CMDB usage.  
    As we don't have SCCM rolled out.
    This is a Workaround. 
    
    
    Date:    15 Feb 2021
      
     ******* Update Version number below when a change is done.*******
     
    History
    Version         Date                Name           Detail
    ---------------------------------------------------------------------------------------
    0.0.1           9 Feb 2021         xxxx       Initial Coding

#>
     

[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true,
        HelpMessage = 'This is the File that has the V-Centers and Creds')]
    [string]$Configuration,
    [Parameter(Mandatory = $true,
        HelpMessage = 'This is the File that has the V-Centers and Creds')]
    [string]$VMOutputCSV,
    [Parameter(Mandatory = $true,
        HelpMessage = 'This is the File that is exported to for the VMGuest info')]
    [string]$VMDiskOutputCSV,
    [Parameter(Mandatory = $true,
        HelpMessage = 'This is the File that is Exported to for the VMDisks')]
    [string]$VMNetOutputCSV,
    [string]$LoggingLocation = "E:\Logs\DCVM-Reporting"  #This is the Location for the Transcript that is setup below.

)

begin
{
    # Setting up the Logging. Making sure the Log Location Exists.
    if (!(Test-Path -Path $LoggingLocation)) { throw "The logging location '$LoggingLocation' is not available." }

    $LogFile = Join-Path -Path $LoggingLocation -ChildPath $("DCVM-DataCollector-" + $(Get-Date -Format 'yyyy-MM-dd') + ".log")

    # Starting the Transcript / Log for this
    Start-Transcript -Path $LogFile
}


Process 
{
    Write-Host "Importing vCenter QA Credential"
    $credentialsStore = Get-Content -Path $Configuration -ErrorAction Stop | ConvertFrom-Csv -ErrorAction Stop

    # Connecting to all the V-Centers from the Configuration file
    $credentialsStore | foreach-object {
        $vcenterCred = New-Object System.Management.Automation.PSCredential($_.name, ($_.encpassword | ConvertTo-SecureString))
    
        Connect-VIServer -Server $_.vCenterTarget -Credential $vcenterCred -Force -ErrorAction Continue
    }

    # Setting up the Environment
    $Counter = 1
    $DateStarted = Get-Date

    $allVMObject = Get-VM

    $vmReportingTable = @()
    $vmDiskReportingTable = @()
    $vmNetworkReportingTable = @()


    #Getting all the VM guest info.     
    foreach ($v in $allVMObject)
    {
        # Counter / Progress Bar.
        $paramWriteProgress = @{
            Activity = 'Generating VM records...'
            Status = "Processing [$counter] of [$($allVMObject.Count)] VM-Guests."
            PercentComplete = (($counter / $allVMObject.Count) * 100)
        
            Write-Progress @paramWriteProgress
                        
            $counter++       
            
            # Setting up Array of all the VM-Guest data wanted    
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


            Write-Verbose -Message $vmReportingItem

            # Appending the PS object from above to the Final Export  $vmReportingTable
            $vmReportingTable += $vmReportingItem

    
            # Getting All the Disk info for the VM Guest
            $disks = $v.Guest.Disks
    
            foreach ($d in $disks)
            {
                $vmDiskReportingItem = [PSCustomObject]@{}
                $vmDiskReportingItem | Add-member -MemberType NoteProperty -Name 'VM UUID' -Value $($v.PersistentId)
                $vmDiskReportingItem | Add-member -MemberType NoteProperty -Name 'CapacityGB' -Value $d.CapacityGB
                $vmDiskReportingItem | Add-member -MemberType NoteProperty -Name 'FreeSpaceGB' -Value $d.FreeSpaceGB
                $vmDiskReportingItem | Add-member -MemberType NoteProperty -Name 'Path' -Value $d.Path
        
                #Appending the PS object from above to the Final Export  $vmDiskReportingTable
                $vmDiskReportingTable += $vmDiskReportingItem
            }

    
            # Getting All the Network info for the VM Guest
            # vmNetworkReportingTable
            $nets = @($v | Get-NetworkAdapter)
    
            foreach ($n in $nets)
            {
                $vmNetReportingItem = [PSCustomObject]@{}
                $vmNetReportingItem | Add-member -MemberType NoteProperty -Name 'VM UUID' -Value $($v.PersistentId)
                $vmNetReportingItem | Add-member -MemberType NoteProperty -Name 'Name' -Value $n.Name
                $vmNetReportingItem | Add-member -MemberType NoteProperty -Name 'MacAddress' -Value $n.MacAddress
                $vmNetReportingItem | Add-member -MemberType NoteProperty -Name 'Type' -Value $n.Type
                $vmNetReportingItem | Add-member -MemberType NoteProperty -Name 'Connected' -Value $n.ConnectionState.Connected
        
                #Appending the PS object from above to the Final Export  $vmNetworkReportingTable
                $vmNetworkReportingTable += $vmNetReportingItem
            }
        }
    }
}
    
End 
{
    ##  This is where all the Output to files happen

    Write-Progress -Activity "Generating VM records..." -Completed

    if ($vmReportingTable.Count -ge 1)
    { 
        $vmReportingTable | Export-csv -NoTypeInformation -Path $VMOutputCSV -Force 
    } 
    else
    { 
        Write-Warning -Message "No new records to add to VM reporting table." 
    }

    if ($vmDiskReportingTable.Count -ge 1)
    { 
        $vmDiskReportingTable | Export-csv -NoTypeInformation -Path $VMDiskOutputCSV -Force 
    } 
    else 
    { 
        Write-Warning -Message "No new records to add to VM disk reporting table."
    }

    if ($vmNetworkReportingTable.Count -ge 1)
    {
        $vmNetworkReportingTable | Export-csv -NoTypeInformation -Path $VMNetOutputCSV -Force 
    } 
    else 
    { 
        Write-Warning -Message "No new records to add to VM network reporting table." 
    }

    #Cleans up all the V-Center Connections
    Disconnect-VIServer -Server * -Confirm:$false -ErrorAction SilentlyContinue
    
    #Lets flush the MEM after writing the Large Data Sets to memory before Export.
    $vmNetworkReportingTable = @()
    $vmDiskReportingTable = @()
    $vmReportingTable = @()
    
    # End Transcript so we have Logging 
    Stop-Transcript
}
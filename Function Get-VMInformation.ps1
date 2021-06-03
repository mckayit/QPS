Function Get-VMInformation
{
    <#
.SYNOPSIS
    Get information from a VM object. Properties inlcude Name, PowerState, vCenterServer, Datacenter, Cluster, VMHost, Datastore, Folder, GuestOS, NetworkName, IPAddress, MacAddress, VMTools
 
 
.NOTES  
    Name: Get-VMInformation
    Author: theSysadminChannel
    Version: 1.0
    DateCreated: 2019-Apr-29
 
 
.EXAMPLE
    For updated help and examples refer to -Online version.
 
 
.LINK
    https://thesysadminchannel.com/get-vminformation-using-powershell-and-powercli -
     
#>
 
    [CmdletBinding()]
 
    param(
        [Parameter(
            Position = 0,
            ParameterSetName = "NonPipeline"
        )]
        [Alias("VM")]
        [string[]]  $Name,
 
 
        [Parameter(
            Position = 1,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            ParameterSetName = "Pipeline"
        )]
        [PSObject[]]  $InputObject
 
    )
 
 
    BEGIN
    {
        if (-not $Global:DefaultVIServer)
        {
            Write-Error "Unable to continue.  Please connect to a vCenter Server." -ErrorAction Stop
        }
 
        #Verifying the object is a VM
        if ($PSBoundParameters.ContainsKey("Name"))
        {
            $InputObject = Get-VM $Name
        }
 
        $i = 1
        $Count = $InputObject.Count
    }
 
    PROCESS
    {
        if (($null -eq $InputObject.VMHost) -and ($null -eq $InputObject.MemoryGB))
        {
            Write-Error "Invalid data type. A virtual machine object was not found" -ErrorAction Stop
        }
 
        foreach ($Object in $InputObject)
        {
            try
            {
                $vCenter = $Object.Uid -replace ".+@"; $vCenter = $vCenter -replace ":.+"
                [PSCustomObject]@{
                    Name                                   = $Object.Name
                    PowerState                             = $Object.PowerState
                    vCenter                                = $vCenter
                    Datacenter                             = $Object.VMHost | Get-Datacenter | select -ExpandProperty Name
                    Cluster                                = $Object.VMhost | Get-Cluster | select -ExpandProperty Name
                    VMHost                                 = $Object.VMhost
                    VMHostModel                            = $Object.MODEL 
                    VMHostManufacturer                     = $Object.VMhost.Manufacturer 
                    VMHostNumCPU                           = $Object.VMhost.NimCPU
                    VMHostCpuTotalMhz                      = $Object.VMhost.CpuTotalMhz
                    VMHostCpuUsageMhz                      = $Object.VMhost.CpuUsageMhz
                    VMHostMemoryTotalMB                    = $Object.VMhost.MemoryTotalMB
                    VMHostMemoryTotalGB                    = $Object.VMhost.MemoryTotalGB
                    VMHostMemoryUsageMB                    = $Object.VMhost.MemoryUsageMB
                    VMHostMemoryUsageGB                    = $Object.VMhost.MemoryUsageGB
                    VMHostProcessorType                    = $Object.VMhost.ProcessorType
                    VMHostHyperthreadingActive             = $Object.VMhost.HyperthreadingActive

                    Datastore                              = ($Object | Get-Datastore | select -ExpandProperty Name) -join ', '
                    FolderName                             = $Object.Folder
                    GuestOS                                = $Object.ExtensionData.Config.GuestFullName
                    NetworkName                            = ($Object | Get-NetworkAdapter | select -ExpandProperty NetworkName) -join ', '
                    IPAddress                              = ($Object.ExtensionData.Summary.Guest.IPAddress) -join ', '
                    MacAddress                             = ($Object | Get-NetworkAdapter | select -ExpandProperty MacAddress) -join ', '
                    VMTools                                = $Object.ExtensionData.Guest.ToolsVersionStatus2
                    GuestStatsOverallCpuUsage              = $Object.ExtensionData.Summary.quickstats.OverallCpuUsage
                    GuestStatsOverallCpuDemand             = $Object.ExtensionData.Summary.quickstats.OverallCpuDemand
                    GuestStatsGuestMemoryUsage             = $Object.ExtensionData.Summary.quickstats.GuestMemoryUsage
                    GuestStatsHostMemoryUsage              = $Object.ExtensionData.Summary.quickstats.HostMemoryUsage
                    GuestStatsGuestHeartbeatStatus         = $Object.ExtensionData.Summary.quickstats.GuestHeartbeatStatus
                    GuestStatsDistributedCpuEntitlement    = $Object.ExtensionData.Summary.quickstats.DistributedCpuEntitlement
                    GuestStatsDistributedMemoryEntitlement = $Object.ExtensionData.Summary.quickstats.DistributedMemoryEntitlement
                    GuestStatsStaticCpuEntitlement         = $Object.ExtensionData.Summary.quickstats.StaticCpuEntitlement
                    GuestStatsStaticMemoryEntitlement      = $Object.ExtensionData.Summary.quickstats.StaticMemoryEntitlement
                    GuestStatsPrivateMemory                = $Object.ExtensionData.Summary.quickstats.PrivateMemory
                }
 
            }
            catch
            {
                Write-Error $_.Exception.Message
 
            }
            finally
            {
                if ($PSBoundParameters.ContainsKey("Name"))
                {
                    $PercentComplete = ($i / $Count).ToString("P")
                    Write-Progress -Activity "Processing VM: $($Object.Name)" -Status "$i/$count : $PercentComplete Complete" -PercentComplete $PercentComplete.Replace("%", "")
                    $i++
                }
                else
                {
                    Write-Progress -Activity "Processing VM: $($Object.Name)" -Status "Completed: $i"
                    $i++
                }
            }
        }
    }
 
    END {}
}
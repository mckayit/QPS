# Setting parameters for the connection
#[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None") ]
#Param(
    # Nutanix cluster IP address
    #[Parameter(Mandatory = $true)]
    #[Alias('IP')] [string] $nxIP,    
    # Nutanix cluster username
    #[Parameter(Mandatory = $true)]
    #[Alias('User')] [string] $nxUser,
    # Nutanix cluster password
    #[Parameter(Mandatory = $true)]
    #[Alias('Password')] [String] $nxPassword
#)

$nxIP = 'qcr-psm-pr-01'
$nxUser = 'ahv-collector'
$nxPassword = 'Sparonk11$'

# Converting the password to a secure string which isn't accepted for our API connectivity
$nxPasswordSec = ConvertTo-SecureString $nxPassword -AsPlainText -Force
Function write-log {
    <#
    .Synopsis
    Write logs for debugging purposes
    .Description
    This function writes logs based on the message including a time stamp for debugging purposes.
    #>
    param (
    $message,
    $sev = "INFO"
    )
    if ($sev -eq "INFO") {
        write-host "$(get-date -format "hh:mm:ss") | INFO | $message"
    }
    elseif ($sev -eq "WARN") {
        write-host "$(get-date -format "hh:mm:ss") | WARN | $message"
    }
    elseif ($sev -eq "ERROR") {
        write-host "$(get-date -format "hh:mm:ss") | ERROR | $message"
    }
    elseif ($sev -eq "CHAPTER") {
        write-host "`n`n### $message`n`n"
    }
} 
# Adding PS cmdlets
Add-PSSnapin -Name NutanixCmdletsPSSnapin 

Function Get-Hosts {
    <#
    .Synopsis
    This function will collect the hosts within the specified cluster.
    .Description
    This function will collect the hosts within the specified cluster using REST API call based on Invoke-RestMethod
    #>
    Param (
    [string] $debug
    )
    $credPair = "$($nxUser):$($nxPassword)"
    $encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($credPair))
    $headers = @{ Authorization = "Basic $encodedCredentials" }
    $URL = "https://$($nxIP):9440/api/nutanix/v3/hosts/list"
    $Payload = @{
        kind   = "host"
        offset = 0
        length = 999
    } 
    $JSON = $Payload | convertto-json
    try {
        $task = Invoke-RestMethod -Uri $URL -method "post" -body $JSON -ContentType 'application/json' -headers $headers;
    }
    catch {
        Start-Sleep 10
        write-log -message "Going once"
        $task = Invoke-RestMethod -Uri $URL -method "post" -body $JSON -ContentType 'application/json' -headers $headers;
    }
    write-log -message "We found $($task.entities.count) hosts in this cluster."
    Return $task
} 
Function Get-VMs {
    <#
    .Synopsis
    This function will collect the VMs within the specified cluster.
    .Description
    This function will collect the VMs within the specified cluster using REST API call based on Invoke-RestMethod
    #>
    Param (
    [string] $debug
    )
    $credPair = "$($nxUser):$($nxPassword)"
    $encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($credPair))
    $headers = @{ Authorization = "Basic $encodedCredentials" }
    write-log -message "Executing VM List Query"
    $URL = "https://$($nxIP):9440/api/nutanix/v3/vms/list"
    $Payload = @{
        kind   = "vm"
        offset = 0
        length = 999
        } 
    $JSON = $Payload | convertto-json
    try {
        $task = Invoke-RestMethod -Uri $URL -method "post" -body $JSON -ContentType 'application/json' -headers $headers;
        }
    catch {
        Start-Sleep 10
        write-log -message "Going once"
        $task = Invoke-RestMethod -Uri $URL -method "post" -body $JSON -ContentType 'application/json' -headers $headers;
        }
    write-log -message "We found $($task.entities.count) VMs."
    Return $task
} 
Function Get-DetailVM {
    <#
    .Synopsis
    This function will collect the speficics of the VM we've specified using the Get-VMs function as input.
    .Description
    This function will collect the speficics of the VM we've specified using the Get-VMs function as input using REST API call based on Invoke-RestMethod
    #>
    Param (
    [string] $uuid,
    [string] $debug
    )
    $credPair = "$($nxUser):$($nxPassword)"
    $encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($credPair))
    $headers = @{ Authorization = "Basic $encodedCredentials" }
    $URL = "https://$($nxIP):9440/api/nutanix/v3/vms/$($uuid)"
    try {
        $task = Invoke-RestMethod -Uri $URL -method "get" -headers $headers;
    }
    catch {
        Start-Sleep 10
        write-log -message "Going once"
    }  
    Return $task
} 
Function Get-DetailHosts {
    Param (
    [string] $uuid,
    [string] $debug
    )
    $credPair = "$($nxUser):$($nxPassword)"
    $encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($credPair))
    $headers = @{ Authorization = "Basic $encodedCredentials" }
    $URL = "https://$($nxIP):9440/api/nutanix/v3/hosts/$($uuid)"
    try {
        $task = Invoke-RestMethod -Uri $URL -method "get" -headers $headers;
    }
    catch {
        Start-Sleep 10
        $task = Invoke-RestMethod -Uri $URL -method "get" -headers $headers;
        write-log -message "Going once"
    }  
    Return $task
} 
# Selecting all the GPUs and their devices IDs in the cluster
$GPU_List = $null
$hosts = Get-Hosts -ClusterPC_IP $nxIP -nxPassword $nxPassword -clusername $nxUser -debug $debug
Foreach ($Hypervisor in $hosts.entities) {
    $detail = Get-DetailHosts -ClusterPC_IP $nxIP -nxPassword $nxPassword -clusername $nxUser -debug $debug -uuid $Hypervisor.metadata.uuid
    [array]$GPU_List += $detail.status.resources.gpu_list
}
write-log -message "Collecting vGPU profiles and Device IDs"
# Connecting to the Nutanix Cluster
$nxServerObj = Connect-NTNXCluster -Server $nxIP -UserName $nxUser -Password $nxPasswordSec -AcceptInvalidSSLCerts -ForcedConnection
write-log -Message "Connecting to cluster $nxIp"
if ($null -eq (get-ntnxclusterinfo)) {
    write-log -message "Cluster connection isn't available, abborting the script"
    break
}
else {
    write-log -message "Connected to Nutanix cluster $nxIP"
}
# Fetching data and putting into CSV
$vms = @(get-ntnxvm | Where-Object {$_.controllerVm -Match "false"}) 
write-log -message "Grabbing VM information"
write-log -message "Currently grabbing information on $($vms.count) VMs"
$FullReport = @()
foreach ($vm in $vms) {                        
    $usedspace = 0
    if (!($vm.nutanixvirtualdiskuuids.count -le $null)) {
            write-log -message "Grabbing information on $($vm.vmName)"
            foreach ($UUID in $VM.nutanixVirtualDiskUuids) {
                $usedspace += (Get-NTNXVirtualDiskStat -Id $UUID -Metrics controller_user_bytes).values[0]
                $myvmdetail = Get-DetailVM -ClusterPC_IP $nxIP -nxPassword $nxPassword -clusername $nxUser -debug $debug -uuid $vm.uuid
            }
        }
    if ($vm.gpusInUse -eq "true") {
        
        $newVMObject = $myvmdetail
        $devid = $newVMObject.spec.resources.gpu_list
        $GPUUsed = $GPU_List | Where-Object {$_.device_id -eq $devid.device_id} 
        $VMGPU = $GPUUsed | select-object {$_.name} -unique
        $VMGPU1 = $VMGPU.'$_.name'
    }
    else {
        $VMGPU1 = $Null
    }
    if ($usedspace -gt 0) {
        $usedspace = [math]::round($usedspace / 1gb, 0)
    }
    $container = "NA"
    if (!($vm.vdiskFilePaths.count -le 0)) {
        $container = $vm.vdiskFilePaths[0].split('/')[1]
    }
    if ($vm.nutanixGuestTools.enabled -eq 'False') { $NGTstate = 'Installed'}
    else { 
        $NGTstate = 'Not Installed'
    }
    
    $props = [ordered]@{
        "VM uuid"                       = $vm.uuid
        "VM Name"                       = $vm.vmName
        "Creation Time"                 = ([Datetime]$myvmdetail.metadata.creation_time).ToUniversalTime().AddHours(10)
        "Container"                     = $container
        "Protection Domain"             = $vm.protectionDomainName
        "Host Placement"                = $vm.hostName
        "Power State"                   = $vm.powerstate
        "Network Name"                  = $myvmdetail.spec.resources.nic_list.subnet_reference.name
        "Network adapters"              = $vm.numNetworkAdapters
        "IP Address(es)"                = $vm.ipAddresses -join ","
        "vCPUs"                         = $vm.numVCpus
        "Number of Cores"               = $myvmdetail.spec.resources.num_sockets
        "Number of vCPUs per core"      = $myvmdetail.spec.resources.num_vcpus_per_socket
        "vRAM (GB)"                     = [math]::round($vm.memoryCapacityInBytes / 1GB, 0)
        "Disk Count"                    = $vm.nutanixVirtualDiskUuids.count
        "Provisioned Space (GB)"        = [math]::round($vm.diskCapacityInBytes / 1GB, 0)
        "Used Space (GB)"               = $usedspace
        "GPU Profile"                   = $VMGPU1
        "VM description"                = $vm.description
        "VM Time Zone"                  = $myvmdetail.spec.resources.hardware_clock_timezone
        "Nutanix Guest Tools installed" = $NGTState
        "NGT Version"                   = $vm.nutanixGuestTools.installedVersion
        "guestOperatingSystem"          = $vm.guestOperatingSystem   # information is not exposed when you are using AHV hypervisor.
        "hypervisorType"                = $vm.hypervisorType         # adds the Hypervisor type.
        "WorkLoadType"                  = "Nutanix"    # Added to help with filling in data in CMDB.
        } #End properties
    $Reportobject = New-Object PSObject -Property $props
    $fullreport += $Reportobject
}
if (!(Test-Path "C:\RVTools\Nutanix-Inventory")){New-Item -ItemType Directory -Path "D:\AHV-Rvtools\"}
#$Date = (Get-Date).tostring("yyyyMMdd")
$fullreport | Export-Csv -Path "D:\AHV-Rvtools\nutanixdata.csv" -NoTypeInformation -UseCulture -verbose:$false 
write-log -message "Writing the information to the CSV"
# Disconnecting from the Nutanix Cluster
#Disconnect-NTNXCluster -Servers *
write-log -message "Closing the connection to the Nutanix cluster $($nxIP)"

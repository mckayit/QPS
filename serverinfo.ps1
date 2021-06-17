#$servername = 'edwqpswsuspr01'
#$servername = 'POLqpsxchpr01'
#  $servers = Get-ADComputer -Filter 'operatingsystem -like "*server*" ' -Properties *
# $creds = Get-Credential

Function get-infozza
{
    $i = 1
    foreach ($server in $servers)
    {
        

        #progress bar.
        $paramWriteProgress = @{
            Activity         = "getting Server info"
            Status           = "Processing [$i] of [$($Servers.Count)] users"
            PercentComplete  = (($i / $Servers.Count) * 100)
            CurrentOperation = "processing the following Server : [ $($server.name)]"
        }
        Write-Progress @paramWriteProgress       
        $i++

    
        [string]$servername = $server.Name


        # get-manufacture

        $manufacture = get-wmiobject -computername $servername win32_computersystem #-Credential $creds

        $Server_Name = $manufacture.name
        #$Domain_ = $manufacture.domain
        $domain_ = ($server.dnshostname.trim($server.name) ).trimstart(".")
        $Manufacturer = $manufacture.manufacturer
        $model = $manufacture.model
        <#
Domain              : prds.qldpol
Manufacturer        : VMware, Inc.
Model               : VMware Virtual Platform
#>


        # Get-OS

        $OS = (Get-WmiObject -ComputerName $servername -Class Win32_OperatingSystem)


        # Get-cpuinfo

        $cpu = Get-WmiObject -ComputerName $servername -Class Win32_Processor | Select-Object -Property Name, Number*



        $cpu_1name = $cpu[0].name
        $cpu_1NumberofCores = $cpu[0].NumberOfCores 
        $cpu_1NumberOfEnabledCore = $cpu[0].NumberOfEnabledCore 
        $cpu_1NumberOfLogicalProcessors = $cpu[0].NumberOfLogicalProcessors

        $cpu_2name = $cpu[1].name
        $cpu_2NumberofCores = $cpu[1].NumberOfCores 
        $cpu_2NumberOfEnabledCore = $cpu[1].NumberOfEnabledCore 
        $cpu_2NumberOfLogicalProcessors = $cpu[1].NumberOfLogicalProcessors

        $cpu_3name = $cpu[2].name
        $cpu_3NumberofCores = $cpu[2].NumberOfCores 
        $cpu_3NumberOfEnabledCore = $cpu[2].NumberOfEnabledCore 
        $cpu_3NumberOfLogicalProcessors = $cpu[2].NumberOfLogicalProcessors

        $cpu_4name = $cpu[3].name
        $cpu_4NumberofCores = $cpu[3].NumberOfCores 
        $cpu_4NumberOfEnabledCore = $cpu[3].NumberOfEnabledCore 
        $cpu_4NumberOfLogicalProcessors = $cpu[3].NumberOfLogicalProcessors

        $cpu_5name = $cpu[4].name
        $cpu_5NumberofCores = $cpu[4].NumberOfCores 
        $cpu_5NumberOfEnabledCore = $cpu[4].NumberOfEnabledCore 
        $cpu_5NumberOfLogicalProcessors = $cpu[4].NumberOfLogicalProcessors

        $cpu_6name = $cpu[5].name
        $cpu_6NumberofCores = $cpu[5].NumberOfCores 
        $cpu_6NumberOfEnabledCore = $cpu[5].NumberOfEnabledCore 
        $cpu_6NumberOfLogicalProcessors = $cpu[5].NumberOfLogicalProcessors

        $cpu_7name = $cpu[6].name
        $cpu_7NumberofCores = $cpu[6].NumberOfCores 
        $cpu_7NumberOfEnabledCore = $cpu[6].NumberOfEnabledCore 
        $cpu_7NumberOfLogicalProcessors = $cpu[6].NumberOfLogicalProcessors

        $cpu_8name = $cpu[7].name
        $cpu_8NumberofCores = $cpu[7].NumberOfCores 
        $cpu_8NumberOfEnabledCore = $cpu[7].NumberOfEnabledCore 
        $cpu_8NumberOfLogicalProcessors = $cpu[7].NumberOfLogicalProcessors



        #get MEM
        $mem = Get-WmiObject -ComputerName $servername -Class Win32_ComputerSystem 
        #$mem1 = $mem.TotalPhysicalMemory 
        $mem1 = [math]::Ceiling($mem.TotalPhysicalMemory / 1024 / 1024 / 1024)


        #get-nicinfo

        $nics = Get-WmiObject -ComputerName $servername  Win32_NetworkAdapterConfiguration



        $Nic1_Index = $nics[0].Index
        $Nic1_ServiceName = $nics[0].ServiceName
        $Nic1_Description = $nics[0].Description
        $Nic1_DHCPEnabled = $nics[0].DHCPEnabled
        $Nic1_IPAddress = $nics[0].IPAddress
        $Nic1_DefaultIPGateway = $nics[0].DefaultIPGateway
        $Nic1_DNSDomain = $nics[0].DNSDomain
        $Nic1_Subnet = $nics[0].ipsubnet
        $Nic1_MACAddress = $nics[0].MACAddress

 
        $Nic2_Index = $nics[1].Index
        $Nic2_ServiceName = $nics[1].ServiceName
        $Nic2_Description = $nics[1].Description
        $Nic2_DHCPEnabled = $nics[1].DHCPEnabled
        $Nic2_IPAddress = $nics[1].IPAddress
        $Nic2_DefaultIPGateway = $nics[1].DefaultIPGateway
        $Nic2_DNSDomain = $nics[1].DNSDomain
        $Nic2_Subnet = $nics[1].ipsubnet
        $Nic2_MACAddress = $nics[1].MACAddress


        $Nic3_Index = $nics[2].Index
        $Nic3_ServiceName = $nics[2].ServiceName
        $Nic3_Description = $nics[2].Description
        $Nic3_DHCPEnabled = $nics[2].DHCPEnabled
        $Nic3_IPAddress = $nics[2].IPAddress
        $Nic3_DefaultIPGateway = $nics[2].DefaultIPGateway
        $Nic3_DNSDomain = $nics[2].DNSDomain
        $Nic3_Subnet = $nics[2].ipsubnet
        $Nic3_MACAddress = $nics[2].MACAddress


        $Nic4_Index = $nics[3].Index
        $Nic4_ServiceName = $nics[3].ServiceName
        $Nic4_Description = $nics[3].Description
        $Nic4_DHCPEnabled = $nics[3].DHCPEnabled
        $Nic4_IPAddress = $nics[3].IPAddress
        $Nic4_DefaultIPGateway = $nics[3].DefaultIPGateway
        $Nic4_DNSDomain = $nics[3].DNSDomain
        $Nic4_Subnet = $nics[3].ipsubnet
        $Nic4_MACAddress = $nics[3].MACAddress


        $Nic5_Index = $nics[4].Index
        $Nic5_ServiceName = $nics[4].ServiceName
        $Nic5_Description = $nics[4].Description
        $Nic5_DHCPEnabled = $nics[4].DHCPEnabled
        $Nic5_IPAddress = $nics[4].IPAddress
        $Nic5_DefaultIPGateway = $nics[4].DefaultIPGateway
        $Nic5_DNSDomain = $nics[4].DNSDomain
        $Nic5_Subnet = $nics[4].ipsubnet
        $Nic5_MACAddress = $nics[4].MACAddress


        $Nic6_Index = $nics[5].Index
        $Nic6_ServiceName = $nics[5].ServiceName
        $Nic6_Description = $nics[5].Description
        $Nic6_DHCPEnabled = $nics[5].DHCPEnabled
        $Nic6_IPAddress = $nics[5].IPAddress
        $Nic6_DefaultIPGateway = $nics[5].DefaultIPGateway
        $Nic6_DNSDomain = $nics[5].DNSDomain
        $Nic6_Subnet = $nics[5].ipsubnet
        $Nic6_MACAddress = $nics[5].MACAddress

        $Nic7_Index = $nics[6].Index
        $Nic7_ServiceName = $nics[6].ServiceName
        $Nic7_Description = $nics[6].Description
        $Nic7_DHCPEnabled = $nics[6].DHCPEnabled
        $Nic7_IPAddress = $nics[6].IPAddress
        $Nic7_DefaultIPGateway = $nics[6].DefaultIPGateway
        $Nic7_DNSDomain = $nics[6].DNSDomain
        $Nic7_Subnet = $nics[6].ipsubnet
        $Nic7_MACAddress = $nics[6].MACAddress

        $Nic8_Index = $nics[7].Index
        $Nic8_ServiceName = $nics[7].ServiceName
        $Nic8_Description = $nics[7].Description
        $Nic8_DHCPEnabled = $nics[7].DHCPEnabled
        $Nic8_IPAddress = $nics[7].IPAddress
        $Nic8_DefaultIPGateway = $nics[7].DefaultIPGateway
        $Nic8_DNSDomain = $nics[7].DNSDomain
        $Nic8_Subnet = $nics[7].ipsubnet
        $Nic8_MACAddress = $nics[7].MACAddress

        $Nic9_Index = $nics[8].Index
        $Nic9_ServiceName = $nics[8].ServiceName
        $Nic9_Description = $nics[8].Description
        $Nic9_DHCPEnabled = $nics[8].DHCPEnabled
        $Nic9_IPAddress = $nics[8].IPAddress
        $Nic9_DefaultIPGateway = $nics[8].DefaultIPGateway
        $Nic9_DNSDomain = $nics[8].DNSDomain
        $Nic9_Subnet = $nics[8].ipsubnet
        $Nic9_MACAddress = $nics[8].MACAddress

        $Nic10_Index = $nics[9].Index
        $Nic10_ServiceName = $nics[9].ServiceName
        $Nic10_Description = $nics[9].Description
        $Nic10_DHCPEnabled = $nics[9].DHCPEnabled
        $Nic10_IPAddress = $nics[9].IPAddress
        $Nic10_DefaultIPGateway = $nics[9].DefaultIPGateway
        $Nic10_DNSDomain = $nics[9].DNSDomain
        $Nic10_Subnet = $nics[9].ipsubnet
        $Nic10_MACAddress = $nics[9].MACAddress
        
        $Nic11_Index = $nics[10].Index
        $Nic11_ServiceName = $nics[10].ServiceName
        $Nic11_Description = $nics[10].Description
        $Nic11_DHCPEnabled = $nics[10].DHCPEnabled
        $Nic11_IPAddress = $nics[10].IPAddress
        $Nic11_DefaultIPGateway = $nics[10].DefaultIPGateway
        $Nic11_DNSDomain = $nics[10].DNSDomain
        $Nic11_Subnet = $nics[10].ipsubnet
        $Nic11_MACAddress = $nics[10].MACAddress
        
        $Nic12_Index = $nics[11].Index
        $Nic12_ServiceName = $nics[11].ServiceName
        $Nic12_Description = $nics[11].Description
        $Nic12_DHCPEnabled = $nics[11].DHCPEnabled
        $Nic12_IPAddress = $nics[11].IPAddress
        $Nic12_DefaultIPGateway = $nics[11].DefaultIPGateway
        $Nic12_DNSDomain = $nics[11].DNSDomain
        $Nic12_Subnet = $nics[11].ipsubnet
        $Nic12_MACAddress = $nics[11].MACAddress
        
        
        $Nic13_Index = $nics[12].Index
        $Nic13_ServiceName = $nics[12].ServiceName
        $Nic13_Description = $nics[12].Description
        $Nic13_DHCPEnabled = $nics[12].DHCPEnabled
        $Nic13_IPAddress = $nics[12].IPAddress
        $Nic13_DefaultIPGateway = $nics[12].DefaultIPGateway
        $Nic13_DNSDomain = $nics[12].DNSDomain
        $Nic13_Subnet = $nics[12].ipsubnet
        $Nic13_MACAddress = $nics[12].MACAddress
        

  
        $Nic14_Index = $nics[13].Index
        $Nic14_ServiceName = $nics[13].ServiceName
        $Nic14_Description = $nics[13].Description
        $Nic14_DHCPEnabled = $nics[13].DHCPEnabled
        $Nic14_IPAddress = $nics[13].IPAddress
        $Nic14_DefaultIPGateway = $nics[13].DefaultIPGateway
        $Nic14_DNSDomain = $nics[13].DNSDomain
        $Nic14_Subnet = $nics[13].ipsubnet
        $Nic14_MACAddress = $nics[13].MACAddress
        
  
        $Nic15_Index = $nics[14].Index
        $Nic15_ServiceName = $nics[14].ServiceName
        $Nic15_Description = $nics[14].Description
        $Nic15_DHCPEnabled = $nics[14].DHCPEnabled
        $Nic15_IPAddress = $nics[14].IPAddress
        $Nic15_DefaultIPGateway = $nics[14].DefaultIPGateway
        $Nic15_DNSDomain = $nics[14].DNSDomain
        $Nic15_Subnet = $nics[14].ipsubnet
        $Nic15_MACAddress = $nics[14].MACAddress
        






        #GET-DISKS
        $drvs1 = get-wmiobject -ComputerName $servername win32_volume | where { $_.driveletter -ne 'X:' -and $_.label -ne 'System Reserved' } | sort driveletter 



        $DRVName1 = $drvs1[0].name
        $DRVLabel1 = $drvs1[0].label
        $DRVFileSystem1 = $drvs1[0].filesystem
        $DRVSizeGB1 = $drvs1[0].Capacity / 1gb
        $DRVFreeSizeGB1 = $drvs1[0].freespace / 1gb

        $DRVName2 = $drvs1[1].name
        $DRVLabel2 = $drvs1[1].label
        $DRVFileSystem2 = $drvs1[1].filesystem
        $DRVSizeGB2 = $drvs1[1].Capacity / 1gb
        $DRVFreeSizeGB2 = $drvs1[1].freespace / 1gb

        $DRVName3 = $drvs1[2].name
        $DRVLabel3 = $drvs1[2].label
        $DRVFileSystem3 = $drvs1[2].filesystem
        $DRVSizeGB3 = $drvs1[2].Capacity / 1gb
        $DRVFreeSizeGB3 = $drvs1[2].freespace / 1gb

        $DRVName4 = $drvs1[3].name
        $DRVLabel4 = $drvs1[3].label
        $DRVFileSystem4 = $drvs1[3].filesystem
        $DRVSizeGB4 = $drvs1[3].Capacity / 1gb
        $DRVFreeSizeGB4 = $drvs1[3].freespace / 1gb

        $DRVName5 = $drvs1[4].name
        $DRVLabel5 = $drvs1[4].label
        $DRVFileSystem5 = $drvs1[4].filesystem
        $DRVSizeGB5 = $drvs1[4].Capacity / 1gb
        $DRVFreeSizeGB5 = $drvs1[4].freespace / 1gb

        $DRVName6 = $drvs1[5].name
        $DRVLabel6 = $drvs1[5].label
        $DRVFileSystem6 = $drvs1[5].filesystem
        $DRVSizeGB6 = $drvs1[5].Capacity / 1gb
        $DRVFreeSizeGB6 = $drvs1[5].freespace / 1gb

        $DRVName7 = $drvs1[6].name
        $DRVLabel7 = $drvs1[6].label
        $DRVFileSystem7 = $drvs1[6].filesystem
        $DRVSizeGB7 = $drvs1[6].Capacity / 1gb
        $DRVFreeSizeGB7 = $drvs1[6].freespace / 1gb

        $DRVName8 = $drvs1[7].name
        $DRVLabel8 = $drvs1[7].label
        $DRVFileSystem8 = $drvs1[7].filesystem
        $DRVSizeGB8 = $drvs1[7].Capacity / 1gb
        $DRVFreeSizeGB8 = $drvs1[7].freespace / 1gb

        $DRVName9 = $drvs1[8].name
        $DRVLabel9 = $drvs1[8].label
        $DRVFileSystem9 = $drvs1[8].filesystem
        $DRVSizeGB9 = $drvs1[8].Capacity / 1gb
        $DRVFreeSizeGB9 = $drvs1[8].freespace / 1gb

        $DRVName10 = $drvs1[9].name
        $DRVLabel10 = $drvs1[9].label
        $DRVFileSystem10 = $drvs1[9].filesystem
        $DRVSizeGB10 = $drvs1[9].Capacity / 1gb
        $DRVFreeSizeGB10 = $drvs1[9].freespace / 1gb

        #output  
        [PSCustomObject]@{
            #$prop1 = [pscustomobject]@{
            ServerName                 = [string]$Server.name     
            Domain                     = [string]$domain_
            Manufacturer               = [string]$Manufacturer 
            model                      = [string]$model
            OSVersion                  = [string]$server.OperatingSystem
            OSArchitecture             = [string]$os.OSArchitecture
            MemoryGB                   = $mem1
            CanonicalName              = $server.CanonicalName
            ADWhenCreated              = $server.WhenCreated
            WhenModified               = $server.Modified
            ServerSID                  = $server.SID
            ADDNSHostName              = $server.DNSHostName


            CPUName1                   = $cpu_1name
            CPUCoreCount1              = $cpu_1NumberOfCores 
            NumberOfEnabledCore1       = $cpu_1NumberOfEnabledCore 
            NumberOfLogicalProcessors1 = $cpu_1NumberOfLogicalProcessors
                
            CPUName2                   = $cpu_2name
            CPUCoreCount2              = $cpu_2NumberOfCores 
            NumberOfEnabledCore2       = $cpu_2NumberOfEnabledCore 
            NumberOfLogicalProcessors2 = $cpu_2NumberOfLogicalProcessors

            CPUName3                   = $cpu_3name
            CPUCoreCount3              = $cpu_3NumberOfCores 
            NumberOfEnabledCore3       = $cpu_3NumberOfEnabledCore 
            NumberOfLogicalProcessors3 = $cpu_3NumberOfLogicalProcessors

            CPUName4                   = $cpu_4name
            CPUCoreCount4              = $cpu_4NumberOfCores 
            NumberOfEnabledCore4       = $cpu_4NumberOfEnabledCore 
            NumberOfLogicalProcessors4 = $cpu_4NumberOfLogicalProcessors

            CPUName5                   = $cpu_5name
            CPUCoreCount5              = $cpu_5NumberOfCores 
            NumberOfEnabledCore5       = $cpu_5NumberOfEnabledCore 
            NumberOfLogicalProcessors5 = $cpu_5NumberOfLogicalProcessors

            CPUName6                   = $cpu_6name
            CPUCoreCount6              = $cpu_6NumberOfCores 
            NumberOfEnabledCore6       = $cpu_6NumberOfEnabledCore 
            NumberOfLogicalProcessors6 = $cpu_6NumberOfLogicalProcessors

            CPUname7                   = $cpu_7name
            CPUCoreCount7              = $cpu_7NumberOfCores 
            NumberOfEnabledCore7       = $cpu_7NumberOfEnabledCore 
            NumberOfLogicalProcessors7 = $cpu_7NumberOfLogicalProcessors

            CPUName8                   = $cpu_8name
            CPUCoreCount8              = $cpu_8NumberOfCores 
            NumberOfEnabledCore8       = $cpu_8NumberOfEnabledCore 
            NumberOfLogicalProcessors8 = $cpu_8NumberOfLogicalProcessors

            Nic1_Index                 = $Nic1_Index 
            Nic1_ServiceName           = $Nic1_ServiceName
            Nic1_Description           = $Nic1_Description
            Nic1_DHCPEnabled           = $Nic1_DHCPEnabled
            Nic1_IPAddress             = [string]$Nic1_IPAddress  
            Nic1_DefaultIPGateway      = [string]$Nic1_DefaultIPGateway
            Nic1_DNSDomain             = $Nic1_DNSDomain
            Nic1_Subnet                = [string]$Nic1_Subnet
            Nic1_MACAddress            = [string]$Nic1_MACAddress

            Nic2_Index                 = $Nic2_Index 
            Nic2_ServiceName           = $Nic2_ServiceName
            Nic2_Description           = $Nic2_Description
            Nic2_DHCPEnabled           = $Nic2_DHCPEnabled
            Nic2_IPAddress             = [string]$Nic2_IPAddress  
            Nic2_DefaultIPGateway      = [string]$Nic2_DefaultIPGateway
            Nic2_DNSDomain             = $Nic2_DNSDomain
            Nic2_Subnet                = [string]$Nic2_Subnet
            Nic2_MACAddress            = [string]$Nic2_MACAddress

   	
            Nic3_Index                 = $Nic3_Index 
            Nic3_ServiceName           = $Nic3_ServiceName
            Nic3_Description           = $Nic3_Description
            Nic3_DHCPEnabled           = $Nic3_DHCPEnabled
            Nic3_IPAddress             = [string]$Nic3_IPAddress  
            Nic3_DefaultIPGateway      = [string]$Nic3_DefaultIPGateway
            Nic3_DNSDomain             = $Nic3_DNSDomain
            Nic3_Subnet                = [string]$Nic3_Subnet
            Nic3_MACAddress            = [string]$Nic3_MACAddress
   	
            Nic4_Index                 = $Nic4_Index 
            Nic4_ServiceName           = $Nic4_ServiceName
            Nic4_Description           = $Nic4_Description
            Nic4_DHCPEnabled           = $Nic4_DHCPEnabled
            Nic4_IPAddress             = [string]$Nic4_IPAddress  
            Nic4_DefaultIPGateway      = [string]$Nic4_DefaultIPGateway
            Nic4_DNSDomain             = $Nic4_DNSDomain
            Nic4_Subnet                = [string]$Nic4_Subnet
            Nic4_MACAddress            = [string]$Nic4_MACAddress
		
            Nic5_Index                 = $Nic5_Index
            Nic5_ServiceName           = $Nic5_ServiceName
            Nic5_Description           = $Nic5_Description
            Nic5_DHCPEnabled           = $Nic5_DHCPEnabled
            Nic5_IPAddress             = [string]$Nic5_IPAddress  
            Nic5_DefaultIPGateway      = [string]$Nic5_DefaultIPGateway
            Nic5_DNSDomain             = $Nic5_DNSDomain
            Nic5_Subnet                = [string]$Nic5_Subnet
            Nic5_MACAddress            = $Nic5_MACAddress

		
            Nic6_Index                 = $Nic6_Index 
            Nic6_ServiceName           = $Nic6_ServiceName
            Nic6_Description           = $Nic6_Description
            Nic6_DHCPEnabled           = $Nic6_DHCPEnabled
            Nic6_IPAddress             = [string]$Nic6_IPAddress  
            Nic6_DefaultIPGateway      = [string]$Nic6_DefaultIPGateway
            Nic6_DNSDomain             = $Nic6_DNSDomain
            Nic6_Subnet                = [string]$Nic6_Subnet
            Nic6_MACAddress            = [string]$Nic6_MACAddress

            Nic7_Index                 = $Nic7_Index 
            Nic7_ServiceName           = $Nic7_ServiceName
            Nic7_Description           = $Nic7_Description
            Nic7_DHCPEnabled           = $Nic7_DHCPEnabled
            Nic7_IPAddress             = [string]$Nic7_IPAddress  
            Nic7_DefaultIPGateway      = [string]$Nic7_DefaultIPGateway
            Nic7_DNSDomain             = $Nic7_DNSDomain
            Nic7_Subnet                = [string]$Nic7_Subnet
            Nic7_MACAddress            = [string]$Nic7_MACAddress
 	
            
            Nic8_Index                 = $Nic8_Index 
            Nic8_ServiceName           = $Nic8_ServiceName
            Nic8_Description           = $Nic8_Description
            Nic8_DHCPEnabled           = $Nic8_DHCPEnabled
            Nic8_IPAddress             = [string]$Nic8_IPAddress  
            Nic8_DefaultIPGateway      = [string]$Nic8_DefaultIPGateway
            Nic8_DNSDomain             = $Nic8_DNSDomain
            Nic8_Subnet                = [string]$Nic8_Subnet
            Nic8_MACAddress            = [string]$Nic8_MACAddress
            
            Nic9_Index                 = $Nic9_Index 
            Nic9_ServiceName           = $Nic9_ServiceName
            Nic9_Description           = $Nic9_Description
            Nic9_DHCPEnabled           = $Nic9_DHCPEnabled
            Nic9_IPAddress             = [string]$Nic9_IPAddress  
            Nic9_DefaultIPGateway      = [string]$Nic9_DefaultIPGateway
            Nic9_DNSDomain             = $Nic9_DNSDomain
            Nic9_Subnet                = [string]$Nic9_Subnet
            Nic9_MACAddress            = [string]$Nic9_MACAddress

            
            Nic10_Index                = $Nic10_Index 
            Nic10_ServiceName          = $Nic10_ServiceName
            Nic10_Description          = $Nic10_Description
            Nic10_DHCPEnabled          = $Nic10_DHCPEnabled
            Nic10_IPAddress            = [string]$Nic10_IPAddress  
            Nic10_DefaultIPGateway     = [string]$Nic10_DefaultIPGateway
            Nic10_DNSDomain            = $Nic10_DNSDomain
            Nic10_Subnet               = [string]$Nic10_Subnet
            Nic10_MACAddress           = [string]$Nic10_MACAddress
            
            Nic11_Index                = $Nic11_Index 
            Nic11_ServiceName          = $Nic11_ServiceName
            Nic11_Description          = $Nic11_Description
            Nic11_DHCPEnabled          = $Nic11_DHCPEnabled
            Nic11_IPAddress            = [string]$Nic11_IPAddress  
            Nic11_DefaultIPGateway     = [string]$Nic11_DefaultIPGateway
            Nic11_DNSDomain            = $Nic11_DNSDomain
            Nic11_Subnet               = [string]$Nic11_Subnet
            Nic11_MACAddress           = [string]$Nic11_MACAddress
            
            Nic12_Index                = $Nic12_Index 
            Nic12_ServiceName          = $Nic12_ServiceName
            Nic12_Description          = $Nic12_Description
            Nic12_DHCPEnabled          = $Nic12_DHCPEnabled
            Nic12_IPAddress            = [string]$Nic12_IPAddress  
            Nic12_DefaultIPGateway     = [string]$Nic12_DefaultIPGateway
            Nic12_DNSDomain            = $Nic12_DNSDomain
            Nic12_Subnet               = [string]$Nic12_Subnet
            Nic12_MACAddress           = [string]$Nic12_MACAddress
            
            Nic13_Index                = $Nic13_Index 
            Nic13_ServiceName          = $Nic13_ServiceName
            Nic13_Description          = $Nic13_Description
            Nic13_DHCPEnabled          = $Nic13_DHCPEnabled
            Nic13_IPAddress            = [string]$Nic13_IPAddress  
            Nic13_DefaultIPGateway     = [string]$Nic13_DefaultIPGateway
            Nic13_DNSDomain            = $Nic13_DNSDomain
            Nic13_Subnet               = [string]$Nic13_Subnet
            Nic13_MACAddress           = [string]$Nic13_MACAddress

            Nic14_Index                = $Nic14_Index 
            Nic14_ServiceName          = $Nic14_ServiceName
            Nic14_Description          = $Nic14_Description
            Nic14_DHCPEnabled          = $Nic14_DHCPEnabled
            Nic14_IPAddress            = [string]$Nic14_IPAddress  
            Nic14_DefaultIPGateway     = [string]$Nic14_DefaultIPGateway
            Nic14_DNSDomain            = $Nic14_DNSDomain
            Nic14_Subnet               = [string]$Nic14_Subnet
            Nic14_MACAddress           = [string]$Nic14_MACAddress
                        
            Nic15_Index                = $Nic15_Index 
            Nic15_ServiceName          = $Nic15_ServiceName
            Nic15_Description          = $Nic15_Description
            Nic15_DHCPEnabled          = $Nic15_DHCPEnabled
            Nic15_IPAddress            = [string]$Nic15_IPAddress  
            Nic15_DefaultIPGateway     = [string]$Nic15_DefaultIPGateway
            Nic15_DNSDomain            = $Nic15_DNSDomain
            Nic15_Subnet               = [string]$Nic15_Subnet
            Nic15_MACAddress           = [string]$Nic15_MACAddress
            
            
            
            DRVName1                   =	$DRVName1
            DRVLabel1                  =	$DRVLabel1
            DRVFileSystem1             =	$DRVFileSystem1
            DRVSizeGB1                 =	$DRVSizeGB1
            DRVFreeSizeGB1             =	$DRVFreeSizeGB1
    
            DRVName2                   =	$DRVName2
            DRVLabel2                  =	$DRVLabel2
            DRVFileSystem2             =	$DRVFileSystem2
            DRVSizeGB2                 =	$DRVSizeGB2
            DRVFreeSizeGB2             =	$DRVFreeSizeGB2
    	
            DRVName3                   =	$DRVName3
            DRVLabel3                  =	$DRVLabel3
            DRVFileSystem3             =	$DRVFileSystem3
            DRVSizeGB3                 =	$DRVSizeGB3
            DRVFreeSizeGB3             =	$DRVFreeSizeGB3
    	
            DRVName4                   =	$DRVName4
            DRVLabel4                  =	$DRVLabel4
            DRVFileSystem4             =	$DRVFileSystem4
            DRVSizeGB4                 =	$DRVSizeGB4
            DRVFreeSizeGB4             =	$DRVFreeSizeGB4
    	
            DRVName5                   =	$DRVName5
            DRVLabel5                  =	$DRVLabel5
            DRVFileSystem5             =	$DRVFileSystem5
            DRVSizeGB5                 =	$DRVSizeGB5
            DRVFreeSizeGB5             =	$DRVFreeSizeGB5
    	
            DRVName6                   =	$DRVName6
            DRVLabel6                  =	$DRVLabel6
            DRVFileSystem6             =	$DRVFileSystem6
            DRVSizeGB6                 =	$DRVSizeGB6
            DRVFreeSizeGB6             =	$DRVFreeSizeGB6
    	
            DRVName7                   =	$DRVName7
            DRVLabel7                  =	$DRVLabel7
            DRVFileSystem7             =	$DRVFileSystem7
            DRVSizeGB7                 =	$DRVSizeGB7
            DRVFreeSizeGB7             =	$DRVFreeSizeGB7
    	
            DRVName8                   =	$DRVName8
            DRVLabel8                  =	$DRVLabel8
            DRVFileSystem8             =	$DRVFileSystem8
            DRVSizeGB8                 =	$DRVSizeGB8
            DRVFreeSizeGB8             =	$DRVFreeSizeGB8
    	
            DRVName9                   =	$DRVName9
            DRVLabel9                  =	$DRVLabel9
            DRVFileSystem9             =	$DRVFileSystem9
            DRVSizeGB9                 =	$DRVSizeGB9 
            DRVFreeSizeGB9             =	$DRVFreeSizeGB9
    	
            DRVName10                  =	$DRVName10     
            DRVLabel10                 =	$DRVLabel10     
            DRVFileSystem10            =	$DRVFileSystem10
            DRVSizeGB10                =	$DRVSizeGB10    
            DRVFreeSizeGB10            =	$DRVFreeSizeGB10



        }

            
        #$prop1 
        #$prop1 |export-csv 'D:\tmp\out.csv'  
               
    }
                
                
}



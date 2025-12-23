try{
    # Import necessary modules
    Import-Module VMware.PowerCLI.VCenter  -ErrorAction Stop
    Import-Module ImportExcel -ErrorAction Stop
}
catch{
    Write-Output "required modules not found. Installing..."
     # Check if VMware PowerCLI and ImportExcel modules are installed, and install if not
    $modules = @("VMware.PowerCLI.VCenter", "ImportExcel")
    foreach ($module in $modules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-Output "$module not found. Installing..."
            Install-Module -Name $module -Scope CurrentUser -Force
        } else {
            Write-Output "$module is already installed."
        }
    }
}



# Set configuration to ignore invalid SSL certificates
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false

# Prompt for the vCenter server address
$vcenterServer = Read-Host -Prompt "Enter the vCenter server address"


# Prompt for credentials
$credential = Get-Credential

# Connect to vCenter using credentials
Connect-VIServer -Server $vcenterServer -Credential $credential

# Define the start and end dates for the last 30 days
$end = Get-Date
$start = $end.AddDays(-30)

# Define the output directory
$outputDir = "C:\vCenterStats"
if (-not (Test-Path -Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir
}
if(Test-Path -Path (Join-Path -Path $outputDir -ChildPath "VUtilizations.xlsx")){
    Remove-Item -Path (Join-Path -Path $outputDir -ChildPath "VUtilizations.xlsx") -Force
}

# Function to get average stat
function Get-AverageStat {
    param (
        [Parameter(Mandatory=$true)]
        [string]$stat,
        [Parameter(Mandatory=$true)]
        [PSObject]$entity
    )
    try {
        $stats = Get-Stat -Entity $entity -Stat $stat -Start $start -Finish $end
        if ($stats) {
            $minVal = $stats | Measure-Object -Property Value -Minimum | Select-Object -ExpandProperty Minimum         
            $maxVal = $stats | Measure-Object -Property Value -Maximum | Select-Object -ExpandProperty Maximum
            $averageVal = ($stats | Measure-Object -Property Value -Average).Average
            return ($minVal,  $maxVal, $averageVal)
        } else {
            return "N/A"
        }
    } catch {
        return "Error"
    }
}

try {

    # Retrieve and display Datacenter details with utilization over the last 30 days
    $datacenterList = Get-Datacenter
    $totalDatacenters = $datacenterList.Count
    #Write-Output ("Total Datacenters {0}" -f ($totalDatacenters))
    $datacenterDetails = @()
 
    for ($i = 0; $i -lt $totalDatacenters; $i++) {
        $datacenter = $datacenterList[$i]
        #Write-Output $i $datacenter.Name
        $progress = [math]::Round(($i / $totalDatacenters) * 100, 2)
        Write-Progress -Activity "Processing Datacenters" -Status "$($i)/$($totalDatacenters) Processing Datacenter $($datacenter.Name)" -PercentComplete $progress
 
        # Calculate CPU and Memory usage for the datacenter
        $datacenterVMs = Get-VM -Location $datacenter
        $clusterCount = Get-Cluster -Location $datacenter
        $hostCount = Get-VMHost -Location $datacenter

        # Retrieve the UUID from the Datacenter's ExtensionData         
        $datacenterUUID = $datacenter.ExtensionData.MoRef.Value
 
        $datacenterDetails += [PSCustomObject]@{
            Name                 = $datacenter.Name
            Id                   = $datacenter.Id
            UUID                 = $datacenterUUID
            'Total VMs'          = $datacenterVMs.Count
            'Total Clusters'     = $clusterCount.Count
            'Total Hosts'        = $hostCount.Count
        }

        # Retrieve and display Host details with utilization over the last 30 days
        $hostList = $datacenter | Get-VMHost
        $totalHosts = $hostList.Count
        $hostDetails = @()

        for ($j = 0; $j -lt $totalHosts; $j++) {
            $vmHost = $hostList[$j]
            #Write-Output ("{0} - {1}" -f ($j, $vmHost.Name))
            $progress = [math]::Round(($j / $totalHosts) * 100, 2)
            Write-Progress -Activity "Processing Hosts" -Status "$($j)/$($totalHosts) Processing Host $($vmHost.Name)" -PercentComplete $progress -id 1

            # Get the datastores accessible by the host
            $datastores = Get-Datastore -RelatedObject $vmHost         
            $totalCapacity = ($datastores | Measure-Object -Property CapacityGB -Sum).Sum         
            $totalFreeSpace = ($datastores | Measure-Object -Property FreeSpaceGB -Sum).Sum         
            $totalUsedSpace = $totalCapacity - $totalFreeSpace         
            $usedSpacePercent = if ($totalCapacity -ne 0) { [math]::Round(($totalUsedSpace / $totalCapacity) * 100, 2) } else { 0 }  
            # Retrieve the UUID from the Host's ExtensionData        
            $hostUUID = $vmHost.ExtensionData.Hardware.SystemInfo.Uuid
            $cluster = Get-Cluster -VMHost $vmHost

            $minCPUVal, $maxCPUVal, $averageCPUVal = Get-AverageStat -stat "cpu.usage.average" -entity $vmHost
            $minmemVal, $maxmemVal, $averagememVal = Get-AverageStat -stat  "mem.usage.average" -entity $vmHost
            
            $hostDetails += [PSCustomObject]@{     
                 Id                = $vmHost.Id
                 Datcenter         = $datacenter.Name 
                 'Datcenter Id'    = $datacenter.Id 
                 'Datcenter UUID'  = $datacenterUUID  
                 Name              = $vmHost.Name   
                 UUID              = $hostUUID  
                 'Cluster Name'    = $cluster.Name
                 'Average CPU Usage (%)'   = $averageCPUVal
                 'Min CPU Usage (%)'   = $minCPUVal
                 'Max CPU Usage (%)'   = $maxCPUVal         
                 'Average Memory Usage (%)'= $averagememVal
                 'Min Memory Usage (%)'= $minmemVal
                 'Max Memory Usage (%)'= $maxmemVal             
                 'Storage Capacity (GB)' = $totalCapacity             
                 'Storage Used (GB)' = $totalUsedSpace            
                 'Storage Used (%)' = $usedSpacePercent        
                 }

            # Retrieve and display VM details with utilization over the last 30 days
            $vmList = $vmHost | Get-VM
            $totalVMs = $vmList.Count
            $vmDetails = @()
            $vmNetworkDetails = @()
            for ($k = 0; $k -lt $totalVMs; $k++) {
                #if ($k -eq 1){break}
                $vm = $vmList[$k]
                #Write-Output $k $vm.Name
                $progress = [math]::Round(($k / $totalVMs) * 100, 2)
                Write-Progress -Activity "Processing VMs" -Status "$($k)/$($totalVMs) Processing VM $($vm.Name)" -PercentComplete $progress -id 2
 
                # Get the storage usage for each VM
                $vmHardDisks = Get-HardDisk -VM $vm
                $storageUsedGB = [math]::Round(($vmHardDisks | Measure-Object -Property CapacityKB -Sum).Sum / 1MB, 2)

                # Retrieve the UUID from the VM's ExtensionData         
                $vmUUID = $vm.ExtensionData.Config.InstanceUuid

                $minCPUVal, $maxCPUVal, $averageCPUVal = Get-AverageStat -stat "cpu.usage.average" -entity $vm
                $minmemVal, $maxmemVal, $averagememVal = Get-AverageStat -stat  "mem.usage.average" -entity $vm

                
                $vmDetails += [PSCustomObject]@{
                    id                  = $vm.Id
                    Name                = $vm.Name
                    'VM Host Id'        = $vmHost.Id
                    Datcenter           = $datacenter.Name 
                    'Datcenter Id'      = $datacenter.Id 
                    UUID                = $vmUUID
                    'Datcenter UUID'    = $datacenterUUID 
                    Host_UUID           = $hostUUID
                    Host                = $vm.VMHost
                    Host_ID             = $vm.VMHostId
                    PowerState          = $vm.PowerState
                    NumCpu              = $vm.NumCpu
                    MemoryMB            = $vm.MemoryMB
                    'Average CPU Usage (%)'   = $averageCPUVal
                    'Min CPU Usage (%)'   = $minCPUVal
                    'Max CPU Usage (%)'   = $maxCPUVal         
                    'Average Memory Usage (%)'= $averagememVal
                    'Min Memory Usage (%)'= $minmemVal
                    'Max Memory Usage (%)'= $maxmemVal 
                    'Storage Used (GB)' = [math]::Round($storageUsedGB, 2)
                }

                if ($vm.ExtensionData.Guest.ToolsStatus -eq 'toolsOk') {
                    foreach ($nic in $vm.ExtensionData.Guest.Net) {
                        if ($nic.IpConfig) {
                            foreach ($ip in $nic.IpConfig.IpAddress) {
                                if ($ip.IpAddress -and $ip.PrefixLength) {
                                    $vmNetworkDetails += [PSCustomObject]@{
                                        'VM Name'          = $vm.Name
                                        'Network Adapter'  = $nic.Network
                                        'IP Address'       = $ip.IpAddress
                                        'PrefixLength'     = $ip.PrefixLength
                                        }
                                }
                            }
                        }
                    }
                }

            }
            # Export VM details to CSV
            $vmDetails | Export-Excel -Path (Join-Path -Path $outputDir -ChildPath "VUtilizations.xlsx") -WorksheetName "VmUtilization" -AutoSize -Append
            $vmNetworkDetails | Export-Excel -Path (Join-Path -Path $outputDir -ChildPath "VUtilizations.xlsx") -WorksheetName "VNetwork" -AutoSize -Append


        }
        # Export Host details to CSV
        $hostDetails | Export-Excel -Path (Join-Path -Path $outputDir -ChildPath "VUtilizations.xlsx") -WorksheetName "VHostUtilization" -AutoSize -Append

    }

    # Export Datacenter details to CSV
    $datacenterDetails | Export-Excel -Path (Join-Path -Path $outputDir -ChildPath "VUtilizations.xlsx") -WorksheetName "VDatacenter" -AutoSize -Append

    
    # Retrieve and display Datastore details with utilization over the last 30 days
    $datastoreList = Get-Datastore
    $totalDatastores = $datastoreList.Count
    $datastoreDetails = @()

    for ($i = 0; $i -lt $totalDatastores; $i++) {
        $datastore = $datastoreList[$i]
        $progress = [math]::Round(($i / $totalDatastores) * 100, 2)
        # Retrieve the UUID from the Datastore's ExtensionData         
        $datastoreUUID = $datastore.ExtensionData.Info.Uuid
        Write-Progress -Activity "Processing Datastores" -Status "Processing Datastore $($datastore.Name)" -PercentComplete $progress
        $datastoreDetails += [PSCustomObject]@{
            Name           = $datastore.Name
            UUID           = $datastoreUUID
            CapacityGB     = $datastore.CapacityGB
            FreeSpaceGB    = $datastore.FreeSpaceGB
            UsedSpaceGB    = $datastore.CapacityGB - $datastore.FreeSpaceGB
            'UsedSpace (%)'= [math]::Round(($datastore.CapacityGB - $datastore.FreeSpaceGB) / $datastore.CapacityGB * 100, 2)
        }
    }

    # Export Datastore details to CSV
    $datastoreDetails | Export-Excel -Path (Join-Path -Path $outputDir -ChildPath "VUtilizations.xlsx") -WorksheetName "VDatastore" -AutoSize -Append


    Write-Output "VM and ESXi host usage statistics saved to $outputDir\VUtilizations.xlsx"


} catch {
    Write-Error "An error occurred: $_"
} finally {
    # Disconnect from vCenter
    Disconnect-VIServer -Server $vcenterServer -Confirm:$false
}
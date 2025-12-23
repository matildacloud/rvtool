## vCenter Utilization Collection Script

This PowerShell script collects comprehensive utilization statistics from a vCenter server for all active and running virtual machines, ESXi hosts, datacenters, and datastores over the last 30 days.

## Overview

The `vcenter_utilization.ps1` script connects to a vCenter server and collects:
- **Datacenter Information**: Total VMs, clusters, and hosts per datacenter
- **ESXi Host Utilization**: CPU and memory usage (min, max, average), storage capacity and usage
- **Virtual Machine Utilization**: CPU and memory usage (min, max, average), storage usage, power state, and configuration
- **VM Network Information**: IP addresses and network adapter details for VMs with VMware Tools installed
- **Datastore Information**: Capacity, free space, and usage percentages

## Prerequisites

- **PowerShell**: Windows PowerShell 5.1 or PowerShell 7+ (PowerShell Core)
- **Required PowerShell Modules**: The script will automatically install these if missing:
  - `VMware.PowerCLI.VCenter` - VMware PowerCLI module for vCenter operations
  - `ImportExcel` - Module for exporting data to Excel format
- **vCenter Access**: 
  - Network connectivity to your vCenter server
  - Valid vCenter server credentials with appropriate permissions
  - vCenter server address (FQDN or IP address)

## Installation

No manual installation required. The script will automatically check for and install required modules if they are not already present on your system.

### Manual Module Installation (Optional)

If you prefer to install modules manually before running the script:

```powershell
Install-Module -Name VMware.PowerCLI.VCenter -Scope CurrentUser -Force
Install-Module -Name ImportExcel -Scope CurrentUser -Force
```

## Usage

### Basic Execution

1. Open PowerShell (as Administrator recommended for module installation)

2. Navigate to the script directory:
   ```powershell
   cd "c:\rvtool"
   ```

3. Execute the script:
   ```powershell
   .\vcenter_utilization.ps1
   ```

4. When prompted, enter:
   - **vCenter Server Address**: The FQDN or IP address of your vCenter server (e.g., `vcenter.example.com` or `192.168.1.100`)
   - **Credentials**: Your vCenter username and password when the credential dialog appears

### Execution Flow

1. **Module Check**: The script checks for required modules and installs them if missing
2. **Connection**: Connects to the vCenter server using provided credentials
3. **Data Collection**: Processes all datacenters, hosts, VMs, and datastores with progress indicators
4. **Export**: Generates an Excel file with multiple worksheets containing all collected data
5. **Cleanup**: Automatically disconnects from vCenter server

## Output

### Output Location

The script generates an Excel file at:
```
C:\vCenterStats\VUtilizations.xlsx
```

**Note**: If the file already exists, it will be automatically deleted and recreated.

### Output Worksheets

The Excel file contains the following worksheets:

1. **VDatacenter**: 
   - Datacenter name, ID, UUID
   - Total VMs, clusters, and hosts count

2. **VHostUtilization**:
   - Host ID, name, UUID
   - Datacenter association
   - Cluster name
   - CPU usage: Min, Max, Average (%)
   - Memory usage: Min, Max, Average (%)
   - Storage capacity, used space, and usage percentage

3. **VmUtilization**:
   - VM ID, name, UUID
   - Host and datacenter association
   - Power state
   - CPU and memory configuration
   - CPU usage: Min, Max, Average (%)
   - Memory usage: Min, Max, Average (%)
   - Storage used (GB)

4. **VNetwork**:
   - VM name
   - Network adapter name
   - IP address
   - Prefix length (subnet mask)
   - **Note**: Only includes VMs with VMware Tools installed and running

5. **VDatastore**:
   - Datastore name and UUID
   - Capacity (GB)
   - Free space (GB)
   - Used space (GB and percentage)

## Data Collection Period

The script collects utilization statistics for the **last 30 days** from the execution date.

## Network Information Collection

IP addresses are only collected for VMs that meet the following criteria:
- VMware Tools is installed and running (`toolsOk` status)
- The VM has active network adapters with IP configuration

## Error Handling

- The script includes error handling and will display error messages if issues occur
- The vCenter connection is automatically closed in the `finally` block, even if errors occur
- Module installation errors will be displayed but the script will attempt to continue

## Troubleshooting

### Module Installation Issues

If module installation fails:
- Run PowerShell as Administrator
- Check your internet connection (modules are downloaded from PowerShell Gallery)
- Verify execution policy: `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser`

### Connection Issues

- Verify vCenter server address is correct and reachable
- Check network connectivity (ping the vCenter server)
- Verify credentials are correct
- For self-signed certificates, the script automatically ignores certificate validation

### Permission Issues

Ensure your vCenter account has permissions to:
- View datacenters, hosts, and VMs
- Access performance statistics
- Read VM guest information (for network details)

### Output Directory Issues

If the output directory cannot be created:
- Ensure you have write permissions to `C:\`
- Or modify the `$outputDir` variable in the script to use a different location

## Example Output

After successful execution, you should see:
```
VM and ESXi host usage statistics saved to C:\vCenterStats\VUtilizations.xlsx
```

## Notes

- The script processes all datacenters, hosts, and VMs in your vCenter environment
- Execution time depends on the size of your vCenter environment
- Large environments may take significant time to process
- The script uses progress indicators to show current processing status
- All statistics are based on the last 30 days of historical data available in vCenter
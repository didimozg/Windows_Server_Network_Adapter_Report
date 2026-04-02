# Windows Server Network Adapter Report

English edition of the PowerShell script for collecting Windows Server network adapter inventory, default gateways,
saved TCP/IP settings, and legacy network profiles left in the registry.

## What the Script Is For

The script helps you build a single report across many servers when you need to:

- count network adapters on each server
- collect IP, DNS, DHCP, MAC, link speed, and connection status
- review configured gateways from adapter settings and from the routing table
- isolate servers that have multiple default gateways
- export old or already removed adapters whose TCP/IP profiles still remain in the system

## How the Script Works

The script reads a server list from `servers.txt` and connects through WMI/DCOM and the remote registry
(`StdRegProv`). This approach is intentionally used for compatibility with Windows Server 2008-2025 and does not rely
on modern `NetAdapter` or `NetTCPIP` modules being present on the target host.

Processing flow:

1. Read the TXT file with server names.
2. Prompt the operator to use the current account or enter alternate credentials.
3. Connect to each server and collect current adapters, IP configuration, and default gateways.
4. Read TCP/IP profile branches and the class registry to identify old network profiles.
5. Exclude service adapters whose names contain `WAN Miniport` or `Microsoft`.
6. Detect servers with multiple default gateways and write them to a dedicated report.
7. Export the final result to CSV and JSON.

## Key Features

- support for Windows Server 2008, 2008 R2, 2012, 2012 R2, 2016, 2019, 2022, and 2025
- works from both Windows PowerShell 5.1 and PowerShell 7+
- interactive choice between the current user and `Get-Credential`
- parallel server processing through background jobs
- automatic use of Windows PowerShell 5.1 worker processes for WMI collection when the main session runs in
  PowerShell 7+
- dedicated report for servers that have multiple default gateways
- human-readable legacy adapter report with saved addressing details

## Contents of the English Edition

- [get_windows_server_network_adapter_report.ps1](./get_windows_server_network_adapter_report.ps1) - main script
- [servers.txt.example](./servers.txt.example) - sample server list file

## Requirements

- WMI/DCOM access to target servers
- permissions to read WMI and the remote registry
- Windows PowerShell 5.1 available on the machine that launches the script for WMI worker processes
- required RPC/WMI ports open between the launcher and target servers

## Preparing the Server List

Create `servers.txt` next to the script based on `servers.txt.example`.

Example:

```text
SRV-DC-01
SRV-FS-01
# lines starting with # are ignored
SRV-APP-01
```

## Example Usage

```powershell
.\get_windows_server_network_adapter_report.ps1
.\get_windows_server_network_adapter_report.ps1 -UseCurrentUser
.\get_windows_server_network_adapter_report.ps1 -ComputerListPath .\servers.txt
.\get_windows_server_network_adapter_report.ps1 -ComputerListPath .\servers.txt -Credential (Get-Credential)
.\get_windows_server_network_adapter_report.ps1 -Parallel
.\get_windows_server_network_adapter_report.ps1 -Parallel -ThrottleLimit 12
.\get_windows_server_network_adapter_report.ps1 -OutputDirectory .\output\manual_run
```

## Parameters

- `-ComputerListPath` - path to the TXT file with one server name per line
- `-OutputDirectory` - destination folder for exports
- `-Credential` - explicit credentials for WMI/DCOM
- `-UseCurrentUser` - use the current account without prompting for credentials
- `-Parallel` - enable parallel processing
- `-ThrottleLimit` - limit how many jobs run at the same time

## Generated Files

By default, exports are written to `.\output\<yyyyMMdd_HHmmss>`.

- `network_adapter_summary.csv` - short per-server summary
- `network_adapter_details.csv` - current network adapters and their settings
- `network_gateway_details.csv` - gateways found in adapter configuration and in the route table
- `network_multiple_gateways.csv` - servers where multiple default gateways were detected
- `network_legacy_adapters.csv` - human-readable list of legacy adapters and saved settings
- `network_adapter_report.json` - full structured report

## How to Read the Legacy Adapter Report

The main columns in `network_legacy_adapters.csv` are:

- `LegacyCategory` - type of legacy trace found in the system
- `WhatWasFound` - short explanation of what remains
- `AdapterName` - readable adapter or driver name
- `AddressingMode` - saved addressing type
- `DHCPEnabled` - whether DHCP was enabled
- `SavedIPAddress` - saved IP addresses
- `SavedSubnetMask` - saved subnet masks
- `SavedDefaultGateway` - saved gateways
- `SavedDnsServers` - saved DNS servers
- `SavedSettingsSummary` - one-line summary of the saved network settings

## Notes

- the script does not require PowerShell Remoting or WinRM
- blocked ICMP does not prevent data collection as long as WMI/DCOM is available
- older Windows versions do not expose a reliable universal hidden-adapter flag, so the report uses the
  `LegacyWithTcpipProfile`, `RegistryOnly`, and `ClassOnly` categories
- for large server lists, tune `ThrottleLimit` to match network and launcher capacity

## License

Distributed under the MIT License. Root license file: [../LICENSE](../LICENSE)

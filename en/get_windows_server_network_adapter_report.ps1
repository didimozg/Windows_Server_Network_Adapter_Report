<#
.SYNOPSIS
    Collects Windows Server network adapter information from a server list stored in a TXT file.
.DESCRIPTION
    The script reads a list of servers from a text file and connects to them through WMI/DCOM.
    This approach is compatible with Windows Server 2008-2025 and does not depend on modern
    NetAdapter or NetTCPIP modules being present on the target system.

    Parallel processing through background jobs is supported. When the script is started from
    PowerShell 7+, Windows PowerShell 5.1 is automatically used as the worker process for WMI collection.

    The script collects:
    - current network adapters and their properties;
    - IP addresses, DNS, DHCP, and gateways;
    - default gateways from adapter configuration and from the IPv4 route table;
    - legacy or stale network profiles from the registry and class registry.

    When no credential parameters are supplied, the script interactively offers:
    - use the current user;
    - enter alternate credentials.
.PARAMETER ComputerListPath
    Path to the TXT file that contains the server list, one name per line.
.PARAMETER OutputDirectory
    Folder where results will be saved. By default, output\<date_time> is created.
.PARAMETER Credential
    Credentials for remote WMI access. If not supplied, the script prompts the operator to
    choose the current user or enter alternate credentials.
.PARAMETER UseCurrentUser
    Run without prompting for credentials and use the current user.
.PARAMETER Parallel
    Process the server list in parallel through background jobs.
.PARAMETER ThrottleLimit
    Maximum number of simultaneously running jobs in parallel mode.
.EXAMPLE
    .\get_windows_server_network_adapter_report.ps1
.EXAMPLE
    .\get_windows_server_network_adapter_report.ps1 -UseCurrentUser
.EXAMPLE
    .\get_windows_server_network_adapter_report.ps1 -ComputerListPath .\servers.txt -Credential (Get-Credential)
.EXAMPLE
    .\get_windows_server_network_adapter_report.ps1 -Parallel -ThrottleLimit 12
#>
[CmdletBinding()]
param(
    [string]$ComputerListPath = '.\servers.txt',
    [string]$OutputDirectory,
    [System.Management.Automation.PSCredential]$Credential,
    [switch]$UseCurrentUser,
    [switch]$Parallel,
    [ValidateRange(1, 64)]
    [int]$ThrottleLimit = 8,
    [switch]$InternalWorker,
    [string]$InternalComputerName,
    [string]$InternalConnectionMode = 'CurrentUser',
    [string]$CredentialPath
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

if ($Credential -and $UseCurrentUser.IsPresent) {
    throw 'The -Credential and -UseCurrentUser parameters cannot be used together.'
}

if ($InternalWorker.IsPresent -and [string]::IsNullOrWhiteSpace($InternalComputerName)) {
    throw 'The -InternalComputerName parameter must be specified in internal worker mode.'
}

$script:ScriptFilePath = if (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
    $PSCommandPath
}
else {
    $MyInvocation.MyCommand.Path
}
$script:CurrentPowerShellExecutablePath = try {
    (Get-Process -Id $PID -ErrorAction Stop).Path
}
catch {
    $null
}
$script:WindowsPowerShellExecutablePath = Join-Path -Path $env:WINDIR -ChildPath 'System32\WindowsPowerShell\v1.0\powershell.exe'
$script:HasGetWmiObject = $null -ne (Get-Command -Name Get-WmiObject -ErrorAction SilentlyContinue)

if ($InternalWorker.IsPresent -and -not $script:HasGetWmiObject) {
    throw 'Internal worker mode requires Windows PowerShell with the Get-WmiObject command.'
}

$script:RegistryHiveLocalMachine = [uint32]2147483650
$script:NetworkClassRegistryPath = 'SYSTEM\CurrentControlSet\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}'
$script:TcpipInterfacesRegistryPath = 'SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces'

function Resolve-ProjectPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return $Path
    }

    return Join-Path -Path $PSScriptRoot -ChildPath $Path
}

function New-OutputDirectory {
    param(
        [string]$RequestedPath
    )

    $resolvedPath = if ([string]::IsNullOrWhiteSpace($RequestedPath)) {
        Resolve-ProjectPath -Path ("output\{0}" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    }
    else {
        Resolve-ProjectPath -Path $RequestedPath
    }

    New-Item -ItemType Directory -Path $resolvedPath -Force | Out-Null
    return $resolvedPath
}

function Get-ObjectPropertyValue {
    param(
        [object]$InputObject,
        [Parameter(Mandatory = $true)]
        [string]$PropertyName
    )

    if ($null -eq $InputObject) {
        return $null
    }

    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function ConvertTo-FlatArray {
    param(
        [object]$Value
    )

    $items = New-Object System.Collections.Generic.List[string]

    if ($null -eq $Value) {
        return @()
    }

    if ($Value -is [string]) {
        if (-not [string]::IsNullOrWhiteSpace($Value)) {
            $items.Add($Value.Trim())
        }

        return $items.ToArray()
    }

    if ($Value -is [System.Collections.IDictionary]) {
        foreach ($key in $Value.Keys) {
            if ($null -eq $key) {
                continue
            }

            $text = [string]$key
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                $items.Add($text.Trim())
            }
        }

        return $items.ToArray()
    }

    if (($Value -is [System.Collections.IEnumerable]) -and -not ($Value -is [string])) {
        foreach ($entry in $Value) {
            foreach ($nestedItem in (ConvertTo-FlatArray -Value $entry)) {
                if (-not [string]::IsNullOrWhiteSpace($nestedItem)) {
                    $items.Add($nestedItem.Trim())
                }
            }
        }

        return $items.ToArray()
    }

    $textValue = [string]$Value
    if (-not [string]::IsNullOrWhiteSpace($textValue)) {
        $items.Add($textValue.Trim())
    }

    return $items.ToArray()
}

function Join-Items {
    param(
        [object]$Value
    )

    return ((ConvertTo-FlatArray -Value $Value) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) -join '; '
}

function ConvertTo-NormalizedGuid {
    param(
        [object]$GuidValue
    )

    if ($null -eq $GuidValue) {
        return $null
    }

    $text = [string]$GuidValue
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    return $text.Trim().Trim('{', '}').ToUpperInvariant()
}

function ConvertTo-NetworkStatusText {
    param(
        [object]$StatusCode
    )

    if ($null -eq $StatusCode) {
        return $null
    }

    switch ([int]$StatusCode) {
        0 { return 'Disconnected' }
        1 { return 'Connecting' }
        2 { return 'Connected' }
        3 { return 'Disconnecting' }
        4 { return 'HardwareNotPresent' }
        5 { return 'HardwareDisabled' }
        6 { return 'HardwareMalfunction' }
        7 { return 'MediaDisconnected' }
        8 { return 'Authenticating' }
        9 { return 'AuthenticationSucceeded' }
        10 { return 'AuthenticationFailed' }
        11 { return 'InvalidAddress' }
        12 { return 'CredentialsRequired' }
        default { return [string]$StatusCode }
    }
}

function ConvertTo-LinkSpeedText {
    param(
        [object]$SpeedValue
    )

    if ($null -eq $SpeedValue) {
        return $null
    }

    $speedText = [string]$SpeedValue
    if ([string]::IsNullOrWhiteSpace($speedText)) {
        return $null
    }

    $parsedSpeed = 0.0
    if (-not [double]::TryParse($speedText, [ref]$parsedSpeed)) {
        return $speedText
    }

    if ($parsedSpeed -ge 1000000000) {
        return '{0:N2} Gbps' -f ($parsedSpeed / 1000000000)
    }

    if ($parsedSpeed -ge 1000000) {
        return '{0:N0} Mbps' -f ($parsedSpeed / 1000000)
    }

    if ($parsedSpeed -ge 1000) {
        return '{0:N0} Kbps' -f ($parsedSpeed / 1000)
    }

    return '{0:N0} bps' -f $parsedSpeed
}

function Get-AddressFamily {
    param(
        [string]$Address
    )

    if ([string]::IsNullOrWhiteSpace($Address)) {
        return $null
    }

    if ($Address.Contains(':')) {
        return 'IPv6'
    }

    return 'IPv4'
}

function Get-ErrorMessageText {
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $messages = New-Object System.Collections.Generic.List[string]
    $currentException = $ErrorRecord.Exception

    while ($null -ne $currentException) {
        if (-not [string]::IsNullOrWhiteSpace($currentException.Message)) {
            $messages.Add($currentException.Message.Trim())
        }

        $currentException = $currentException.InnerException
    }

    if ($messages.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($ErrorRecord.ToString())) {
        $messages.Add($ErrorRecord.ToString().Trim())
    }

    return ($messages | Select-Object -Unique) -join ' | '
}

function Test-IsExcludedAdapterInfo {
    param(
        [string[]]$Values
    )

    foreach ($value in $Values) {
        if ([string]::IsNullOrWhiteSpace($value)) {
            continue
        }

        $normalizedValue = $value.ToLowerInvariant()
        if ($normalizedValue -like '*wan miniport*' -or $normalizedValue -like '*microsoft*') {
            return $true
        }
    }

    return $false
}

function Get-FirstNonEmptyString {
    param(
        [string[]]$Values
    )

    foreach ($value in $Values) {
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value.Trim()
        }
    }

    return $null
}

function ConvertTo-YesNoText {
    param(
        [object]$Value
    )

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return 'Unknown'
    }

    switch ($text.Trim().ToLowerInvariant()) {
        'true' { return 'Yes' }
        'false' { return 'No' }
        '1' { return 'Yes' }
        '0' { return 'No' }
        default { return $text }
    }
}

function Get-LegacyDetectionLabel {
    param(
        [string]$DetectionType
    )

    switch ($DetectionType) {
        'LegacyWithTcpipProfile' { return 'Legacy adapter with saved TCP/IP profile' }
        'RegistryOnly' { return 'Only the TCP/IP profile remains in the registry' }
        'ClassOnly' { return 'Only the driver or adapter registry entry remains' }
        default { return $DetectionType }
    }
}

function Get-LegacyDetectionDescription {
    param(
        [string]$DetectionType
    )

    switch ($DetectionType) {
        'LegacyWithTcpipProfile' { return 'The adapter is no longer active, but both the driver entry and network settings remain in the system.' }
        'RegistryOnly' { return 'The adapter is no longer present, but TCP/IP settings are still saved under the Interfaces registry branch.' }
        'ClassOnly' { return 'The adapter entry remains in the class registry, but the TCP/IP profile is no longer present.' }
        default { return $null }
    }
}

function Get-LegacyAddressingMode {
    param(
        [string]$StaticIPAddress,
        [string]$DhcpIPAddress,
        [object]$EnableDHCP
    )

    if (-not [string]::IsNullOrWhiteSpace($StaticIPAddress)) {
        return 'Static configuration'
    }

    if (-not [string]::IsNullOrWhiteSpace($DhcpIPAddress) -or [string]$EnableDHCP -eq '1') {
        return 'DHCP'
    }

    return 'No saved network settings'
}

function Get-CleanReadableNetworkValue {
    param(
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $cleanValues = @(
        $Value -split ';' |
            ForEach-Object { $_.Trim() } |
            Where-Object {
                -not [string]::IsNullOrWhiteSpace($_) -and
                $_ -notin @('0.0.0.0', '::', '255.255.255.255')
            }
    )

    if ($cleanValues.Count -eq 0) {
        return $null
    }

    return ($cleanValues | Select-Object -Unique) -join '; '
}

function Get-ServerListFromFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ListPath
    )

    $resolvedListPath = Resolve-ProjectPath -Path $ListPath

    if (-not (Test-Path -LiteralPath $resolvedListPath)) {
        $examplePath = Resolve-ProjectPath -Path '.\servers.txt.example'
        if (Test-Path -LiteralPath $examplePath) {
            throw "Server list file not found: $resolvedListPath. Use this template: $examplePath"
        }

        throw "Server list file not found: $resolvedListPath"
    }

    $servers = New-Object System.Collections.Generic.List[string]
    foreach ($line in (Get-Content -LiteralPath $resolvedListPath -ErrorAction Stop)) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        $trimmedLine = $line.Trim()
        if ($trimmedLine.StartsWith('#')) {
            continue
        }

        $servers.Add($trimmedLine)
    }

    if ($servers.Count -eq 0) {
        throw "The file '$resolvedListPath' does not contain any servers to process."
    }

    return $servers.ToArray() | Sort-Object -Unique
}

function Resolve-AuthenticationContext {
    param(
        [System.Management.Automation.PSCredential]$ProvidedCredential,
        [switch]$RunAsCurrentUser
    )

    if ($ProvidedCredential) {
        return [PSCustomObject]@{
            Credential = $ProvidedCredential
            Mode = 'Credential'
            DisplayName = 'Provided credentials'
        }
    }

    if ($RunAsCurrentUser.IsPresent) {
        return [PSCustomObject]@{
            Credential = $null
            Mode = 'CurrentUser'
            DisplayName = 'Current user'
        }
    }

    $canPrompt = $true
    try {
        $null = $Host.UI.RawUI
    }
    catch {
        $canPrompt = $false
    }

    if (-not $canPrompt) {
        return [PSCustomObject]@{
            Credential = $null
            Mode = 'CurrentUser'
            DisplayName = 'Current user'
        }
    }

    Write-Host 'Choose how to connect to the servers:' -ForegroundColor Yellow
    Write-Host '1. Use the current user' -ForegroundColor DarkYellow
    Write-Host '2. Enter alternate credentials' -ForegroundColor DarkYellow

    while ($true) {
        $choice = Read-Host 'Enter 1 or 2'
        switch ($choice) {
            '1' {
                return [PSCustomObject]@{
                    Credential = $null
                    Mode = 'CurrentUser'
                    DisplayName = 'Current user'
                }
            }
            '2' {
                return [PSCustomObject]@{
                    Credential = Get-Credential -Message 'Enter credentials for remote WMI access'
                    Mode = 'Credential'
                    DisplayName = 'Prompted credentials'
                }
            }
            default {
                Write-Warning 'You must choose 1 or 2.'
            }
        }
    }
}

function Get-WorkerPowerShellExecutablePath {
    $currentExecutableName = if (-not [string]::IsNullOrWhiteSpace($script:CurrentPowerShellExecutablePath)) {
        [System.IO.Path]::GetFileName($script:CurrentPowerShellExecutablePath).ToLowerInvariant()
    }
    else {
        $null
    }

    if (
        $script:HasGetWmiObject -and
        -not [string]::IsNullOrWhiteSpace($script:CurrentPowerShellExecutablePath) -and
        $currentExecutableName -in @('powershell.exe', 'pwsh.exe')
    ) {
        return $script:CurrentPowerShellExecutablePath
    }

    if (Test-Path -LiteralPath $script:WindowsPowerShellExecutablePath) {
        return $script:WindowsPowerShellExecutablePath
    }

    throw 'Could not determine a PowerShell engine for WMI collection. Windows PowerShell 5.1 is required.'
}

function Get-WorkerInvocationArguments {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetComputerName,
        [Parameter(Mandatory = $true)]
        [string]$ConnectionMode,
        [string]$CredentialFilePath,
        [Parameter(Mandatory = $true)]
        [string]$PowerShellExecutablePath
    )

    $invokeArguments = @('-NoProfile')
    if ([System.IO.Path]::GetFileName($PowerShellExecutablePath).ToLowerInvariant() -eq 'powershell.exe') {
        $invokeArguments += @('-ExecutionPolicy', 'Bypass')
    }

    $invokeArguments += @(
        '-File', $script:ScriptFilePath,
        '-InternalWorker',
        '-InternalComputerName', $TargetComputerName,
        '-InternalConnectionMode', $ConnectionMode
    )

    if (-not [string]::IsNullOrWhiteSpace($CredentialFilePath)) {
        $invokeArguments += @('-CredentialPath', $CredentialFilePath)
    }

    return $invokeArguments
}

function New-TemporaryCredentialFile {
    param(
        [System.Management.Automation.PSCredential]$Credential
    )

    if ($null -eq $Credential) {
        return $null
    }

    $credentialFilePath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ('network-adapter-report-credential-{0}.clixml' -f ([guid]::NewGuid().ToString('N')))
    $Credential | Export-Clixml -Path $credentialFilePath
    return $credentialFilePath
}

function Convert-WorkerExecutionToResult {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetComputerName,
        [Parameter(Mandatory = $true)]
        [string[]]$OutputLines,
        [Parameter(Mandatory = $true)]
        [int]$ExitCode,
        [Parameter(Mandatory = $true)]
        [string]$ConnectionMode
    )

    $jsonText = ($OutputLines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join [Environment]::NewLine
    if ($ExitCode -eq 0 -and -not [string]::IsNullOrWhiteSpace($jsonText)) {
        try {
            return $jsonText | ConvertFrom-Json -Depth 12
        }
        catch {
            $parseErrorMessage = 'Failed to parse worker process output: {0}' -f $_.Exception.Message
            return New-ErrorResult -ComputerName $TargetComputerName -ConnectionMode $ConnectionMode -ErrorMessage $parseErrorMessage
        }
    }

    $errorText = if (-not [string]::IsNullOrWhiteSpace($jsonText)) {
        $jsonText
    }
    else {
        'The worker process finished without output. Exit code: {0}' -f $ExitCode
    }

    return New-ErrorResult -ComputerName $TargetComputerName -ConnectionMode $ConnectionMode -ErrorMessage $errorText
}

function Invoke-InventoryWorkerProcess {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetComputerName,
        [Parameter(Mandatory = $true)]
        [string]$ConnectionMode,
        [string]$CredentialFilePath,
        [Parameter(Mandatory = $true)]
        [string]$PowerShellExecutablePath
    )

    $invokeArguments = Get-WorkerInvocationArguments -TargetComputerName $TargetComputerName -ConnectionMode $ConnectionMode -CredentialFilePath $CredentialFilePath -PowerShellExecutablePath $PowerShellExecutablePath
    $outputLines = @(& $PowerShellExecutablePath @invokeArguments 2>&1 | ForEach-Object { [string]$_ })
    $exitCode = if ($null -ne $LASTEXITCODE) { [int]$LASTEXITCODE } else { 0 }

    return Convert-WorkerExecutionToResult -TargetComputerName $TargetComputerName -OutputLines $outputLines -ExitCode $exitCode -ConnectionMode $ConnectionMode
}

function Receive-CompletedInventoryJobResults {
    param(
        [Parameter(Mandatory = $true)]
        [object]$JobList,
        [Parameter(Mandatory = $true)]
        [object]$ResultList,
        [Parameter(Mandatory = $true)]
        [string]$ConnectionMode
    )

    $completedJobs = @($JobList | Where-Object { $_.State -notin @('NotStarted', 'Running') })
    foreach ($job in $completedJobs) {
        $targetComputerName = if ($job.PSObject.Properties['TargetComputerName']) {
            [string]$job.PSObject.Properties['TargetComputerName'].Value
        }
        else {
            [string]$job.Name
        }

        $jobOutput = @()
        try {
            $jobOutput = @(Receive-Job -Job $job -ErrorAction SilentlyContinue)
        }
        catch {
            $jobOutput = @()
        }

        $structuredOutput = $null
        if ($jobOutput.Count -gt 0) {
            $structuredOutput = $jobOutput | Select-Object -Last 1
        }

        if ($null -ne $structuredOutput -and $structuredOutput.PSObject.Properties['TargetComputerName']) {
            $ResultList.Add(
                (Convert-WorkerExecutionToResult -TargetComputerName ([string]$structuredOutput.TargetComputerName) -OutputLines @($structuredOutput.OutputLines) -ExitCode ([int]$structuredOutput.ExitCode) -ConnectionMode $ConnectionMode)
            )
        }
        else {
            $fallbackErrorText = if ($job.ChildJobs.Count -gt 0 -and $job.ChildJobs[0].Error.Count -gt 0) {
                (($job.ChildJobs[0].Error | ForEach-Object { $_.ToString() }) -join ' | ')
            }
            else {
                'The background job completed without returning data.'
            }

            $ResultList.Add(
                (New-ErrorResult -ComputerName $targetComputerName -ConnectionMode $ConnectionMode -ErrorMessage $fallbackErrorText)
            )
        }

        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        [void]$JobList.Remove($job)
    }
}

function Invoke-NetworkInventoryCollection {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$ServerList,
        [Parameter(Mandatory = $true)]
        [pscustomobject]$AuthenticationContext,
        [switch]$UseParallel,
        [ValidateRange(1, 64)]
        [int]$MaxParallelJobs = 8
    )

    $workerPowerShellExecutablePath = Get-WorkerPowerShellExecutablePath
    $credentialFilePath = New-TemporaryCredentialFile -Credential $AuthenticationContext.Credential

    try {
        if (-not $UseParallel.IsPresent) {
            $sequentialResults = New-Object System.Collections.Generic.List[object]
            foreach ($server in $ServerList) {
                Write-Host "Processing: $server" -ForegroundColor Cyan
                $sequentialResults.Add(
                    (Invoke-InventoryWorkerProcess -TargetComputerName $server -ConnectionMode $AuthenticationContext.Mode -CredentialFilePath $credentialFilePath -PowerShellExecutablePath $workerPowerShellExecutablePath)
                )
            }

            return $sequentialResults.ToArray()
        }

        if (-not (Get-Command -Name Start-Job -ErrorAction SilentlyContinue)) {
            Write-Warning 'The Start-Job command is not available. Falling back to sequential mode.'
            return Invoke-NetworkInventoryCollection -ServerList $ServerList -AuthenticationContext $AuthenticationContext -MaxParallelJobs $MaxParallelJobs
        }

        if ($PSVersionTable.PSVersion.Major -lt 5) {
            Write-Warning 'PowerShell 5.1 or later is recommended for stable parallel processing. Falling back to sequential mode.'
            return Invoke-NetworkInventoryCollection -ServerList $ServerList -AuthenticationContext $AuthenticationContext -MaxParallelJobs $MaxParallelJobs
        }

        Write-Host ("Parallel mode is enabled. Up to {0} servers will be processed at the same time." -f $MaxParallelJobs) -ForegroundColor DarkGreen

        $activeJobs = New-Object System.Collections.Generic.List[object]
        $parallelResults = New-Object System.Collections.Generic.List[object]
        foreach ($server in $ServerList) {
            while ($activeJobs.Count -ge $MaxParallelJobs) {
                $null = Wait-Job -Job $activeJobs.ToArray() -Any
                Receive-CompletedInventoryJobResults -JobList $activeJobs -ResultList $parallelResults -ConnectionMode $AuthenticationContext.Mode
            }

            Write-Host "Starting job for: $server" -ForegroundColor Cyan
            $jobArguments = Get-WorkerInvocationArguments -TargetComputerName $server -ConnectionMode $AuthenticationContext.Mode -CredentialFilePath $credentialFilePath -PowerShellExecutablePath $workerPowerShellExecutablePath
            $job = Start-Job -Name ('NetworkAdapter-{0}' -f $server) -ArgumentList @($workerPowerShellExecutablePath, $jobArguments, $server) -ScriptBlock {
                param($powerShellExecutablePath, $invokeArguments, $targetComputerName)

                $outputLines = @(& $powerShellExecutablePath @invokeArguments 2>&1 | ForEach-Object { [string]$_ })
                $exitCode = if ($null -ne $LASTEXITCODE) { [int]$LASTEXITCODE } else { 0 }

                [PSCustomObject]@{
                    TargetComputerName = $targetComputerName
                    ExitCode = $exitCode
                    OutputLines = @($outputLines)
                }
            }

            $job | Add-Member -NotePropertyName TargetComputerName -NotePropertyValue $server -Force
            $activeJobs.Add($job)
        }

        while ($activeJobs.Count -gt 0) {
            $null = Wait-Job -Job $activeJobs.ToArray() -Any
            Receive-CompletedInventoryJobResults -JobList $activeJobs -ResultList $parallelResults -ConnectionMode $AuthenticationContext.Mode
        }

        return $parallelResults.ToArray()
    }
    finally {
        if (-not [string]::IsNullOrWhiteSpace($credentialFilePath) -and (Test-Path -LiteralPath $credentialFilePath)) {
            Remove-Item -LiteralPath $credentialFilePath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-WmiObjectCompat {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Class,
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential,
        [string]$Namespace = 'root\cimv2',
        [string]$Filter
    )

    $wmiParams = @{
        Class = $Class
        ComputerName = $ComputerName
        Namespace = $Namespace
        ErrorAction = 'Stop'
    }

    if (-not [string]::IsNullOrWhiteSpace($Filter)) {
        $wmiParams.Filter = $Filter
    }

    if ($Credential) {
        $wmiParams.Credential = $Credential
    }

    return Get-WmiObject @wmiParams
}

function Get-RegistryProvider {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential
    )

    $wmiParams = @{
        List = $true
        Class = 'StdRegProv'
        ComputerName = $ComputerName
        Namespace = 'root\default'
        ErrorAction = 'Stop'
    }

    if ($Credential) {
        $wmiParams.Credential = $Credential
    }

    return Get-WmiObject @wmiParams
}

function Invoke-RegistryMethodCompat {
    param(
        [Parameter(Mandatory = $true)]
        [object]$RegistryProvider,
        [Parameter(Mandatory = $true)]
        [string]$MethodName,
        [object[]]$ArgumentList
    )

    try {
        return Invoke-WmiMethod -InputObject $RegistryProvider -Name $MethodName -ArgumentList $ArgumentList -ErrorAction Stop
    }
    catch {
        return $null
    }
}

function Get-RegistrySubKeys {
    param(
        [Parameter(Mandatory = $true)]
        [object]$RegistryProvider,
        [Parameter(Mandatory = $true)]
        [string]$SubKeyPath
    )

    $result = Invoke-RegistryMethodCompat -RegistryProvider $RegistryProvider -MethodName 'EnumKey' -ArgumentList @($script:RegistryHiveLocalMachine, $SubKeyPath)
    if ($null -eq $result -or $result.ReturnValue -ne 0 -or $null -eq $result.sNames) {
        return @()
    }

    return @($result.sNames)
}

function Get-RegistryValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$RegistryProvider,
        [Parameter(Mandatory = $true)]
        [string]$SubKeyPath,
        [Parameter(Mandatory = $true)]
        [string]$ValueName,
        [ValidateSet('String', 'MultiString', 'DWord')]
        [string]$ValueType = 'String'
    )

    switch ($ValueType) {
        'String' {
            $lookupPlan = @(
                @{ Method = 'GetStringValue'; Property = 'sValue' }
                @{ Method = 'GetExpandedStringValue'; Property = 'sValue' }
            )
        }
        'MultiString' {
            $lookupPlan = @(
                @{ Method = 'GetMultiStringValue'; Property = 'sValue' }
                @{ Method = 'GetStringValue'; Property = 'sValue' }
                @{ Method = 'GetExpandedStringValue'; Property = 'sValue' }
            )
        }
        'DWord' {
            $lookupPlan = @(
                @{ Method = 'GetDWORDValue'; Property = 'uValue' }
            )
        }
    }

    foreach ($lookup in $lookupPlan) {
        $result = Invoke-RegistryMethodCompat -RegistryProvider $RegistryProvider -MethodName $lookup.Method -ArgumentList @($script:RegistryHiveLocalMachine, $SubKeyPath, $ValueName)
        if ($null -eq $result -or $result.ReturnValue -ne 0) {
            continue
        }

        $value = Get-ObjectPropertyValue -InputObject $result -PropertyName $lookup.Property
        if ($null -ne $value) {
            return $value
        }
    }

    return $null
}

function Get-NetworkClassRegistryMap {
    param(
        [Parameter(Mandatory = $true)]
        [object]$RegistryProvider
    )

    $map = @{}
    foreach ($subKey in (Get-RegistrySubKeys -RegistryProvider $RegistryProvider -SubKeyPath $script:NetworkClassRegistryPath)) {
        $fullPath = '{0}\{1}' -f $script:NetworkClassRegistryPath, $subKey
        $interfaceGuid = ConvertTo-NormalizedGuid -GuidValue (Get-RegistryValue -RegistryProvider $RegistryProvider -SubKeyPath $fullPath -ValueName 'NetCfgInstanceId' -ValueType 'String')
        if ([string]::IsNullOrWhiteSpace($interfaceGuid)) {
            continue
        }

        $map[$interfaceGuid] = [PSCustomObject]@{
            RegistryClassKey = $subKey
            InterfaceGuid = $interfaceGuid
            DriverDescription = Join-Items -Value (Get-RegistryValue -RegistryProvider $RegistryProvider -SubKeyPath $fullPath -ValueName 'DriverDesc' -ValueType 'String')
            ProviderName = Join-Items -Value (Get-RegistryValue -RegistryProvider $RegistryProvider -SubKeyPath $fullPath -ValueName 'ProviderName' -ValueType 'String')
            ComponentId = Join-Items -Value (Get-RegistryValue -RegistryProvider $RegistryProvider -SubKeyPath $fullPath -ValueName 'ComponentId' -ValueType 'String')
            Manufacturer = Join-Items -Value (Get-RegistryValue -RegistryProvider $RegistryProvider -SubKeyPath $fullPath -ValueName 'Mfg' -ValueType 'String')
            PnpInstanceId = Join-Items -Value (Get-RegistryValue -RegistryProvider $RegistryProvider -SubKeyPath $fullPath -ValueName 'MatchingDeviceId' -ValueType 'String')
        }
    }

    return $map
}

function Get-TcpipInterfaceProfiles {
    param(
        [Parameter(Mandatory = $true)]
        [object]$RegistryProvider
    )

    $profiles = @{}
    foreach ($subKey in (Get-RegistrySubKeys -RegistryProvider $RegistryProvider -SubKeyPath $script:TcpipInterfacesRegistryPath)) {
        $fullPath = '{0}\{1}' -f $script:TcpipInterfacesRegistryPath, $subKey
        $interfaceGuid = ConvertTo-NormalizedGuid -GuidValue $subKey
        if ([string]::IsNullOrWhiteSpace($interfaceGuid)) {
            continue
        }

        $profiles[$interfaceGuid] = [PSCustomObject]@{
            RegistryKey = $subKey
            InterfaceGuid = $interfaceGuid
            EnableDHCP = Join-Items -Value (Get-RegistryValue -RegistryProvider $RegistryProvider -SubKeyPath $fullPath -ValueName 'EnableDHCP' -ValueType 'DWord')
            StaticIPAddress = Join-Items -Value (Get-RegistryValue -RegistryProvider $RegistryProvider -SubKeyPath $fullPath -ValueName 'IPAddress' -ValueType 'MultiString')
            StaticSubnetMask = Join-Items -Value (Get-RegistryValue -RegistryProvider $RegistryProvider -SubKeyPath $fullPath -ValueName 'SubnetMask' -ValueType 'MultiString')
            StaticDefaultGateway = Join-Items -Value (Get-RegistryValue -RegistryProvider $RegistryProvider -SubKeyPath $fullPath -ValueName 'DefaultGateway' -ValueType 'MultiString')
            StaticGatewayMetric = Join-Items -Value (Get-RegistryValue -RegistryProvider $RegistryProvider -SubKeyPath $fullPath -ValueName 'DefaultGatewayMetric' -ValueType 'MultiString')
            DhcpIPAddress = Join-Items -Value (Get-RegistryValue -RegistryProvider $RegistryProvider -SubKeyPath $fullPath -ValueName 'DhcpIPAddress' -ValueType 'String')
            DhcpSubnetMask = Join-Items -Value (Get-RegistryValue -RegistryProvider $RegistryProvider -SubKeyPath $fullPath -ValueName 'DhcpSubnetMask' -ValueType 'String')
            DhcpDefaultGateway = Join-Items -Value (Get-RegistryValue -RegistryProvider $RegistryProvider -SubKeyPath $fullPath -ValueName 'DhcpDefaultGateway' -ValueType 'MultiString')
            DhcpGatewayMetric = Join-Items -Value (Get-RegistryValue -RegistryProvider $RegistryProvider -SubKeyPath $fullPath -ValueName 'DhcpDefaultGatewayMetric' -ValueType 'MultiString')
            NameServer = Join-Items -Value (Get-RegistryValue -RegistryProvider $RegistryProvider -SubKeyPath $fullPath -ValueName 'NameServer' -ValueType 'String')
            DhcpNameServer = Join-Items -Value (Get-RegistryValue -RegistryProvider $RegistryProvider -SubKeyPath $fullPath -ValueName 'DhcpNameServer' -ValueType 'String')
            Domain = Join-Items -Value (Get-RegistryValue -RegistryProvider $RegistryProvider -SubKeyPath $fullPath -ValueName 'Domain' -ValueType 'String')
            DhcpDomain = Join-Items -Value (Get-RegistryValue -RegistryProvider $RegistryProvider -SubKeyPath $fullPath -ValueName 'DhcpDomain' -ValueType 'String')
            InterfaceMetric = Join-Items -Value (Get-RegistryValue -RegistryProvider $RegistryProvider -SubKeyPath $fullPath -ValueName 'InterfaceMetric' -ValueType 'DWord')
        }
    }

    return $profiles
}

function New-ErrorResult {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        [Parameter(Mandatory = $true)]
        [string]$ConnectionMode,
        [Parameter(Mandatory = $true)]
        [string]$ErrorMessage
    )

    return [PSCustomObject]@{
        Summary = [PSCustomObject]@{
            ComputerName = $ComputerName
            CollectionStatus = 'Error'
            CollectedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            Domain = $null
            Manufacturer = $null
            Model = $null
            OperatingSystem = $null
            ConnectionMode = $ConnectionMode
            TotalAdapters = 0
            PhysicalAdapters = 0
            ConnectedAdapters = 0
            IpEnabledAdapters = 0
            ConfiguredGatewayCount = 0
            UniqueConfiguredGatewayCount = 0
            RouteTableDefaultGatewayCount = 0
            HasMultipleDefaultGateways = 'No'
            LegacyAdapterCount = 0
            LegacyRegistryOnlyCount = 0
            LegacyClassOnlyCount = 0
            ErrorMessage = $ErrorMessage
        }
        Adapters = @()
        Gateways = @()
        LegacyAdapters = @()
        MultipleGatewayAlerts = @()
    }
}

function Get-NetworkInventoryForServer {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory = $true)]
        [string]$ConnectionMode
    )

    try {
        $computerSystem = Get-WmiObjectCompat -Class 'Win32_ComputerSystem' -ComputerName $ComputerName -Credential $Credential
        $operatingSystem = Get-WmiObjectCompat -Class 'Win32_OperatingSystem' -ComputerName $ComputerName -Credential $Credential
        $networkAdapters = @(Get-WmiObjectCompat -Class 'Win32_NetworkAdapter' -ComputerName $ComputerName -Credential $Credential | Sort-Object -Property Index)
        $adapterConfigurations = @(Get-WmiObjectCompat -Class 'Win32_NetworkAdapterConfiguration' -ComputerName $ComputerName -Credential $Credential | Sort-Object -Property Index)
        $registryProvider = Get-RegistryProvider -ComputerName $ComputerName -Credential $Credential

        $defaultRoutes = @()
        try {
            $defaultRoutes = @(Get-WmiObjectCompat -Class 'Win32_IP4RouteTable' -ComputerName $ComputerName -Credential $Credential -Filter "Destination='0.0.0.0' AND Mask='0.0.0.0'" | Sort-Object -Property InterfaceIndex, Metric1, NextHop)
        }
        catch {
            $defaultRoutes = @()
        }

        $configByIndex = @{}
        $configByGuid = @{}
        foreach ($config in $adapterConfigurations) {
            $configIndexKey = [string](Get-ObjectPropertyValue -InputObject $config -PropertyName 'Index')
            if (-not [string]::IsNullOrWhiteSpace($configIndexKey)) {
                $configByIndex[$configIndexKey] = $config
            }

            $settingId = ConvertTo-NormalizedGuid -GuidValue (Get-ObjectPropertyValue -InputObject $config -PropertyName 'SettingID')
            if (-not [string]::IsNullOrWhiteSpace($settingId)) {
                $configByGuid[$settingId] = $config
            }
        }

        $adapterRows = @(foreach ($adapter in $networkAdapters) {
            $index = Get-ObjectPropertyValue -InputObject $adapter -PropertyName 'Index'
            $indexKey = [string]$index
            $interfaceGuid = ConvertTo-NormalizedGuid -GuidValue (Get-ObjectPropertyValue -InputObject $adapter -PropertyName 'GUID')
            $config = $null

            if ($configByIndex.ContainsKey($indexKey)) {
                $config = $configByIndex[$indexKey]
            }
            elseif (-not [string]::IsNullOrWhiteSpace($interfaceGuid) -and $configByGuid.ContainsKey($interfaceGuid)) {
                $config = $configByGuid[$interfaceGuid]
            }

            $netConnectionStatusCode = Get-ObjectPropertyValue -InputObject $adapter -PropertyName 'NetConnectionStatus'
            $speed = Get-ObjectPropertyValue -InputObject $adapter -PropertyName 'Speed'

            [PSCustomObject]@{
                ComputerName = $ComputerName
                OperatingSystem = [string](Get-ObjectPropertyValue -InputObject $operatingSystem -PropertyName 'Caption')
                Name = [string](Get-ObjectPropertyValue -InputObject $adapter -PropertyName 'Name')
                NetConnectionID = [string](Get-ObjectPropertyValue -InputObject $adapter -PropertyName 'NetConnectionID')
                Description = [string](Get-ObjectPropertyValue -InputObject $adapter -PropertyName 'Description')
                ProductName = [string](Get-ObjectPropertyValue -InputObject $adapter -PropertyName 'ProductName')
                InterfaceGuid = $interfaceGuid
                Index = $index
                InterfaceIndex = Get-ObjectPropertyValue -InputObject $adapter -PropertyName 'InterfaceIndex'
                AdapterType = [string](Get-ObjectPropertyValue -InputObject $adapter -PropertyName 'AdapterType')
                Manufacturer = [string](Get-ObjectPropertyValue -InputObject $adapter -PropertyName 'Manufacturer')
                MACAddress = [string](Get-ObjectPropertyValue -InputObject $adapter -PropertyName 'MACAddress')
                PhysicalAdapter = [string](Get-ObjectPropertyValue -InputObject $adapter -PropertyName 'PhysicalAdapter')
                NetEnabled = [string](Get-ObjectPropertyValue -InputObject $adapter -PropertyName 'NetEnabled')
                NetConnectionStatusCode = $netConnectionStatusCode
                NetConnectionStatus = ConvertTo-NetworkStatusText -StatusCode $netConnectionStatusCode
                SpeedBitsPerSecond = $speed
                LinkSpeed = ConvertTo-LinkSpeedText -SpeedValue $speed
                PNPDeviceID = [string](Get-ObjectPropertyValue -InputObject $adapter -PropertyName 'PNPDeviceID')
                ServiceName = [string](Get-ObjectPropertyValue -InputObject $adapter -PropertyName 'ServiceName')
                DHCPEnabled = [string](Get-ObjectPropertyValue -InputObject $config -PropertyName 'DHCPEnabled')
                DHCPServer = [string](Get-ObjectPropertyValue -InputObject $config -PropertyName 'DHCPServer')
                IPEnabled = [string](Get-ObjectPropertyValue -InputObject $config -PropertyName 'IPEnabled')
                IPAddress = Join-Items -Value (Get-ObjectPropertyValue -InputObject $config -PropertyName 'IPAddress')
                IPSubnet = Join-Items -Value (Get-ObjectPropertyValue -InputObject $config -PropertyName 'IPSubnet')
                DefaultGateway = Join-Items -Value (Get-ObjectPropertyValue -InputObject $config -PropertyName 'DefaultIPGateway')
                GatewayMetric = Join-Items -Value (Get-ObjectPropertyValue -InputObject $config -PropertyName 'GatewayCostMetric')
                DNSServers = Join-Items -Value (Get-ObjectPropertyValue -InputObject $config -PropertyName 'DNSServerSearchOrder')
                DNSDomain = [string](Get-ObjectPropertyValue -InputObject $config -PropertyName 'DNSDomain')
                DNSHostName = [string](Get-ObjectPropertyValue -InputObject $config -PropertyName 'DNSHostName')
                IPConnectionMetric = Join-Items -Value (Get-ObjectPropertyValue -InputObject $config -PropertyName 'IPConnectionMetric')
                WINSPrimaryServer = [string](Get-ObjectPropertyValue -InputObject $config -PropertyName 'WINSPrimaryServer')
                WINSSecondaryServer = [string](Get-ObjectPropertyValue -InputObject $config -PropertyName 'WINSSecondaryServer')
                TcpipNetbiosOptions = [string](Get-ObjectPropertyValue -InputObject $config -PropertyName 'TcpipNetbiosOptions')
            }
        })

        $adapterRows = @($adapterRows | Where-Object {
            -not [string]::IsNullOrWhiteSpace($_.InterfaceGuid) -or
            -not [string]::IsNullOrWhiteSpace($_.MACAddress) -or
            -not [string]::IsNullOrWhiteSpace($_.PNPDeviceID) -or
            -not [string]::IsNullOrWhiteSpace($_.NetConnectionID)
        })

        $excludedInterfaceGuidSet = @{}
        $excludedInterfaceIndexSet = @{}
        foreach ($excludedAdapterRow in @($adapterRows | Where-Object {
            Test-IsExcludedAdapterInfo -Values @(
                $_.Name,
                $_.NetConnectionID,
                $_.Description,
                $_.ProductName,
                $_.Manufacturer
            )
        })) {
            if (-not [string]::IsNullOrWhiteSpace($excludedAdapterRow.InterfaceGuid)) {
                $excludedInterfaceGuidSet[$excludedAdapterRow.InterfaceGuid] = $true
            }

            $excludedInterfaceIndexKey = [string]$excludedAdapterRow.InterfaceIndex
            if (-not [string]::IsNullOrWhiteSpace($excludedInterfaceIndexKey)) {
                $excludedInterfaceIndexSet[$excludedInterfaceIndexKey] = $true
            }
        }

        $adapterRows = @($adapterRows | Where-Object {
            -not (Test-IsExcludedAdapterInfo -Values @(
                $_.Name,
                $_.NetConnectionID,
                $_.Description,
                $_.ProductName,
                $_.Manufacturer
            ))
        })

        $adapterByGuid = @{}
        $adapterByInterfaceIndex = @{}
        foreach ($adapterRow in $adapterRows) {
            if (-not [string]::IsNullOrWhiteSpace($adapterRow.InterfaceGuid)) {
                $adapterByGuid[$adapterRow.InterfaceGuid] = $adapterRow
            }

            $interfaceIndexKey = [string]$adapterRow.InterfaceIndex
            if (-not [string]::IsNullOrWhiteSpace($interfaceIndexKey) -and -not $adapterByInterfaceIndex.ContainsKey($interfaceIndexKey)) {
                $adapterByInterfaceIndex[$interfaceIndexKey] = $adapterRow
            }
        }

        $gatewayRows = New-Object System.Collections.Generic.List[object]

        foreach ($config in $adapterConfigurations) {
            $settingId = ConvertTo-NormalizedGuid -GuidValue (Get-ObjectPropertyValue -InputObject $config -PropertyName 'SettingID')
            if (-not [string]::IsNullOrWhiteSpace($settingId) -and $excludedInterfaceGuidSet.ContainsKey($settingId)) {
                continue
            }

            $matchedAdapter = if (-not [string]::IsNullOrWhiteSpace($settingId) -and $adapterByGuid.ContainsKey($settingId)) { $adapterByGuid[$settingId] } else { $null }
            $gatewayValues = @(ConvertTo-FlatArray -Value (Get-ObjectPropertyValue -InputObject $config -PropertyName 'DefaultIPGateway'))
            $gatewayMetrics = @(ConvertTo-FlatArray -Value (Get-ObjectPropertyValue -InputObject $config -PropertyName 'GatewayCostMetric'))

            for ($i = 0; $i -lt $gatewayValues.Count; $i++) {
                $metricValue = $null
                if ($gatewayMetrics.Count -gt $i) {
                    $metricValue = $gatewayMetrics[$i]
                }
                elseif ($gatewayMetrics.Count -gt 0) {
                    $metricValue = $gatewayMetrics -join '; '
                }

                $gatewayRows.Add([PSCustomObject]@{
                    ComputerName = $ComputerName
                    Source = 'AdapterConfiguration'
                    AddressFamily = Get-AddressFamily -Address $gatewayValues[$i]
                    Gateway = $gatewayValues[$i]
                    Metric = $metricValue
                    InterfaceGuid = $settingId
                    InterfaceIndex = Get-ObjectPropertyValue -InputObject $config -PropertyName 'InterfaceIndex'
                    AdapterName = if ($matchedAdapter) { $matchedAdapter.Name } else { $null }
                    NetConnectionID = if ($matchedAdapter) { $matchedAdapter.NetConnectionID } else { $null }
                    AdapterDescription = if ($matchedAdapter) { $matchedAdapter.Description } else { [string](Get-ObjectPropertyValue -InputObject $config -PropertyName 'Description') }
                    RawDestination = $null
                    RawMask = $null
                })
            }
        }

        foreach ($route in $defaultRoutes) {
            $routeInterfaceIndex = Get-ObjectPropertyValue -InputObject $route -PropertyName 'InterfaceIndex'
            $matchedAdapter = $null
            $routeInterfaceKey = [string]$routeInterfaceIndex
            if (-not [string]::IsNullOrWhiteSpace($routeInterfaceKey) -and $excludedInterfaceIndexSet.ContainsKey($routeInterfaceKey)) {
                continue
            }

            if (-not [string]::IsNullOrWhiteSpace($routeInterfaceKey) -and $adapterByInterfaceIndex.ContainsKey($routeInterfaceKey)) {
                $matchedAdapter = $adapterByInterfaceIndex[$routeInterfaceKey]
            }

            $gatewayRows.Add([PSCustomObject]@{
                ComputerName = $ComputerName
                Source = 'RouteTableIPv4'
                AddressFamily = 'IPv4'
                Gateway = [string](Get-ObjectPropertyValue -InputObject $route -PropertyName 'NextHop')
                Metric = [string](Get-ObjectPropertyValue -InputObject $route -PropertyName 'Metric1')
                InterfaceGuid = if ($matchedAdapter) { $matchedAdapter.InterfaceGuid } else { $null }
                InterfaceIndex = $routeInterfaceIndex
                AdapterName = if ($matchedAdapter) { $matchedAdapter.Name } else { $null }
                NetConnectionID = if ($matchedAdapter) { $matchedAdapter.NetConnectionID } else { $null }
                AdapterDescription = if ($matchedAdapter) { $matchedAdapter.Description } else { $null }
                RawDestination = [string](Get-ObjectPropertyValue -InputObject $route -PropertyName 'Destination')
                RawMask = [string](Get-ObjectPropertyValue -InputObject $route -PropertyName 'Mask')
            })
        }

        $networkClassMap = Get-NetworkClassRegistryMap -RegistryProvider $registryProvider
        $tcpipProfiles = Get-TcpipInterfaceProfiles -RegistryProvider $registryProvider

        $legacyRows = New-Object System.Collections.Generic.List[object]
        $processedLegacyGuids = @{}

        foreach ($profileGuid in ($tcpipProfiles.Keys | Sort-Object)) {
            if ($adapterByGuid.ContainsKey($profileGuid)) {
                continue
            }

            if ($excludedInterfaceGuidSet.ContainsKey($profileGuid)) {
                continue
            }

            $legacyProfile = $tcpipProfiles[$profileGuid]
            $classInfo = if ($networkClassMap.ContainsKey($profileGuid)) { $networkClassMap[$profileGuid] } else { $null }
            if ($classInfo -and (Test-IsExcludedAdapterInfo -Values @(
                $classInfo.DriverDescription,
                $classInfo.ProviderName,
                $classInfo.ComponentId,
                $classInfo.Manufacturer
            ))) {
                continue
            }

            $detectionType = if ($classInfo) { 'LegacyWithTcpipProfile' } else { 'RegistryOnly' }

            $legacyRows.Add([PSCustomObject]@{
                ComputerName = $ComputerName
                DetectionType = $detectionType
                InterfaceGuid = $legacyProfile.InterfaceGuid
                RegistryKey = $legacyProfile.RegistryKey
                RegistryClassKey = if ($classInfo) { $classInfo.RegistryClassKey } else { $null }
                AdapterName = if ($classInfo) { $classInfo.DriverDescription } else { $null }
                ProviderName = if ($classInfo) { $classInfo.ProviderName } else { $null }
                ComponentId = if ($classInfo) { $classInfo.ComponentId } else { $null }
                Manufacturer = if ($classInfo) { $classInfo.Manufacturer } else { $null }
                PnpInstanceId = if ($classInfo) { $classInfo.PnpInstanceId } else { $null }
                EnableDHCP = $legacyProfile.EnableDHCP
                StaticIPAddress = $legacyProfile.StaticIPAddress
                StaticSubnetMask = $legacyProfile.StaticSubnetMask
                StaticDefaultGateway = $legacyProfile.StaticDefaultGateway
                StaticGatewayMetric = $legacyProfile.StaticGatewayMetric
                DhcpIPAddress = $legacyProfile.DhcpIPAddress
                DhcpSubnetMask = $legacyProfile.DhcpSubnetMask
                DhcpDefaultGateway = $legacyProfile.DhcpDefaultGateway
                DhcpGatewayMetric = $legacyProfile.DhcpGatewayMetric
                NameServer = $legacyProfile.NameServer
                DhcpNameServer = $legacyProfile.DhcpNameServer
                Domain = $legacyProfile.Domain
                DhcpDomain = $legacyProfile.DhcpDomain
                InterfaceMetric = $legacyProfile.InterfaceMetric
            })

            $processedLegacyGuids[$profileGuid] = $true
        }

        foreach ($classGuid in ($networkClassMap.Keys | Sort-Object)) {
            if ($adapterByGuid.ContainsKey($classGuid) -or $processedLegacyGuids.ContainsKey($classGuid)) {
                continue
            }

            if ($excludedInterfaceGuidSet.ContainsKey($classGuid)) {
                continue
            }

            $classInfo = $networkClassMap[$classGuid]
            if (Test-IsExcludedAdapterInfo -Values @(
                $classInfo.DriverDescription,
                $classInfo.ProviderName,
                $classInfo.ComponentId,
                $classInfo.Manufacturer
            )) {
                continue
            }

            $legacyRows.Add([PSCustomObject]@{
                ComputerName = $ComputerName
                DetectionType = 'ClassOnly'
                InterfaceGuid = $classInfo.InterfaceGuid
                RegistryKey = $null
                RegistryClassKey = $classInfo.RegistryClassKey
                AdapterName = $classInfo.DriverDescription
                ProviderName = $classInfo.ProviderName
                ComponentId = $classInfo.ComponentId
                Manufacturer = $classInfo.Manufacturer
                PnpInstanceId = $classInfo.PnpInstanceId
                EnableDHCP = $null
                StaticIPAddress = $null
                StaticSubnetMask = $null
                StaticDefaultGateway = $null
                StaticGatewayMetric = $null
                DhcpIPAddress = $null
                DhcpSubnetMask = $null
                DhcpDefaultGateway = $null
                DhcpGatewayMetric = $null
                NameServer = $null
                DhcpNameServer = $null
                Domain = $null
                DhcpDomain = $null
                InterfaceMetric = $null
            })
        }

        $gatewayRows = @($gatewayRows.ToArray())
        $legacyRows = @($legacyRows.ToArray())

        $configuredGatewayRows = @($gatewayRows | Where-Object {
            $_.Source -eq 'AdapterConfiguration' -and
            -not [string]::IsNullOrWhiteSpace($_.Gateway)
        })
        $routeTableGatewayRows = @($gatewayRows | Where-Object {
            $_.Source -eq 'RouteTableIPv4' -and
            -not [string]::IsNullOrWhiteSpace($_.Gateway)
        })

        $configuredGatewayAssignments = @(
            $configuredGatewayRows |
                ForEach-Object {
                    '{0}|{1}|{2}|{3}' -f $_.Gateway, $_.InterfaceGuid, $_.InterfaceIndex, $_.NetConnectionID
                } |
                Sort-Object -Unique
        )
        $uniqueConfiguredGatewayValues = @(
            $configuredGatewayRows |
                Select-Object -ExpandProperty Gateway |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Sort-Object -Unique
        )
        $uniqueRouteGatewayValues = @(
            $routeTableGatewayRows |
                Select-Object -ExpandProperty Gateway |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Sort-Object -Unique
        )

        $multipleGatewayAlerts = @()
        if ($configuredGatewayAssignments.Count -gt 1 -or $uniqueRouteGatewayValues.Count -gt 1) {
            $configuredGatewayByAdapter = @(
                $configuredGatewayRows |
                    ForEach-Object {
                        if (-not [string]::IsNullOrWhiteSpace($_.NetConnectionID)) {
                            '{0}: {1}' -f $_.NetConnectionID, $_.Gateway
                        }
                        elseif (-not [string]::IsNullOrWhiteSpace($_.AdapterName)) {
                            '{0}: {1}' -f $_.AdapterName, $_.Gateway
                        }
                        else {
                            'InterfaceIndex {0}: {1}' -f $_.InterfaceIndex, $_.Gateway
                        }
                    } |
                    Sort-Object -Unique
            )

            $issueText = if ($configuredGatewayAssignments.Count -gt 1) {
                'Multiple default gateways were found in the network adapter configuration.'
            }
            else {
                'Multiple default routes were found in the route table.'
            }

            $multipleGatewayAlerts = @([PSCustomObject]@{
                ComputerName = $ComputerName
                OperatingSystem = [string](Get-ObjectPropertyValue -InputObject $operatingSystem -PropertyName 'Caption')
                ConfiguredGatewayEntryCount = $configuredGatewayAssignments.Count
                UniqueConfiguredGatewayCount = $uniqueConfiguredGatewayValues.Count
                ConfiguredGatewayList = Join-Items -Value $uniqueConfiguredGatewayValues
                ConfiguredGatewayByAdapter = Join-Items -Value $configuredGatewayByAdapter
                RouteTableGatewayCount = $uniqueRouteGatewayValues.Count
                RouteTableGatewayList = Join-Items -Value $uniqueRouteGatewayValues
                Issue = $issueText
                Recommendation = 'Verify whether this server is expected to use multiple default gateways.'
            })
        }

        $legacyRows = @($legacyRows | ForEach-Object {
            $savedIpAddress = Get-CleanReadableNetworkValue -Value (Get-FirstNonEmptyString -Values @($_.StaticIPAddress, $_.DhcpIPAddress))
            $savedSubnetMask = Get-CleanReadableNetworkValue -Value (Get-FirstNonEmptyString -Values @($_.StaticSubnetMask, $_.DhcpSubnetMask))
            $savedGateway = Get-CleanReadableNetworkValue -Value (Get-FirstNonEmptyString -Values @($_.StaticDefaultGateway, $_.DhcpDefaultGateway))
            $savedDnsServers = Get-CleanReadableNetworkValue -Value (Get-FirstNonEmptyString -Values @($_.NameServer, $_.DhcpNameServer))
            $savedDomain = Get-CleanReadableNetworkValue -Value (Get-FirstNonEmptyString -Values @($_.Domain, $_.DhcpDomain))
            $adapterDisplayName = Get-FirstNonEmptyString -Values @($_.AdapterName, $_.Manufacturer, 'Adapter could not be determined')
            $registryLocations = @()
            if (-not [string]::IsNullOrWhiteSpace($_.RegistryKey)) {
                $registryLocations += 'Tcpip\Interfaces\{0}' -f $_.RegistryKey
            }
            if (-not [string]::IsNullOrWhiteSpace($_.RegistryClassKey)) {
                $registryLocations += 'Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}\{0}' -f $_.RegistryClassKey
            }

            $savedSettingsSummary = Join-Items -Value @(
                if ($savedIpAddress) { 'IP: {0}' -f $savedIpAddress }
                if ($savedSubnetMask) { 'Mask: {0}' -f $savedSubnetMask }
                if ($savedGateway) { 'Gateway: {0}' -f $savedGateway }
                if ($savedDnsServers) { 'DNS: {0}' -f $savedDnsServers }
                if ($savedDomain) { 'DNS suffix: {0}' -f $savedDomain }
            )

            [PSCustomObject]@{
                ComputerName = $_.ComputerName
                LegacyCategory = Get-LegacyDetectionLabel -DetectionType $_.DetectionType
                WhatWasFound = Get-LegacyDetectionDescription -DetectionType $_.DetectionType
                AdapterName = $adapterDisplayName
                AddressingMode = Get-LegacyAddressingMode -StaticIPAddress $_.StaticIPAddress -DhcpIPAddress $_.DhcpIPAddress -EnableDHCP $_.EnableDHCP
                DHCPEnabled = ConvertTo-YesNoText -Value $_.EnableDHCP
                SavedIPAddress = $savedIpAddress
                SavedSubnetMask = $savedSubnetMask
                SavedDefaultGateway = $savedGateway
                SavedDnsServers = $savedDnsServers
                SavedDnsSuffix = $savedDomain
                SavedSettingsSummary = $savedSettingsSummary
                Manufacturer = $_.Manufacturer
                ProviderName = $_.ProviderName
                PnpInstanceId = $_.PnpInstanceId
                RegistryLocations = Join-Items -Value $registryLocations
                InterfaceGuid = $_.InterfaceGuid
                DetectionType = $_.DetectionType
            }
        })

        $summary = [PSCustomObject]@{
            ComputerName = $ComputerName
            CollectionStatus = 'Success'
            CollectedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            Domain = [string](Get-ObjectPropertyValue -InputObject $computerSystem -PropertyName 'Domain')
            Manufacturer = [string](Get-ObjectPropertyValue -InputObject $computerSystem -PropertyName 'Manufacturer')
            Model = [string](Get-ObjectPropertyValue -InputObject $computerSystem -PropertyName 'Model')
            OperatingSystem = [string](Get-ObjectPropertyValue -InputObject $operatingSystem -PropertyName 'Caption')
            ConnectionMode = $ConnectionMode
            TotalAdapters = @($adapterRows).Count
            PhysicalAdapters = @($adapterRows | Where-Object { $_.PhysicalAdapter -eq 'True' }).Count
            ConnectedAdapters = @($adapterRows | Where-Object { $_.NetConnectionStatus -eq 'Connected' }).Count
            IpEnabledAdapters = @($adapterRows | Where-Object { $_.IPEnabled -eq 'True' }).Count
            ConfiguredGatewayCount = @($gatewayRows | Where-Object { $_.Source -eq 'AdapterConfiguration' }).Count
            UniqueConfiguredGatewayCount = $uniqueConfiguredGatewayValues.Count
            RouteTableDefaultGatewayCount = @($gatewayRows | Where-Object { $_.Source -eq 'RouteTableIPv4' }).Count
            HasMultipleDefaultGateways = if ($multipleGatewayAlerts.Count -gt 0) { 'Yes' } else { 'No' }
            LegacyAdapterCount = @($legacyRows).Count
            LegacyRegistryOnlyCount = @($legacyRows | Where-Object { $_.DetectionType -eq 'RegistryOnly' }).Count
            LegacyClassOnlyCount = @($legacyRows | Where-Object { $_.DetectionType -eq 'ClassOnly' }).Count
            ErrorMessage = $null
        }

        return [PSCustomObject]@{
            Summary = $summary
            Adapters = @($adapterRows)
            Gateways = @($gatewayRows)
            LegacyAdapters = @($legacyRows)
            MultipleGatewayAlerts = @($multipleGatewayAlerts)
        }
    }
    catch {
        $detailedErrorMessage = Get-ErrorMessageText -ErrorRecord $_
        if ($_.InvocationInfo -and -not [string]::IsNullOrWhiteSpace($_.InvocationInfo.PositionMessage)) {
            $detailedErrorMessage = '{0} | {1}' -f $detailedErrorMessage, ($_.InvocationInfo.PositionMessage.Trim() -replace '\s+', ' ')
        }

        return New-ErrorResult -ComputerName $ComputerName -ConnectionMode $ConnectionMode -ErrorMessage $detailedErrorMessage
    }
}

function Export-NetworkInventoryReport {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Results,
        [Parameter(Mandatory = $true)]
        [string]$ReportOutputDirectory
    )

    $summaryRows = @($Results | ForEach-Object { $_.Summary })
    $adapterRows = @($Results | ForEach-Object { $_.Adapters })
    $gatewayRows = @($Results | ForEach-Object { $_.Gateways })
    $legacyRows = @($Results | ForEach-Object { $_.LegacyAdapters })
    $multipleGatewayRows = @($Results | ForEach-Object { $_.MultipleGatewayAlerts })

    $summaryPath = Join-Path -Path $ReportOutputDirectory -ChildPath 'network_adapter_summary.csv'
    $adapterPath = Join-Path -Path $ReportOutputDirectory -ChildPath 'network_adapter_details.csv'
    $gatewayPath = Join-Path -Path $ReportOutputDirectory -ChildPath 'network_gateway_details.csv'
    $legacyPath = Join-Path -Path $ReportOutputDirectory -ChildPath 'network_legacy_adapters.csv'
    $multipleGatewayPath = Join-Path -Path $ReportOutputDirectory -ChildPath 'network_multiple_gateways.csv'
    $jsonPath = Join-Path -Path $ReportOutputDirectory -ChildPath 'network_adapter_report.json'

    $summaryRows |
        Sort-Object -Property ComputerName |
        Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8

    $adapterRows |
        Sort-Object -Property ComputerName, Index, Name |
        Export-Csv -Path $adapterPath -NoTypeInformation -Encoding UTF8

    $gatewayRows |
        Sort-Object -Property ComputerName, Source, InterfaceIndex, Gateway |
        Export-Csv -Path $gatewayPath -NoTypeInformation -Encoding UTF8

    $legacyRows |
        Sort-Object -Property ComputerName, LegacyCategory, AdapterName |
        Export-Csv -Path $legacyPath -NoTypeInformation -Encoding UTF8

    if ($multipleGatewayRows.Count -gt 0) {
        $multipleGatewayRows |
            Sort-Object -Property ComputerName |
            Export-Csv -Path $multipleGatewayPath -NoTypeInformation -Encoding UTF8
    }
    else {
        '"ComputerName","OperatingSystem","ConfiguredGatewayEntryCount","UniqueConfiguredGatewayCount","ConfiguredGatewayList","ConfiguredGatewayByAdapter","RouteTableGatewayCount","RouteTableGatewayList","Issue","Recommendation"' |
            Set-Content -Path $multipleGatewayPath -Encoding UTF8
    }

    $Results | ConvertTo-Json -Depth 8 | Set-Content -Path $jsonPath -Encoding UTF8

    return [PSCustomObject]@{
        SummaryRows = @($summaryRows)
        MultipleGatewayRows = @($multipleGatewayRows)
        SummaryPath = $summaryPath
        AdapterPath = $adapterPath
        GatewayPath = $gatewayPath
        MultipleGatewayPath = $multipleGatewayPath
        LegacyPath = $legacyPath
        JsonPath = $jsonPath
    }
}

function Invoke-NetworkInventoryReport {
    if ($InternalWorker.IsPresent) {
        $workerCredential = $null
        if (-not [string]::IsNullOrWhiteSpace($CredentialPath)) {
            if (-not (Test-Path -LiteralPath $CredentialPath)) {
                $workerErrorResult = New-ErrorResult -ComputerName $InternalComputerName -ConnectionMode $InternalConnectionMode -ErrorMessage ("Credential file not found: {0}" -f $CredentialPath)
                $workerErrorResult | ConvertTo-Json -Depth 8 -Compress
                return
            }

            $workerCredential = Import-Clixml -Path $CredentialPath
        }

        $workerResult = Get-NetworkInventoryForServer -ComputerName $InternalComputerName -Credential $workerCredential -ConnectionMode $InternalConnectionMode
        $workerResult | ConvertTo-Json -Depth 8 -Compress
        return
    }

    $authenticationContext = Resolve-AuthenticationContext -ProvidedCredential $Credential -RunAsCurrentUser:$UseCurrentUser
    $serverList = Get-ServerListFromFile -ListPath $ComputerListPath
    $resolvedOutputDirectory = New-OutputDirectory -RequestedPath $OutputDirectory

    Write-Host 'Collecting network adapter information...' -ForegroundColor Green
    Write-Host "Connection mode: $($authenticationContext.DisplayName)" -ForegroundColor DarkGreen
    Write-Host "Server list file: $(Resolve-ProjectPath -Path $ComputerListPath)" -ForegroundColor DarkGreen
    Write-Host "Output folder: $resolvedOutputDirectory" -ForegroundColor DarkGreen

    if ($Parallel.IsPresent) {
        Write-Host ("Parallel mode requested. PowerShell version: {0}" -f $PSVersionTable.PSVersion.ToString()) -ForegroundColor DarkGreen
    }

    $results = @(Invoke-NetworkInventoryCollection -ServerList $serverList -AuthenticationContext $authenticationContext -UseParallel:$Parallel -MaxParallelJobs $ThrottleLimit)
    $exportResult = Export-NetworkInventoryReport -Results $results -ReportOutputDirectory $resolvedOutputDirectory

    $successCount = @($exportResult.SummaryRows | Where-Object { $_.CollectionStatus -eq 'Success' }).Count
    $errorCount = @($exportResult.SummaryRows | Where-Object { $_.CollectionStatus -eq 'Error' }).Count
    $multipleGatewayServerCount = @($exportResult.MultipleGatewayRows).Count

    Write-Host "Done. Servers processed successfully: $successCount" -ForegroundColor Green
    if ($errorCount -gt 0) {
        Write-Warning "Servers processed with errors: $errorCount. See network_adapter_summary.csv for details."
    }
    if ($multipleGatewayServerCount -gt 0) {
        Write-Warning "Servers with multiple default gateways found: $multipleGatewayServerCount. See network_multiple_gateways.csv for details."
    }

    Write-Host "Summary report: $($exportResult.SummaryPath)" -ForegroundColor DarkGreen
    Write-Host "Adapter details: $($exportResult.AdapterPath)" -ForegroundColor DarkGreen
    Write-Host "Gateways: $($exportResult.GatewayPath)" -ForegroundColor DarkGreen
    Write-Host "Servers with multiple gateways: $($exportResult.MultipleGatewayPath)" -ForegroundColor DarkGreen
    Write-Host "Legacy adapters: $($exportResult.LegacyPath)" -ForegroundColor DarkGreen
    Write-Host "Full JSON: $($exportResult.JsonPath)" -ForegroundColor DarkGreen
}

Invoke-NetworkInventoryReport












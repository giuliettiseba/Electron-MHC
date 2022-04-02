
Import-Module "$PSScriptRoot\bin\MilestonePSTools.dll"
enum VmsTaskState {
    Completed
    Error
    Idle
    InProgress
    Success
    Unknown
}

class VmsTaskResult {
    [int] $Progress
    [string] $Path
    [string] $ErrorCode
    [string] $ErrorText
    [VmsTaskState] $State

    VmsTaskResult () {}

    VmsTaskResult([VideoOS.ConfigurationApi.ClientService.ConfigurationItem] $InvokeItem) {
        foreach ($p in $InvokeItem.Properties) {
            switch ($p.ValueType) {
                'Progress' {
                    $this.($p.Key) = [int]$p.Value
                }
                'Tick' {
                    $this.($p.Key) = [bool]::Parse($p.Value)
                }
                default {
                    $this.($p.Key) = $p.Value
                }
            }
        }
    }
}

class VmsHardwareScanResult : VmsTaskResult {
    [uri] $HardwareAddress
    [string] $UserName
    [string] $Password
    [bool] $MacAddressExistsGlobal
    [bool] $MacAddressExistsLocal
    [bool] $HardwareScanValidated
    [string] $MacAddress
    [string] $HardwareDriverPath

    # Property hidden so that this type can be cleanly exported to CSV or something
    # without adding a column with a complex object in it.
    hidden [VideoOS.Platform.ConfigurationItems.RecordingServer] $RecordingServer

    VmsHardwareScanResult() {}

    VmsHardwareScanResult([VideoOS.ConfigurationApi.ClientService.ConfigurationItem] $InvokeItem) {
        foreach ($p in $InvokeItem.Properties) {
            switch ($p.ValueType) {
                'Progress' {
                    $this.($p.Key) = [int]$p.Value
                }
                'Tick' {
                    $this.($p.Key) = [bool]::Parse($p.Value)
                }
                default {
                    $this.($p.Key) = $p.Value
                }
            }
        }
    }
}

# Contains the output from the script passed to LocalJobRunner.AddJob, in addition to any errors thrown in the script if present.
class LocalJobResult {
    [object[]] $Output
    [System.Management.Automation.ErrorRecord[]] $Errors
}

# Contains the IAsyncResult object returned by PowerShell.BeginInvoke() as well as the PowerShell instance we need to
class LocalJob {
    [System.Management.Automation.PowerShell] $PowerShell
    [System.IAsyncResult] $Result
}

# Centralizes the complexity of running multiple commands/scripts at a time and receiving the results, including errors, when they complete.
class LocalJobRunner : IDisposable {
    hidden [System.Management.Automation.Runspaces.RunspacePool] $RunspacePool
    hidden [System.Collections.Generic.List[LocalJob]] $Jobs
    [timespan] $JobPollingInterval = (New-Timespan -Seconds 1)

    # Default constructor creates an underlying runspace pool with a max size matching the number of processors
    LocalJobRunner () {
        $this.Initialize($env:NUMBER_OF_PROCESSORS)
    }

    # Optionally you may manually specify a max size for the underlying runspace pool.
    LocalJobRunner ([int]$MaxSize) {
        $this.Initialize($MaxSize)
    }

    hidden [void] Initialize([int]$MaxSize) {
        $this.Jobs = New-Object System.Collections.Generic.List[LocalJob]
        $this.RunspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxSize)
        $this.RunspacePool.Open()
    }

    # Accepts a scriptblock and a set of parameters. A new powewershell instance will be created, attached to a runspacepool, and the results can be collected later in a call to ReceiveJobs.
    [LocalJob] AddJob([scriptblock]$scriptblock, [hashtable]$parameters) {
        $parameters = if ($null -eq $parameters) { $parameters = @{} } else { $parameters }
        $shell = [powershell]::Create()
        $shell.RunspacePool = $this.RunspacePool
        $asyncResult = $shell.AddScript($scriptblock).AddParameters($parameters).BeginInvoke()
        $job = [LocalJob]@{
            PowerShell = $shell
            Result     = $asyncResult
        }
        $this.Jobs.Add($job)
        return $job
    }

    # Returns the output from specific jobs
    [LocalJobResult[]] ReceiveJobs([LocalJob[]]$localJobs) {
        $completedJobs = $localJobs | Where-Object { $_.Result.IsCompleted }
        $completedJobs | Foreach-Object { $this.Jobs.Remove($_) }
        $results = $completedJobs | Foreach-Object {
            [LocalJobResult]@{
                Output = $_.PowerShell.EndInvoke($_.Result)
                Errors = $_.PowerShell.Streams.Error
            }

            $_.PowerShell.Dispose()
        }
        return $results
    }

    # Returns the output from any completed jobs in an object that also includes any errors if present.
    [LocalJobResult[]] ReceiveJobs() {
        return $this.ReceiveJobs($this.Jobs)
    }

    # Block until all jobs have completed. The list of jobs will be polled on an interval of JobPollingInterval, which is 1 second by default.
    [void] Wait() {
        $this.Wait($this.Jobs)
    }

    # Block until all jobs have completed. The list of jobs will be polled on an interval of JobPollingInterval, which is 1 second by default.
    [void] Wait([LocalJob[]]$jobList) {
        while ($jobList.Result.IsCompleted -contains $false) {
            Start-Sleep -Seconds $this.JobPollingInterval.TotalSeconds
        }
    }

    # Returns $true if there are any jobs available to be received using ReceiveJobs. Use to implement your own polling strategy instead of using Wait.
    [bool] HasPendingJobs() {
        return ($this.Jobs.Count -gt 0)
    }

    # Make sure to dispose of this class so that the underlying runspace pool gets disposed.
    [void] Dispose() {
        $this.Jobs.Clear()
        $this.RunspacePool.Close()
        $this.RunspacePool.Dispose()
    }
}

class VmsCameraStreamConfig {
    [string] $Name
    [string] $DisplayName
    [bool] $Enabled
    [bool] $LiveDefault
    [string] $LiveMode
    [bool] $Recorded
    [hashtable] $Settings
    [hashtable] $ValueTypeInfo
    hidden [VideoOS.Platform.ConfigurationItems.Camera] $Camera
}

class VmsStreamDeviceStatus : VideoOS.Platform.SDK.Proxy.Status2.MediaStreamDeviceStatusBase {
    [string] $DeviceName
    [string] $DeviceType
    [string] $RecorderName
    [guid]   $RecorderId
    [bool]   $Motion

    VmsStreamDeviceStatus () {}
    VmsStreamDeviceStatus ([VideoOS.Platform.SDK.Proxy.Status2.MediaStreamDeviceStatusBase]$status) {
        $this.DbMoveInProgress = $status.DbMoveInProgress
        $this.DbRepairInProgress = $status.DbRepairInProgress
        if ($null -ne $status.DeviceId) {
            $this.DeviceId = $status.DeviceId
        }
        $this.Enabled = $status.Enabled
        $this.Error = $status.Error
        $this.ErrorNoConnection = $status.ErrorNoConnection
        $this.ErrorNotLicensed = $status.ErrorNotLicensed
        $this.ErrorOverflow = $status.ErrorOverflow
        $this.ErrorWritingGop = $status.ErrorWritingGop
        $this.IsChange = $status.IsChange
        $this.Recording = $status.Recording
        $this.Started = $status.Started
        if ($null -ne $status.Time) {
            $this.Time = $status.Time
        }
        if ($null -ne $status.Motion) {
            $this.Motion = $status.Motion
        }
    }
}
function Assert-VmsConnected {
    [CmdletBinding()]
    param (
    )

    process {
        if ($null -eq [MilestonePSTools.Connection.MilestoneConnection]::Instance) {
            $message = 'Not connected to a Management Server.'
            if ($script:Messages) {
                $message = $script:Messages.NotConnectedToAManagementServer
            }
            throw ([VideoOS.Platform.CommunicationMIPException]::new($message))
        }
    }
}
function Assert-VmsVersion {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [Version]
        $MinimumVersion
    )

    process {
        $site = Get-Site
        $currentVersion = [version]$site.Properties['ServerVersion']
        if ($currentVersion -lt $MinimumVersion) {
            $callingFunction = (Get-PSCallStack)[1].Command
            $message = "$callingFunction requires a minimum server version of $MinimumVersion. The current site is running version $currentVersion."
            $exception = [notsupportedexception]::new($message)

            $errorParams = @{
                Message = $message
                ErrorId = 'MinVmsVersionNotMet'
                Exception = $exception
                Category = 'NotImplemented'
                RecommendedAction = "Upgrade server to Milestone VMS version $MinimumVersion or later."
            }
            Write-Error @errorParams
        }
    }
}
class VmsConfigChildItemSettings {
    [string]    $Name
    [hashtable] $Properties
    [hashtable] $ValueTypeInfo
}

function ConvertFrom-ConfigChildItem {
    [CmdletBinding()]
    [OutputType([VmsConfigChildItemSettings])]
    param(
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [VideoOS.Platform.ConfigurationItems.IConfigurationChildItem]
        $InputObject,

        [Parameter()]
        [switch]
        $RawValues
    )

    process {
        # When we look up display values for raw values, sometimes
        # the raw value matches the value of a valuetypeinfo property
        # like MinValue or MaxValue. We don't want to display "MinValue"
        # as the display value for a setting, so this list of valuetypeinfo
        # entry names should be ignored.
        $ignoredNames = 'MinValue', 'MaxValue', 'StepValue'
        $properties = @{}
        $valueTypeInfos = @{}
        foreach ($key in $InputObject.Properties.Keys) {
            # Sometimes the Keys are the same as KeyFullName and other times
            # they are short, easy to read names. So just in case, we'll test
            # the key by splitting it and seeing how many parts there are. A
            # KeysFullName value looks like 'device:0.0/RecorderMode/75f374ab-8dd2-4fd0-b8f5-155fa730702c'
            $keyParts = $key -split '/', 3
            $keyName = if ($keyParts.Count -gt 1) { $keyParts[1] } else { $key }

            $value = $InputObject.Properties.GetValue($key)
            $valueTypeInfo = $InputObject.Properties.GetValueTypeInfoCollection($key)

            if (-not $RawValues) {
                <#
                  Unless -RawValues was used, we'll check to see if there's a
                  display name available for the value for the current setting.
                  If a ValueTypeInfo entry has a Value matching the raw value,
                  and the Name of that value isn't one of the internal names we
                  want to ignore, we'll replace $value with the ValueTypeInfo
                  Name. Here's a reference ValueTypeInfo table for RecorderMode:

                  TranslationId                        Name       Value
                  -------------                        ----       -----
                  b9f5c797-ebbf-55ad-ccdd-8539a65a0241 Disabled   0
                  535863a8-2f16-3709-557e-59e2eb8139a7 Continuous 1
                  8226588f-03da-49b8-57e5-ddf8c508dd2d Motion     2

                  So if the raw value of RecorderMode is 0, we would return
                  "Disabled" unless the -RawValues switch is used.
                #>

                $friendlyValue = ($valueTypeInfo | Select-Object | Where-Object {
                        $_.Value -eq $value -and $_.Name -notin $ignoredNames
                    }).Name
                if (-not [string]::IsNullOrWhiteSpace($friendlyValue)) {
                    $value = $friendlyValue
                }
            }

            $properties[$keyName] = $value
            $valueTypeInfos[$keyName] = $valueTypeInfo
        }

        [VmsConfigChildItemSettings]@{
            Name          = $InputObject.DisplayName
            Properties    = $properties
            ValueTypeInfo = $valueTypeInfos
        }
    }
}
function ConvertFrom-StreamUsage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.StreamUsageChildItem]
        $StreamUsage
    )

    process {
        $streamName = $StreamUsage.StreamReferenceIdValues.Keys | Where-Object {
            $StreamUsage.StreamReferenceIdValues.$_ -eq $StreamUsage.StreamReferenceId
        }
        Write-Output $streamName
    }
}
function ConvertTo-Uri {
    <#
    .SYNOPSIS
    Accepts an IPv4 or IPv6 address and converts it to an http or https URI

    .DESCRIPTION
    Accepts an IPv4 or IPv6 address and converts it to an http or https URI. IPv6 addresses need to
    be wrapped in square brackets when used in a URI. This function is used to help normalize data
    into an expected URI format.

    .PARAMETER IPAddress
    Specifies an IPAddress object of either Internetwork or InternetworkV6.

    .PARAMETER UseHttps
    Specifies whether the resulting URI should use https as the scheme instead of http.

    .PARAMETER HttpPort
    Specifies an alternate port to override the default http/https ports.

    .EXAMPLE
    '192.168.1.1' | ConvertTo-Uri
    #>
    [CmdletBinding()]
    [OutputType([uri])]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [IPAddress]
        $IPAddress,

        [Parameter()]
        [switch]
        $UseHttps,

        [Parameter()]
        [int]
        $HttpPort = 80
    )

    process {
        $builder = [uribuilder]::new()
        $builder.Scheme = if ($UseHttps) { 'https' } else { 'http' }
        $builder.Host = if ($IPAddress.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetworkV6) {
            "[$IPAddress]"
        }
        else {
            $IPAddress
        }
        $builder.Port = $HttpPort
        Write-Output $builder.Uri
    }
}
function Copy-ConfigurationItem {
    [CmdletBinding()]
    param (
        [parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [pscustomobject]
        $InputObject,
        [parameter(Mandatory, Position = 1)]
        [VideoOS.ConfigurationApi.ClientService.ConfigurationItem]
        $DestinationItem
    )

    process {
        if (!$DestinationItem.ChildrenFilled) {
            Write-Verbose "$($DestinationItem.DisplayName) has not been retrieved recursively. Retrieving child items now."
            $DestinationItem = $DestinationItem | Get-ConfigurationItem -Recurse -Sort
        }

        $srcStack = New-Object -TypeName System.Collections.Stack
        $srcStack.Push($InputObject)
        $dstStack = New-Object -TypeName System.Collections.Stack
        $dstStack.Push($DestinationItem)

        Write-Verbose "Configuring $($DestinationItem.DisplayName) ($($DestinationItem.Path))"
        while ($dstStack.Count -gt 0) {
            $dirty = $false
            $src = $srcStack.Pop()
            $dst = $dstStack.Pop()

            if (($src.ItemCategory -ne $dst.ItemCategory) -or ($src.ItemType -ne $dst.ItemType)) {
                Write-Error "Source and Destination ConfigurationItems are different"
                return
            }

            if ($src.EnableProperty.Enabled -ne $dst.EnableProperty.Enabled) {
                Write-Verbose "$(if ($src.EnableProperty.Enabled) { "Enabling"} else { "Disabling" }) $($dst.DisplayName)"
                $dst.EnableProperty.Enabled = $src.EnableProperty.Enabled
                $dirty = $true
            }

            $srcChan = $src.Properties | Where-Object { $_.Key -eq "Channel"} | Select-Object -ExpandProperty Value
            $dstChan = $dst.Properties | Where-Object { $_.Key -eq "Channel"} | Select-Object -ExpandProperty Value
            if ($srcChan -ne $dstChan) {
                Write-Error "Sorting mismatch between source and destination configuration."
                return
            }

            foreach ($srcProp in $src.Properties) {
                $dstProp = $dst.Properties | Where-Object Key -eq $srcProp.Key
                if ($null -eq $dstProp) {
                    Write-Verbose "Key '$($srcProp.Key)' not found on $($dst.Path)"
                    Write-Verbose "Available keys`r`n$($dst.Properties | Select-Object Key, Value | Format-Table)"
                    continue
                }
                if (!$srcProp.IsSettable -or $srcProp.ValueType -eq 'PathList' -or $srcProp.ValueType -eq 'Path') { continue }
                if ($srcProp.Value -ne $dstProp.Value) {
                    Write-Verbose "Changing $($dstProp.DisplayName) to $($srcProp.Value) on $($dst.Path)"
                    $dstProp.Value = $srcProp.Value
                    $dirty = $true
                }
            }
            if ($dirty) {
                if ($dst.ItemCategory -eq "ChildItem") {
                    $result = $lastParent | Set-ConfigurationItem
                } else {
                    $result = $dst | Set-ConfigurationItem
                }

                if (!$result.ValidatedOk) {
                    foreach ($errorResult in $result.ErrorResults) {
                        Write-Error $errorResult.ErrorText
                    }
                }
            }

            if ($src.Children.Count -eq $dst.Children.Count -and $src.Children.Count -gt 0) {
                foreach ($child in $src.Children) {
                    $srcStack.Push($child)
                }
                foreach ($child in $dst.Children) {
                    $dstStack.Push($child)
                }
                if ($dst.ItemCategory -eq "Item") {
                    $lastParent = $dst
                }
            } elseif ($src.Children.Count -ne 0) {
                Write-Warning "Number of child items is not equal on $($src.DisplayName)"
            }
        }
    }
}
class CidrInfo {
    [string] $Cidr
    [IPAddress] $Address
    [int] $Mask

    [IPAddress] $Start
    [IPAddress] $End
    [IPAddress] $SubnetMask
    [IPAddress] $HostMask

    [int] $TotalAddressCount
    [int] $HostAddressCount

    CidrInfo([string] $Cidr) {
        [System.Net.IPAddress]$this.Address, [int]$this.Mask = $Cidr -split '/'
        if ($this.Address.AddressFamily -notin @([System.Net.Sockets.AddressFamily]::InterNetwork, [System.Net.Sockets.AddressFamily]::InterNetworkV6)) {
            throw "CidrInfo is not compatible with AddressFamily $($this.Address.AddressFamily). Expected InterNetwork or InterNetworkV6."
        }
        $min, $max = if ($this.Address.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) { 0, 32 } else { 0, 128 }
        if ($this.Mask -lt $min -or $this.Mask -gt $max) {
            throw "CIDR mask value out of range. Expected a value between $min and $max for AddressFamily $($this.Address.AddressFamily)"
        }
        $hostMaskLength = $max - $this.Mask
        $this.Cidr = $Cidr
        $this.TotalAddressCount = [math]::pow(2, $hostMaskLength)
        # RFC 3021 support is assumed. When the range supports only two hosts, RFC 3021 defines it usable for point-to-point communications but not all systems support this.
        $this.HostAddressCount = if ($hostMaskLength -eq 0) { 1 } elseif ($hostMaskLength -eq 1) { 2 } else { $this.TotalAddressCount - 2 }

        $addressBytes = $this.Address.GetAddressBytes()
        $netMaskBytes = [byte[]]::new($addressBytes.Count)
        $hostMaskBytes = [byte[]]::new($addressBytes.Count)
        $bitCounter = 0
        for ($octet = 0; $octet -lt $addressBytes.Count; $octet++) {
            for ($bit = 0; $bit -lt 8; $bit++) {
                $bitCounter += 1
                $bitValue = 0
                if ($bitCounter -le $this.Mask) {
                    $bitValue = 1
                }
                $netMaskBytes[$octet] = $netMaskBytes[$octet] -bor ( $bitValue -shl ( 7 - $bit ) )
                $hostMaskBytes[$octet] = $netMaskBytes[$octet] -bxor 255
            }
        }
        $this.SubnetMask = [ipaddress]::new($netMaskBytes)
        $this.HostMask = [IPAddress]::new($hostMaskBytes)

        $startBytes = [byte[]]::new($addressBytes.Count)
        $endBytes = [byte[]]::new($addressBytes.Count)
        for ($octet = 0; $octet -lt $addressBytes.Count; $octet++) {
            $startBytes[$octet] = $addressBytes[$octet] -band $netMaskBytes[$octet]
            $endBytes[$octet] = $addressBytes[$octet] -bor $hostMaskBytes[$octet]
        }
        $this.Start = [IPAddress]::new($startBytes)
        $this.End = [IPAddress]::new($endBytes)
    }
}

function Expand-IPRange {
    <#
    .SYNOPSIS
    Expands a start and end IP address or a CIDR notation into an array of IP addresses within the given range.

    .DESCRIPTION
    Accepts start and end IP addresses in the form of IPv4 or IPv6 addresses, and returns each IP
    address falling within the range including the Start and End values.

    The Start and End IP addresses must be in the same address family (IPv4 or IPv6) and if the
    addresses are IPv6, they must have the same scope ID.

    .PARAMETER Start
    Specifies the first IP address in the range to be expanded.

    .PARAMETER End
    Specifies the last IP address in the range to be expanded. Must be greater than or equal to Start.

    .PARAMETER Cidr
    Specifies an IP address range in CIDR notation. Example: 192.168.0.0/23 represents 192.168.0.0-192.168.1.255.

    .PARAMETER AsString
    Specifies that each IP address in the range should be returned as a string instead of an [IPAddress] object.

    .EXAMPLE
    PS C:\> Expand-IPRange -Start 192.168.1.1 -End 192.168.2.255
    Returns 511 IPv4 IPAddress objects.

    .EXAMPLE
    PS C:\> Expand-IPRange -Start fe80::5566:e22e:3f34:5a0f -End fe80::5566:e22e:3f34:5a16
    Returns 8 IPv6 IPAddress objects.

    .EXAMPLE
    PS C:\> Expand-IPRange -Start 10.1.1.100 -End 10.1.10.50 -AsString
    Returns 2255 IPv4 addresses as strings.

    .EXAMPLE
    PS C:\> Expand-IPRange -Cidr 172.16.16.0/23
    Returns IPv4 IPAddress objects from 172.16.16.0 to 172.16.17.255.
    #>
    [CmdletBinding(DefaultParameterSetName = 'FromRange')]
    [OutputType([System.Net.IPAddress], [string])]
    param(
        [Parameter(Mandatory, ParameterSetName = 'FromRange')]
        [ValidateScript({
            if ($_.AddressFamily -in @([System.Net.Sockets.AddressFamily]::InterNetwork, [System.Net.Sockets.AddressFamily]::InterNetworkV6)) {
                return $true
            }
            throw "Start IPAddress is from AddressFamily '$($_.AddressFamily)'. Expected InterNetwork or InterNetworkV6."
        })]
        [System.Net.IPAddress]
        $Start,

        [Parameter(Mandatory, ParameterSetName = 'FromRange')]
        [ValidateScript({
            if ($_.AddressFamily -in @([System.Net.Sockets.AddressFamily]::InterNetwork, [System.Net.Sockets.AddressFamily]::InterNetworkV6)) {
                return $true
            }
            throw "Start IPAddress is from AddressFamily '$($_.AddressFamily)'. Expected InterNetwork or InterNetworkV6."
        })]
        [System.Net.IPAddress]
        $End,

        [Parameter(Mandatory, ParameterSetName = 'FromCidr')]
        [ValidateNotNullOrEmpty()]
        [string]
        $Cidr,

        [Parameter()]
        [switch]
        $AsString
    )

    process {
        if ($PSCmdlet.ParameterSetName -eq 'FromCidr') {
            $cidrInfo = [CidrInfo]$Cidr
            $Start = $cidrInfo.Start
            $End = $cidrInfo.End
        }

        if (-not $Start.AddressFamily.Equals($End.AddressFamily)) {
            throw 'Expand-IPRange received Start and End addresses from different IP address families (IPv4 and IPv6). Both addresses must be of the same IP address family.'
        }

        if ($Start.ScopeId -ne $End.ScopeId) {
            throw 'Expand-IPRange received IPv6 Start and End addresses with different ScopeID values. The ScopeID values must be identical.'
        }

        # Assert that the End IP is greater than or equal to the Start IP.
        $startBytes = $Start.GetAddressBytes()
        $endBytes = $End.GetAddressBytes()
        for ($i = 0; $i -lt $startBytes.Length; $i++) {
            if ($endBytes[$i] -lt $startBytes[$i]) {
                throw 'Expand-IPRange must receive an End IPAddress which is greater than or equal to the Start IPAddress'
            }
            if ($endBytes[$i] -gt $startBytes[$i]) {
                # We can break early if a higher-order byte from the End address is greater than the matching byte of the Start address
                break
            }
        }

        $current = $Start
        while ($true) {
            if ($AsString) {
                Write-Output $current.ToString()
            }
            else {
                Write-Output $current
            }

            if ($current.Equals($End)) {
                break
            }

            $bytes = $current.GetAddressBytes()
            for ($i = $bytes.Length - 1; $i -ge 0; $i--) {
                if ($bytes[$i] -lt 255) {
                    $bytes[$i] += 1
                    break
                }
                $bytes[$i] = 0
            }
            if ($null -ne $current.ScopeId) {
                $current = [System.Net.IPAddress]::new($bytes, $current.ScopeId)
            }
            else {
                $current = [System.Net.IPAddress]::new($bytes)
            }
        }
    }
}
function FillChildren {
    [CmdletBinding()]
    [OutputType([VideoOS.ConfigurationApi.ClientService.ConfigurationItem])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.ConfigurationApi.ClientService.ConfigurationItem]
        $ConfigurationItem,

        [Parameter()]
        [int]
        $Depth = 1
    )

    process {
        $stack = New-Object System.Collections.Generic.Stack[VideoOS.ConfigurationApi.ClientService.ConfigurationItem]
        $stack.Push($ConfigurationItem)
        while ($stack.Count -gt 0) {
            $Depth = $Depth - 1
            $item = $stack.Pop()
            $item.Children = $item | Get-ConfigurationItem -ChildItems
            $item.ChildrenFilled = $true
            if ($Depth -gt 0) {
                $item.Children | Foreach-Object {
                    $stack.Push($_)
                }
            }
        }
        Write-Output $ConfigurationItem
    }
}
function Find-XProtectDeviceDialog {
    [CmdletBinding()]
    param ()

    process {
        Add-Type -AssemblyName PresentationFramework
        $xaml = [xml]@"
        <Window
                xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
                xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                xmlns:local="clr-namespace:Search_XProtect"
                Title="Search XProtect" Height="500" Width="800"
                FocusManager.FocusedElement="{Binding ElementName=cboItemType}">
            <Grid>
                <GroupBox Name="gboAdvanced" Header="Advanced Parameters" HorizontalAlignment="Left" Height="94" Margin="506,53,0,0" VerticalAlignment="Top" Width="243"/>
                <Label Name="lblItemType" Content="Item Type" HorizontalAlignment="Left" Margin="57,22,0,0" VerticalAlignment="Top"/>
                <ComboBox Name="cboItemType" HorizontalAlignment="Left" Margin="124,25,0,0" VerticalAlignment="Top" Width="120" TabIndex="0">
                    <ComboBoxItem Content="Camera" HorizontalAlignment="Left" Width="118"/>
                    <ComboBoxItem Content="Hardware" HorizontalAlignment="Left" Width="118"/>
                    <ComboBoxItem Content="InputEvent" HorizontalAlignment="Left" Width="118"/>
                    <ComboBoxItem Content="Metadata" HorizontalAlignment="Left" Width="118"/>
                    <ComboBoxItem Content="Microphone" HorizontalAlignment="Left" Width="118"/>
                    <ComboBoxItem Content="Output" HorizontalAlignment="Left" Width="118"/>
                    <ComboBoxItem Content="Speaker" HorizontalAlignment="Left" Width="118"/>
                </ComboBox>
                <Label Name="lblName" Content="Name" HorizontalAlignment="Left" Margin="77,53,0,0" VerticalAlignment="Top" IsEnabled="False"/>
                <Label Name="lblPropertyName" Content="Property Name" HorizontalAlignment="Left" Margin="519,80,0,0" VerticalAlignment="Top" IsEnabled="False"/>
                <ComboBox Name="cboPropertyName" HorizontalAlignment="Left" Margin="614,84,0,0" VerticalAlignment="Top" Width="120" IsEnabled="False" TabIndex="5"/>
                <TextBox Name="txtName" HorizontalAlignment="Left" Height="23" Margin="124,56,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="187" IsEnabled="False" TabIndex="1"/>
                <Button Name="btnSearch" Content="Search" HorizontalAlignment="Left" Margin="306,154,0,0" VerticalAlignment="Top" Width="75" TabIndex="7" IsEnabled="False"/>
                <DataGrid Name="dgrResults" HorizontalAlignment="Left" Height="207" Margin="36,202,0,0" VerticalAlignment="Top" Width="719" IsReadOnly="True"/>
                <Label Name="lblAddress" Content="IP Address" HorizontalAlignment="Left" Margin="53,84,0,0" VerticalAlignment="Top" IsEnabled="False"/>
                <TextBox Name="txtAddress" HorizontalAlignment="Left" Height="23" Margin="124,87,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120" IsEnabled="False" TabIndex="2"/>
                <Label Name="lblEnabledFilter" Content="Enabled/Disabled" HorizontalAlignment="Left" Margin="506,22,0,0" VerticalAlignment="Top" IsEnabled="False"/>
                <ComboBox Name="cboEnabledFilter" HorizontalAlignment="Left" Margin="614,26,0,0" VerticalAlignment="Top" Width="120" IsEnabled="False" TabIndex="4">
                    <ComboBoxItem Content="Enabled" HorizontalAlignment="Left" Width="118"/>
                    <ComboBoxItem Content="Disabled" HorizontalAlignment="Left" Width="118"/>
                    <ComboBoxItem Name="cbiEnabledAll" Content="All" HorizontalAlignment="Left" Width="118" IsSelected="True"/>
                </ComboBox>
                <Label Name="lblMACAddress" Content="MAC Address" HorizontalAlignment="Left" Margin="37,115,0,0" VerticalAlignment="Top" IsEnabled="False"/>
                <TextBox Name="txtMACAddress" HorizontalAlignment="Left" Height="23" Margin="124,118,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120" IsEnabled="False" TabIndex="3"/>
                <Label Name="lblPropertyValue" Content="Property Value" HorizontalAlignment="Left" Margin="522,108,0,0" VerticalAlignment="Top" IsEnabled="False"/>
                <TextBox Name="txtPropertyValue" HorizontalAlignment="Left" Height="23" Margin="614,111,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120" IsEnabled="False" TabIndex="6"/>
                <Button Name="btnExportCSV" Content="Export CSV" HorizontalAlignment="Left" Margin="680,429,0,0" VerticalAlignment="Top" Width="75" TabIndex="9" IsEnabled="False"/>
                <Label Name="lblNoResults" Content="No results found!" HorizontalAlignment="Left" Margin="345,175,0,0" VerticalAlignment="Top" Foreground="Red" Visibility="Hidden"/>
                <Button Name="btnResetForm" Content="Reset Form" HorizontalAlignment="Left" Margin="414,154,0,0" VerticalAlignment="Top" Width="75" TabIndex="8"/>
                <Label Name="lblTotalResults" Content="Total Results:" HorizontalAlignment="Left" Margin="32,423,0,0" VerticalAlignment="Top" FontWeight="Bold"/>
                <TextBox Name="txtTotalResults" HorizontalAlignment="Left" Height="23" Margin="120,427,0,0" VerticalAlignment="Top" Width="53" IsEnabled="False"/>
                <Label Name="lblPropertyNameBlank" Content="Property Name cannot be blank if Property&#xD;&#xA;Value has an entry." HorizontalAlignment="Left" Margin="507,152,0,0" VerticalAlignment="Top" Foreground="Red" Width="248" Height="45" Visibility="Hidden"/>
                <Label Name="lblPropertyValueBlank" Content="Property Value cannot be blank if Property&#xA;Name has a selection." HorizontalAlignment="Left" Margin="507,152,0,0" VerticalAlignment="Top" Foreground="Red" Width="248" Height="45" Visibility="Hidden"/>
            </Grid>
        </Window>
"@

        function Clear-Results {
            $var_dgrResults.Columns.Clear()
            $var_dgrResults.Items.Clear()
            $var_txtTotalResults.Clear()
            $var_lblNoResults.Visibility = "Hidden"
            $var_lblPropertyNameBlank.Visibility = "Hidden"
            $var_lblPropertyValueBlank.Visibility = "Hidden"
        }

        $reader = [system.xml.xmlnodereader]::new($xaml)
        $window = [windows.markup.xamlreader]::Load($reader)
        $searchResults = $null

        # Create variables based on form control names.
        # Variable will be named as 'var_<control name>'
        $xaml.SelectNodes("//*[@Name]") | ForEach-Object {
            #"trying item $($_.Name)"
            try {
                Set-Variable -Name "var_$($_.Name)" -Value $window.FindName($_.Name) -ErrorAction Stop
            } catch {
                throw
            }
        }
        # Get-Variable var_*

        $iconBase64 = "AAABAAEAICAAAAEAIACoEAAAFgAAACgAAAAgAAAAQAAAAAEAIAAAAAAAABAAAMMOAADDDgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADamQCA2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA2pkAgNqZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADamQCA2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA2pkAgNqZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADamQCA2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA2pkAgNqZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADamQCA2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA2pkAgNqZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADamQCA2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAA2pkAgNqZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAADamQCA2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgNqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQCAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAIAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//5////8P///+B////AP///gB///wAP//4AB//8AAP/+AAB//AAAP/gAAB/wAAAP4AAAB8AAAAOAAAABAAAAAAAAAACAAAABwAAAA+AAAAfwAAAP+AAAH/wAAD/+AAB//wAA//+AAf//wAP//+AH///wD///+B////w////+f/8="
        $iconBytes = [Convert]::FromBase64String($iconBase64)
        $window.Icon = $iconBytes

        $assembly = [System.Reflection.Assembly]::GetAssembly([VideoOS.Platform.ConfigurationItems.Hardware])

        $excludedItems = "Folder|Path|Icon|Enabled|DisplayName|RecordingFramerate|ItemCategory|Wrapper|Address|Channel"

        $var_cboItemType.Add_SelectionChanged( {
                param($sender, $e)
                $itemType = $e.AddedItems[0].Content

                $var_cboPropertyName.Items.Clear()
                $var_dgrResults.Columns.Clear()
                $var_dgrResults.Items.Clear()
                $var_txtTotalResults.Clear()
                $var_txtPropertyValue.Clear()
                $var_lblNoResults.Visibility = "Hidden"
                $var_lblPropertyNameBlank.Visibility = "Hidden"
                $var_lblPropertyValueBlank.Visibility = "Hidden"

                $properties = ($assembly.GetType("VideoOS.Platform.ConfigurationItems.$itemType").DeclaredProperties | Where-Object { $_.PropertyType.Name -eq 'String' }).Name + ([VideoOS.Platform.ConfigurationItems.IConfigurationChildItem].DeclaredProperties | Where-Object { $_.PropertyType.Name -eq 'String' }).Name | Where-Object { $_ -notmatch $excludedItems }
                foreach ($property in $properties) {
                    $newComboboxItem = [System.Windows.Controls.ComboBoxItem]::new()
                    $newComboboxItem.AddChild($property)
                    $var_cboPropertyName.Items.Add($newComboboxItem)
                }

                $sortDescription = [System.ComponentModel.SortDescription]::new("Content", "Ascending")
                $var_cboPropertyName.Items.SortDescriptions.Add($sortDescription)

                $var_cboEnabledFilter.IsEnabled = $true
                $var_lblEnabledFilter.IsEnabled = $true
                $var_cboPropertyName.IsEnabled = $true
                $var_lblPropertyName.IsEnabled = $true
                $var_txtPropertyValue.IsEnabled = $true
                $var_lblPropertyValue.IsEnabled = $true
                $var_txtName.IsEnabled = $true
                $var_lblName.IsEnabled = $true
                $var_btnSearch.IsEnabled = $true

                if ($itemType -eq "Hardware") {
                    $var_txtAddress.IsEnabled = $true
                    $var_lblAddress.IsEnabled = $true
                    $var_txtMACAddress.IsEnabled = $true
                    $var_lblMACAddress.IsEnabled = $true
                } else {
                    $var_txtAddress.IsEnabled = $false
                    $var_txtAddress.Clear()
                    $var_lblAddress.IsEnabled = $false
                    $var_txtMACAddress.IsEnabled = $false
                    $var_txtMACAddress.Clear()
                    $var_lblMACAddress.IsEnabled = $false
                }
            })

        $var_txtName.Add_TextChanged( {
                Clear-Results
            })

        $var_txtAddress.Add_TextChanged( {
                Clear-Results
            })

        $var_txtMACAddress.Add_TextChanged( {
                Clear-Results
            })

        $var_cboEnabledFilter.Add_SelectionChanged( {
                Clear-Results
            })

        $var_cboPropertyName.Add_SelectionChanged( {
                Clear-Results
            })

        $var_txtPropertyValue.Add_TextChanged( {
                Clear-Results
            })

        $var_btnSearch.Add_Click( {
                if (-not [string]::IsNullOrEmpty($var_cboPropertyName.Text) -and [string]::IsNullOrEmpty($var_txtPropertyValue.Text)) {
                    $var_lblPropertyValueBlank.Visibility = "Visible"
                    Return
                } elseif ([string]::IsNullOrEmpty($var_cboPropertyName.Text) -and -not [string]::IsNullOrEmpty($var_txtPropertyValue.Text)) {
                    $var_lblPropertyNameBlank.Visibility = "Visible"
                    Return
                }

                $script:searchResults = Find-XProtectDeviceSearch -ItemType $var_cboItemType.Text -Name $var_txtName.Text -Address $var_txtAddress.Text -MAC $var_txtMACAddress.Text -Enabled $var_cboEnabledFilter.Text -PropertyName $var_cboPropertyName.Text -PropertyValue $var_txtPropertyValue.Text
                if ($null -ne $script:searchResults) {
                    $var_btnExportCSV.IsEnabled = $true
                } else {
                    $var_btnExportCSV.IsEnabled = $false
                }
            })

        $var_btnExportCSV.Add_Click( {
                $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
                $saveDialog.Title = "Save As CSV"
                $saveDialog.Filter = "Comma delimited (*.csv)|*.csv"

                $saveAs = $saveDialog.ShowDialog()

                if ($saveAs -eq $true) {
                    $script:searchResults | Export-Csv -Path $saveDialog.FileName -NoTypeInformation
                }
            })

        $var_btnResetForm.Add_Click( {
                $var_dgrResults.Columns.Clear()
                $var_dgrResults.Items.Clear()
                $var_cboItemType.SelectedItem = $null
                $var_cboEnabledFilter.IsEnabled = $false
                $var_lblEnabledFilter.IsEnabled = $false
                $var_cbiEnabledAll.IsSelected = $true
                $var_cboPropertyName.IsEnabled = $false
                $var_cboPropertyName.Items.Clear()
                $var_lblPropertyName.IsEnabled = $false
                $var_txtPropertyValue.IsEnabled = $false
                $var_txtPropertyValue.Clear()
                $var_lblPropertyValue.IsEnabled = $false
                $var_txtName.IsEnabled = $false
                $var_txtName.Clear()
                $var_lblName.IsEnabled = $false
                $var_btnSearch.IsEnabled = $false
                $var_btnExportCSV.IsEnabled = $false
                $var_txtAddress.IsEnabled = $false
                $var_txtAddress.Clear()
                $var_lblAddress.IsEnabled = $false
                $var_txtMACAddress.IsEnabled = $false
                $var_txtMACAddress.Clear()
                $var_lblMACAddress.IsEnabled = $false
                $var_txtTotalResults.Clear()
                $var_lblNoResults.Visibility = "Hidden"
                $var_lblPropertyNameBlank.Visibility = "Hidden"
                $var_lblPropertyValueBlank.Visibility = "Hidden"
            })

        $null = $window.ShowDialog()
    }
}

function Find-XProtectDeviceSearch {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ItemType,
        [Parameter(Mandatory = $false)]
        [string]$Name,
        [Parameter(Mandatory = $false)]
        [string]$Address,
        [Parameter(Mandatory = $false)]
        [string]$MAC,
        [Parameter(Mandatory = $false)]
        [string]$Enabled,
        [Parameter(Mandatory = $false)]
        [string]$PropertyName,
        [Parameter(Mandatory = $false)]
        [string]$PropertyValue
    )

    process {
        $var_dgrResults.Columns.Clear()
        $var_dgrResults.Items.Clear()
        $var_lblNoResults.Visibility = "Hidden"
        $var_lblPropertyNameBlank.Visibility = "Hidden"
        $var_lblPropertyValueBlank.Visibility = "Hidden"

        if ([string]::IsNullOrEmpty($PropertyName) -or [string]::IsNullOrEmpty($PropertyValue)) {
            $PropertyName = "Id"
            $PropertyValue = $null
        }

        if ($ItemType -eq "Hardware" -and $null -eq [string]::IsNullOrEmpty($MAC)) {
            $results = [array](Find-XProtectDevice -ItemType $ItemType -MacAddress $MAC -EnableFilter $Enabled -Properties @{Name = $Name; Address = $Address; $PropertyName = $PropertyValue })
        } elseif ($ItemType -eq "Hardware" -and $null -ne [string]::IsNullOrEmpty($MAC)) {
            $results = [array](Find-XProtectDevice -ItemType $ItemType -EnableFilter $Enabled -Properties @{Name = $Name; Address = $Address; $PropertyName = $PropertyValue })
        } else {
            $results = [array](Find-XProtectDevice -ItemType $ItemType -EnableFilter $Enabled -Properties @{Name = $Name; $PropertyName = $PropertyValue })
        }

        if ($null -ne $results) {
            #$columnNames = ($results | Get-Member | Where-Object {$_.MemberType -eq 'NoteProperty'}).Name
            $columnNames = $results[0].PsObject.Properties | ForEach-Object { $_.Name }
        } else {
            $var_lblNoResults.Visibility = "Visible"
        }

        foreach ($columnName in $columnNames) {
            $newColumn = [System.Windows.Controls.DataGridTextColumn]::new()
            $newColumn.Header = $columnName
            $newColumn.Binding = New-Object System.Windows.Data.Binding($columnName)
            $newColumn.Width = "SizeToCells"
            $var_dgrResults.Columns.Add($newColumn)
        }

        if ($ItemType -eq "Hardware") {
            foreach ($result in $results) {
                $var_dgrResults.AddChild([pscustomobject]@{Hardware = $result.Hardware; RecordingServer = $result.RecordingServer })
            }
        } else {
            foreach ($result in $results) {
                $var_dgrResults.AddChild([pscustomobject]@{$columnNames[0] = $result.((Get-Variable -Name columnNames).Value[0]); Hardware = $result.Hardware; RecordingServer = $result.RecordingServer })
            }
        }

        $var_txtTotalResults.Text = $results.count
    }
    end {
        return $results
    }
}
function Get-DevicesByRecorder {
    <#
    .SYNOPSIS
        Gets all enabled cameras in a hashtable indexed by recording server id.
    .DESCRIPTION
        This cmdlet quickly returns a hashtable where the keys are recording
        server ID's and the values are lists of "VideoOS.Platform.Item" objects.

        The cmdlet will complete much quicker than if we were to use
        Get-RecordingServer | Get-VmsCamera, because it does not rely on the
        configuration API at all. Instead, it has the same functionality as
        Smart Client where the command "sees" only the devices that are enabled
        and loaded by the Recording Server.
    .EXAMPLE
        Get-CamerasByRecorder
        Name                           Value
        ----                           -----
        bb82b2cd-0bb9-4c88-9cb8-128... {Canon VB-M40 (192.168.101.64) - Camera 1}
        f9dc2bcd-faea-4138-bf5a-32c... {Axis P1375 (10.1.77.178) - Camera 1, Test Cam}

        This is what the output would look like on a small system.
    .OUTPUTS
        [hashtable]
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Id')]
        [guid[]]
        $RecordingServerId,

        [Parameter()]
        [Alias('Kind')]
        [ValidateSet('Camera', 'Microphone', 'Speaker', 'Metadata', IgnoreCase = $false)]
        [string[]]
        $DeviceType = 'Camera'
    )

    process {
        $config = [videoos.platform.configuration]::Instance
        $serverKind = [VideoOS.Platform.Kind]::Server
        $selectedKinds = @(($DeviceType | ForEach-Object { [VideoOS.Platform.Kind]::$_ }))
        $systemHierarchy = [VideoOS.Platform.ItemHierarchy]::SystemDefined

        $stack = [Collections.Generic.Stack[VideoOS.Platform.Item]]::new()
        $rootItems = $config.GetItems($systemHierarchy)
        foreach ($mgmtSrv in $rootItems | Where-Object { $_.FQID.Kind -eq $serverKind }) {
            foreach ($recorder in $mgmtSrv.GetChildren()) {
                if ($recorder.FQID.Kind -eq $serverKind -and ($RecordingServerId.Count -eq 0 -or $recorder.FQID.ObjectId -in $RecordingServerId)) {
                    $stack.Push($recorder)
                }
            }
        }

        $result = @{}
        $lastServerId = $null
        while ($stack.Count -gt 0) {
            $item = $stack.Pop()
            if ($item.FQID.Kind -eq $serverKind) {
                $lastServerId = $item.FQID.ObjectId
                $result.$lastServerId = [Collections.Generic.List[VideoOS.Platform.Item]]::new()
            } elseif ($item.FQID.Kind -in $selectedKinds -and $item.FQID.FolderType -eq 'No') {
                $result.$lastServerId.Add($item)
                continue
            }

            if ($item.HasChildren -ne 'No' -and ($item.FQID.Kind -eq $serverKind -or $item.FQID.Kind -in $selectedKinds)) {
                foreach ($child in $item.GetChildren()) {
                    if ($child.FQID.Kind -in $selectedKinds) {
                        $stack.Push($child)
                    }
                }
            }
        }
        Write-Output $result
    }
}
function Get-HttpSslCertThumbprint {
    <#
    .SYNOPSIS
        Gets the certificate thumbprint from the sslcert binding information put by netsh http show sslcert ipport=$IPPort
    .DESCRIPTION
        Gets the certificate thumbprint from the sslcert binding information put by netsh http show sslcert ipport=$IPPort.
        Returns $null if no binding is present for the given ip:port value.
    .PARAMETER IPPort
        The ip:port string representing the binding to retrieve the thumbprint from.
    .EXAMPLE
        Get-MobileServerSslCertThumbprint 0.0.0.0:8082
        Gets the sslcert thumbprint for the binding found matching 0.0.0.0:8082 which is the default HTTPS IP and Port for
        XProtect Mobile Server. The value '0.0.0.0' represents 'all interfaces' and 8082 is the default https port.
    #>
    [CmdletBinding()]
    param (
        [parameter(Mandatory)]
        [string]
        $IPPort
    )
    process {
        $netshOutput = [string](netsh.exe http show sslcert ipport=$IPPort)

        if (!$netshOutput.Contains('Certificate Hash')) {
            Write-Error "No SSL certificate binding found for $ipPort"
            return
        }

        if ($netshOutput -match "Certificate Hash\s+:\s+(\w+)\s+") {
            $Matches[1]
        } else {
            Write-Error "Certificate Hash not found for $ipPort"
        }
    }
}
function Get-ProcessOutput
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]
        $FilePath,
        [Parameter()]
        [string[]]
        $ArgumentList
    )
    
    process {
        try {
            $process = New-Object System.Diagnostics.Process
            $process.StartInfo.UseShellExecute = $false
            $process.StartInfo.RedirectStandardOutput = $true
            $process.StartInfo.RedirectStandardError = $true
            $process.StartInfo.FileName = $FilePath
            $process.StartInfo.CreateNoWindow = $true

            if($ArgumentList) { $process.StartInfo.Arguments = $ArgumentList }
            Write-Verbose "Executing $($FilePath) with the following arguments: $([string]::Join(' ', $ArgumentList))"
            $null = $process.Start()
    
            [pscustomobject]@{
                StandardOutput = $process.StandardOutput.ReadToEnd()
                StandardError = $process.StandardError.ReadToEnd()
                ExitCode = $process.ExitCode
            }
        }
        finally {
            $process.Dispose()
        }
        
    }
}
function GetCodecValueFromStream {
    param([VideoOS.Platform.ConfigurationItems.StreamChildItem]$Stream)

    $res = $Stream.Properties.GetValue("Codec")
    if ($null -ne $res) {
        ($Stream.Properties.GetValueTypeInfoCollection("Codec") | Where-Object Value -eq $res).Name
        return
    }
}
function GetFpsValueFromStream {
    param([VideoOS.Platform.ConfigurationItems.StreamChildItem]$Stream)

    $res = $Stream.Properties.GetValue("FPS")
    if ($null -ne $res) {
        $val = ($Stream.Properties.GetValueTypeInfoCollection("FPS") | Where-Object Value -eq $res).Name
        if ($null -eq $val) {
            $res
        }
        else {
            $val
        }
        return
    }

    $res = $Stream.Properties.GetValue("Framerate")
    if ($null -ne $res) {
        $val = ($Stream.Properties.GetValueTypeInfoCollection("Framerate") | Where-Object Value -eq $res).Name
        if ($null -eq $val) {
            $res
        }
        else {
            $val
        }
        return
    }
}
function GetResolutionValueFromStream {
    param([VideoOS.Platform.ConfigurationItems.StreamChildItem]$Stream)

    $res = $Stream.Properties.GetValue("StreamProperty")
    if ($null -ne $res) {
        ($Stream.Properties.GetValueTypeInfoCollection("StreamProperty") | Where-Object Value -eq $res).Name
        return
    }

    $res = $Stream.Properties.GetValue("Resolution")
    if ($null -ne $res) {
        ($Stream.Properties.GetValueTypeInfoCollection("Resolution") | Where-Object Value -eq $res).Name
        return
    }
}
function Set-CertKeyPermission {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        # Specifies the certificate store path to locate the certificate specified in Thumbprint. Example: Cert:\LocalMachine\My
        [Parameter()]
        [string]
        $CertificateStore = 'Cert:\LocalMachine\My',

        # Specifies the thumbprint of the certificate to which private key access should be updated.
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]
        $Thumbprint,

        # Specifies the Windows username for the identity to which permissions should be granted.
        [Parameter(Mandatory)]
        [string]
        $UserName,

        # Specifies the level of access to grant to the private key.
        [Parameter()]
        [ValidateSet('Read', 'FullControl')]
        [string]
        $Permission = 'Read',

        # Specifies the access type for the Access Control List rule.
        [Parameter()]
        [ValidateSet('Allow', 'Deny')]
        [string]
        $PermissionType = 'Allow'
    )

    process {
        <#
            There is a LOT of error checking in this function as it seems that certificates are not
            always consistently storing their private keys in predictable places. I've found private
            keys for RSA certs in ProgramData\Microsoft\Crypto\Keys instead of
            ProgramData\Microsoft\Crypto\RSA\MachineKeys, I've seen the UniqueName property contain
            a value representing the file name of the certificate private key file somewhere in the
            ProgramData\Microsoft\Crypto folder, and I've seen the UniqueName property contain a
            full file path to the private key file. I've also found that some RSA certs require you
            to use the RSA extension method to retrieve the private key, even though it seems like
            you should expect to find it in the PrivateKey property when retrieving the certificate
            from Get-ChildItem Cert:\LocalMachine\My.
        #>

        $certificate = Get-ChildItem -Path $CertificateStore | Where-Object Thumbprint -eq $Thumbprint
        Write-Verbose "Processing certificate for $($certificate.Subject) with thumbprint $($certificate.Thumbprint)"
        if ($null -eq $certificate) {
            Write-Error "Certificate not found in certificate store '$CertificateStore' matching thumbprint '$Thumbprint'"
            return
        }
        if (-not $certificate.HasPrivateKey) {
            Write-Error "Certificate with friendly name '$($certificate.FriendlyName)' issued to subject '$($certificate.Subject)' does not have a private key attached."
            return
        }
        $privateKey = $null
        switch ($certificate.PublicKey.EncodedKeyValue.Oid.FriendlyName) {
            'RSA' {
                $privateKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($certificate)
            }

            'ECC' {
                $privateKey = [System.Security.Cryptography.X509Certificates.ECDsaCertificateExtensions]::GetECDsaPrivateKey($certificate)
            }

            'DSA' {
                Write-Error "Use of DSA-based certificates is not recommended, and not supported by this command. See https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.dsa?view=net-5.0"
                return
            }

            Default { Write-Error "`$certificate.PublicKey.EncodedKeyValue.Oid.FriendlyName was '$($certificate.PublicKey.EncodedKeyValue.Oid.FriendlyName)'. Expected RSA, DSA or ECC."; return }
        }
        if ($null -eq $privateKey) {
            Write-Error "Certificate with friendly name '$($certificate.FriendlyName)' issued to subject '$($certificate.Subject)' does not have a private key attached."
            return
        }
        if ([string]::IsNullOrWhiteSpace($privateKey.Key.UniqueName)) {
            Write-Error "Certificate with friendly name '$($certificate.FriendlyName)' issued to subject '$($certificate.Subject)' does not have a value for the private key's UniqueName property so we cannot find the file on the filesystem associated with the private key."
            return
        }

        if (Test-Path -LiteralPath $privateKey.Key.UniqueName) {
            $privateKeyFile = Get-Item -Path $privateKey.Key.UniqueName
        }
        else {
            $privateKeyFile = Get-ChildItem -Path (Join-Path -Path ([system.environment]::GetFolderPath([system.environment+specialfolder]::CommonApplicationData)) -ChildPath ([io.path]::combine('Microsoft', 'Crypto'))) -Filter $privateKey.Key.UniqueName -Recurse -ErrorAction Ignore
            if ($null -eq $privateKeyFile) {
                Write-Error "No private key file found matching UniqueName '$($privateKey.Key.UniqueName)'"
                return
            }
            if ($privateKeyFile.Count -gt 1) {
                Write-Error "Found more than one private key file matching UniqueName '$($privateKey.Key.UniqueName)'"
                return
            }
        }

        $privateKeyPath = $privateKeyFile.FullName
        if (-not (Test-Path -Path $privateKeyPath)) {
            Write-Error "Expected to find private key file at '$privateKeyPath' but the file does not exist. You may need to re-install the certificate in the certificate store"
            return
        }

        $acl = Get-Acl -Path $privateKeyPath
        $rule = [Security.AccessControl.FileSystemAccessRule]::new($UserName, $Permission, $PermissionType)
        $acl.AddAccessRule($rule)
        if ($PSCmdlet.ShouldProcess($privateKeyPath, "Add FileSystemAccessRule")) {
            $acl | Set-Acl -Path $privateKeyPath
        }
    }
}
function Add-VmsFailoverGroup {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([VideoOS.Platform.ConfigurationItems.FailoverGroup])]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ValueFromPipeline)]
        [string]
        $Name,

        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateNotNull()]
        [string]
        $Description = ''
    )

    begin {
        Assert-VmsVersion -MinimumVersion 21.2 -ErrorAction Stop
    }

    process {
        $ms = Get-ManagementServer
        if ($PSCmdlet.ShouldProcess($ms.Name, "Create failover group named '$Name'")) {
            $task = $ms.FailoverGroupFolder.AddFailoverGroup($Name, $Description)
            if ($task.State -ne 'Success') {
                Write-Error "Add-VmsFailoverGroup encounted an error. $($task.ErrorText.Trim('.'))."
                return
            }
            $id = $task.Path.Substring(14, 36)
            Get-VmsFailoverGroup -Id $id
        }
    }
}
function Add-VmsHardware {
    [CmdletBinding()]
    [OutputType([VideoOS.Platform.ConfigurationItems.Hardware])]
    param (
        [Parameter(ParameterSetName = 'FromHardwareScan', Mandatory, ValueFromPipeline)]
        [VmsHardwareScanResult[]]
        $HardwareScan,

        [Parameter(ParameterSetName = 'Manual', Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.RecordingServer]
        $RecordingServer,

        [Parameter(ParameterSetName = 'Manual', Mandatory, ValueFromPipelineByPropertyName)]
        [Alias('Address')]
        [uri]
        $HardwareAddress,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]
        $Name,

        [Parameter(ParameterSetName = 'Manual')]
        [int]
        $DriverNumber,

        [Parameter(ParameterSetName = 'Manual', ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [string]
        $HardwareDriverPath,

        [Parameter(ParameterSetName = 'Manual', Mandatory)]
        [pscredential]
        $Credential,

        [Parameter()]
        [switch]
        $SkipConfig,

        [Parameter()]
        [switch]
        $Force
    )

    process {
        $tasks = New-Object System.Collections.Generic.List[VideoOS.Platform.ConfigurationItems.ServerTask]
        switch ($PSCmdlet.ParameterSetName) {
            'Manual' {
                if ([string]::IsNullOrWhiteSpace($HardwareDriverPath)) {
                    if ($MyInvocation.BoundParameters.ContainsKey('DriverNumber')) {
                        $hardwareDriver = $RecordingServer.HardwareDriverFolder.HardwareDrivers | Where-Object Number -eq $DriverNumber
                        if ($null -ne $hardwareDriver) {
                            Write-Verbose "Mapped DriverNumber $DriverNumber to $($hardwareDriver.Name)"
                            $HardwareDriverPath = $hardwareDriver.Path
                        } else {
                            Write-Error "Failed to find hardware driver matching driver number $DriverNumber on Recording Server '$($RecordingServer.Name)'"
                            return
                        }
                    } else {
                        Write-Error "Add-VmsHardware cannot continue without either the HardwareDriverPath or the user-friendly driver number found in the supported hardware list."
                        return
                    }
                }
                $serverTask = $RecordingServer.AddHardware($HardwareAddress, $HardwareDriverPath, $Credential.UserName, $Credential.Password)
                $tasks.Add($serverTask)
            }
            'FromHardwareScan' {
                if ($HardwareScan.HardwareScanValidated -contains $false) {
                    Write-Warning "One or more scanned hardware could not be validated. These entries will be skipped."
                }
                if ($HardwareScan.MacAddressExistsLocal -contains $true) {
                    Write-Warning "One or more scanned hardware already exist on the target recording server. These entries will be skipped."
                }
                if ($HardwareScan.MacAddressExistsGlobal -contains $true -and -not $Force) {
                    Write-Warning "One or more scanned hardware already exist on another recording server. These entries will be skipped since the Force switch was not used."
                }
                foreach ($scan in $HardwareScan | Where-Object { $_.HardwareScanValidated -and -not $_.MacAddressExistsLocal }) {
                    if ($scan.MacAddressExistsGlobal -and -not $Force) {
                        continue
                    }
                    Write-Verbose "Adding $($scan.HardwareAddress) to $($scan.RecordingServer.Name) using driver identified by $($scan.HardwareDriverPath)"
                    $serverTask = $scan.RecordingServer.AddHardware($scan.HardwareAddress, $scan.HardwareDriverPath, $scan.UserName, $scan.Password)
                    $tasks.Add($serverTask)
                }
            }
        }
        if ($tasks.Count -eq 0) {
            return
        }
        Write-Verbose "Awaiting $($tasks.Count) AddHardware requests"
        Write-Verbose "Tasks: $([string]::Join(', ', $tasks.Path))"
        Wait-VmsTask -Path $tasks.Path -Title "Adding hardware to recording server(s) on site $((Get-Site).Name)" -Cleanup | Foreach-Object {
            $vmsTask = [VmsTaskResult]$_
            if ($vmsTask.State -eq [VmsTaskState]::Success) {
                $hardwareId = $vmsTask.Path.Substring(9, 36)
                if (-not $SkipConfig) {
                    Set-NewHardwareConfig -HardwarePath $vmsTask.Path -Name $Name
                }
                $newHardware = Get-Hardware -HardwareId $hardwareId
                if ($null -ne $newHardware) {
                    Write-Output $newHardware
                }
            } else {
                Write-Error "Add-VmsHardware failed with error code $($vmsTask.ErrorCode). $($vmsTask.ErrorText)"
            }
        }
    }
}

function Set-NewHardwareConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]
        $HardwarePath,

        [Parameter()]
        [string]
        $Name
    )

    process {
        $hardwareId = $HardwarePath.Substring(9, 36)
        $newHardware = Get-Hardware -HardwareId $hardwareId

        $systemInfo = [videoos.platform.configuration]::Instance.FindSystemInfo((Get-Site).FQID.ServerId, $true)
        $version = $systemInfo.Properties.ProductVersion -as [version]
        $itemTypes = @('Camera')
        if (-not [string]::IsNullOrWhiteSpace($Name)) {
            $itemTypes += 'Microphone', 'Speaker', 'Metadata', 'InputEvent', 'Output'
        }
        if ($version -ge '20.2') {
            $newHardware.FillChildren($itemTypes)
        }

        $newHardware.Enabled = $true
        if (-not [string]::IsNullOrWhiteSpace($Name)) {
            $newHardware.Name = $Name
        }
        $newHardware.Save()

        foreach ($itemType in $itemTypes) {
            foreach ($item in $newHardware."$($itemType)Folder"."$($itemType)s") {
                if (-not [string]::IsNullOrWhiteSpace($Name)) {
                    $newName = '{0} - {1} {2}' -f $Name, $itemType.Replace('Event', ''), ($item.Channel + 1)
                    $item.Name = $newName
                }
                if ($itemType -eq 'Camera' -and $item.Channel -eq 0) {
                    $item.Enabled = $true
                }
                $item.Save()
            }
        }
    }
}
function ConvertFrom-ConfigurationItem {
    [CmdletBinding()]
    param(
        # Specifies the Milestone Configuration API 'Path' value of the configuration item. For example, 'Hardware[a6756a0e-886a-4050-a5a5-81317743c32a]' where the guid is the ID of an existing Hardware item.
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]
        $Path,

        # Specifies the Milestone 'ItemType' value such as 'Camera', 'Hardware', or 'InputEvent'
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]
        $ItemType
    )

    begin {
        $assembly = [System.Reflection.Assembly]::GetAssembly([VideoOS.Platform.ConfigurationItems.Hardware])
        $serverId = (Get-Site -ErrorAction Stop).FQID.ServerId
    }

    process {
        if ($Path -eq '/') {
            [VideoOS.Platform.ConfigurationItems.ManagementServer]::new($serverId)
        } else {
            $instance = $assembly.CreateInstance("VideoOS.Platform.ConfigurationItems.$ItemType", $false, [System.Reflection.BindingFlags]::Default, $null, (@($serverId, $Path)), $null, $null)
            Write-Output $instance
        }
    }
}
function Export-VmsHardware {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param (
        [Parameter(ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.RecordingServer[]]
        $RecordingServer,

        [Parameter()]
        [ValidateSet('All', 'Enabled', 'Disabled')]
        [string]
        $EnableFilter = 'Enabled',

        [Parameter()]
        [string]
        $Path,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        try {
            Get-Site -ErrorAction Stop | Select-Site
        }
        catch {
            throw
        }
    }

    process {
        $version = [version](Get-ManagementServer).Version
        $supportsFillFeature = $version -ge '20.2'
        if (-not $supportsFillFeature) {
            Write-Warning "You are running version $version. Some Configuration API features introduced in 2020 R2 are unavailable, resulting in slower processing."
        }

        if ($null -eq $RecordingServer) {
            Write-Verbose 'Getting a list of all recording servers'
            $RecordingServer = Get-RecordingServer
        }

        $deviceGroupCache = @{}
        'Camera', 'Microphone', 'Speaker', 'Metadata', 'Input', 'Output' | Foreach-Object {
            $deviceType = $_
            Write-Verbose "Processing $deviceType device group hierarchy"
            Get-DeviceGroup -DeviceCategory $deviceType -Path / -Recurse | Foreach-Object {
                $group = $_
                if ($deviceType -eq 'Input') {
                    $deviceType = 'InputEvent'
                }
                if (($group."$($deviceType)Folder"."$($deviceType)s").Count -gt 0) {
                    $groupPath = Resolve-VmsDeviceGroupPath -DeviceGroup $group
                    foreach ($device in $group."$($deviceType)Folder"."$($deviceType)s") {
                        if (-not $deviceGroupCache.ContainsKey($device.Id)) {
                            $deviceGroupCache.($device.Id) = New-Object System.Collections.Generic.List[string]
                        }
                        $deviceGroupCache.($device.Id).Add($groupPath)
                    }
                }
            }
        }

        foreach ($recorder in $RecordingServer) {
            $root = $recorder | Get-ConfigurationItem
            if ($supportsFillFeature) {
                $svc = Get-IConfigurationService
                $itemTypes = 'Hardware', 'Camera', 'Microphone', 'Speaker', 'Metadata', 'InputEvent', 'Output', 'HardwareDriver', 'Storage'
                $filters = $itemTypes | Foreach-Object {
                    [VideoOS.ConfigurationApi.ClientService.ItemFilter]::new($_, @(), [VideoOS.ConfigurationApi.ClientService.EnableFilter]::$EnableFilter)
                }
                $root.Children = $svc.GetChildItemsHierarchy($recorder.Path, $itemTypes, $filters)
                $root.ChildrenFilled = $true
            } else {
                $root = $root | FillChildren -Depth 4
            }
            $storageCache = @{}
            ($root.Children | Where-Object ItemType -eq 'StorageFolder').Children | ForEach-Object { $storageCache[$_.Path] = $_ }
            $driverCache = @{}
            ($root.Children | Where-Object ItemType -eq 'HardwareDriverFolder').Children | ForEach-Object { $driverCache[$_.Path] = $_ | Get-ConfigurationItemProperty -Key Number }

            foreach ($hardware in ($root.Children | Where-Object ItemType -eq 'HardwareFolder').Children) {
                $passwordTask = $hardware | Invoke-Method -MethodId ReadPasswordHardware | Invoke-Method -MethodId ReadPasswordHardware
                $password = ($passwordTask.Properties | Where-Object Key -eq 'Password').Value
                if ($null -eq $password) {
                    Write-Warning "Failed to retrieve password for $($hardware.DisplayName)"
                }

                $row = [ordered]@{
                    Address = $hardware | Get-ConfigurationItemProperty -Key Address
                    UserName = $hardware | Get-ConfigurationItemProperty -Key UserName
                    Password = $password
                    DriverNumber = $driverCache[($hardware | Get-ConfigurationItemProperty -Key HardwareDriverPath)]
                    DriverFamily = [string]::Empty
                    StorageName = '' # Probably will want to use the first enabled camera's storage path as import only supports one path for all enabled devices
                    HardwareName = $hardware.DisplayName
                    Coordinates = ''

                    CameraName = ''
                    MicrophoneName = ''
                    SpeakerName = ''
                    MetadataName = ''
                    InputName = ''
                    OutputName = ''

                    EnabledCameraChannels = ''
                    EnabledMicrophoneChannels = ''
                    EnabledSpeakerChannels = ''
                    EnabledMetadataChannels = ''
                    EnabledInputChannels = ''
                    EnabledOutputChannels = ''

                    CameraGroup = $null
                    MicrophoneGroup = $null
                    SpeakerGroup = $null
                    MetadataGroup = $null
                    InputGroup = $null
                    OutputGroup = $null

                    RecordingServer = $root.DisplayName
                    UseDefaultCredentials = $false
                    Description = $hardware | Get-ConfigurationItemProperty -Key Description
                }

                foreach ($deviceType in 'Camera', 'Microphone', 'Speaker', 'Metadata', 'Input', 'Output') {
                    $modifiedItemTypeName = $deviceType
                    if ($deviceType -eq 'Input') {
                        $modifiedItemTypeName = 'InputEvent'
                    }
                    $devices = ($hardware.Children | Where-Object ItemType -eq "$($modifiedItemTypeName)Folder").Children | Sort-Object { [int]($_ | Get-ConfigurationItemProperty -Key Channel) }
                    if ($devices.Count -eq 0) {
                        continue
                    }
                    $row["$($deviceType)Name"] = ($devices.DisplayName | Foreach-Object { $_.Replace(';', ':')}) -join ';'
                    $enabledDevices = @()
                    foreach ($device in $devices) {
                        if ([string]::IsNullOrWhiteSpace($row.StorageName)) {
                            $row.StorageName = $storageCache.($device | Get-ConfigurationItemProperty -Key RecordingStorage).DisplayName
                        }
                        if ([string]::IsNullOrWhiteSpace($row.Coordinates)) {
                            $geocoordinate = $device | Get-ConfigurationItemProperty -Key GisPoint | ConvertFrom-GisPoint
                            $row.Coordinates = if ($geocoordinate.IsUnknown) { '' } else { $geocoordinate.ToString() }
                        }
                        $channel = $device | Get-ConfigurationItemProperty -Key Channel
                        if ($device.EnableProperty.Enabled) {
                            $enabledDevices += $channel
                        }
                        $row["$($deviceType)Group"] = $deviceGroupCache[($device.Properties | Where-Object Key -eq 'Id' | Select-Object -ExpandProperty Value)] -join ';'
                    }
                    $row["Enabled$($deviceType)Channels"] = $enabledDevices -join ';'


                }

                $obj = [pscustomobject]$row
                if (-not [string]::IsNullOrWhiteSpace($Path)) {
                    $obj | Export-Csv -Path $Path -Append -NoTypeInformation
                }
                if ([string]::IsNullOrWhiteSpace($Path) -or $PassThru) {
                    Write-Output $obj
                }
            }
        }
    }
}
function Export-VmsLicenseRequest {
    [CmdletBinding()]
    [OutputType([System.IO.FileInfo])]
    param (
        [Parameter(Mandatory)]
        [string]
        $Path,

        [Parameter()]
        [switch]
        $Force,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        $ms = Get-ManagementServer -ErrorAction Stop
        if ([Version]$ms.Version -lt '20.2') {
            throw "Management of Milestone XProtect VMS licensing using MIP SDK was introduced in version 2020 R2 (v20.2). This function is not compatible with the current Management Server version, v$($ms.Version)."
        }
    }

    process {
        try {
            $filePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
            if ((Test-Path $filePath) -and -not $Force) {
                Write-Error "File '$Path' already exists. To overwrite an existing file, specify the -Force switch."
                return
            }

            $result = $ms.LicenseInformationFolder.LicenseInformations[0].RequestLicense()
            if ($result.State -ne 'Success') {
                Write-Error "Failed to create license request. $($result.ErrorText.Trim('.'))."
                return
            }

            $content = [Convert]::FromBase64String($result.GetProperty('License'))
            [io.file]::WriteAllBytes($filePath, $content)

            if ($PassThru) {
                Get-Item -Path $filePath
            }
        }
        catch {
            Write-Error $_
        }
    }
}
function Find-ConfigurationItem {
    [CmdletBinding()]
    param (
        # Specifies all, or part of the display name of the configuration item to search for. For example, if you want to find a camera named "North West Parking" and you specify the value 'Parking', you will get results for any camera where 'Parking' appears in the name somewhere. The search is not case sensitive.
        [Parameter()]
        [string]
        $Name,

        # Specifies the type(s) of items to include in the results. The default is to include only 'Camera' items.
        [Parameter()]
        [string[]]
        $ItemType = 'Camera',

        # Specifies whether all matching items should be included, or whether only enabled, or disabled items should be included in the results. The default is to include all items regardless of state.
        [Parameter()]
        [ValidateSet('All', 'Disabled', 'Enabled')]
        [string]
        $EnableFilter = 'All',

        # An optional hashtable of additional property keys and values to filter results. Properties must be string types, and the results will be included if the property key exists, and the value contains the provided string.
        [Parameter()]
        [hashtable]
        $Properties = @{}
    )

    process {
        $svc = Get-IConfigurationService -ErrorAction Stop
        $itemFilter = [VideoOS.ConfigurationApi.ClientService.ItemFilter]::new()
        $itemFilter.EnableFilter = [VideoOS.ConfigurationApi.ClientService.EnableFilter]::$EnableFilter

        $propertyFilters = New-Object System.Collections.Generic.List[VideoOS.ConfigurationApi.ClientService.PropertyFilter]
        if (-not [string]::IsNullOrWhiteSpace($Name) -and $Name -ne '*') {
            $Properties.Name = $Name
        }
        foreach ($key in $Properties.Keys) {
            $propertyFilters.Add([VideoOS.ConfigurationApi.ClientService.PropertyFilter]::new(
                    $key,
                    [VideoOS.ConfigurationApi.ClientService.Operator]::Contains,
                    $Properties.$key
                ))
        }
        $itemFilter.PropertyFilters = $propertyFilters

        foreach ($type in $ItemType) {
            $itemFilter.ItemType = $type
            $svc.QueryItems($itemFilter, [int]::MaxValue) | Foreach-Object {
                Write-Output $_
            }
        }
    }
}

$ItemTypeArgCompleter = {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)

    ([VideoOS.ConfigurationAPI.ItemTypes] | Get-Member -Static -MemberType Property).Name | Where-Object {
        $_ -like "$wordToComplete*"
    } | Foreach-Object {
        "'$_'"
    }
}
Register-ArgumentCompleter -CommandName Find-ConfigurationItem -ParameterName ItemType -ScriptBlock $ItemTypeArgCompleter
Register-ArgumentCompleter -CommandName ConvertFrom-ConfigurationItem -ParameterName ItemType -ScriptBlock $ItemTypeArgCompleter
function Find-XProtectDevice {
    [CmdletBinding()]
    param(
        # Specifies the ItemType such as Camera, Microphone, or InputEvent. Default is 'Camera'.
        [Parameter()]
        [ValidateSet('Hardware', 'Camera', 'Microphone', 'Speaker', 'InputEvent', 'Output', 'Metadata')]
        [string[]]
        $ItemType = 'Camera',

        # Specifies name, or part of the name of the device(s) to find.
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]
        $Name,

        # Specifies all or part of the IP or hostname of the hardware device to search for.
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]
        $Address,

        # Specifies all or part of the MAC address of the hardware device to search for. Note: Searching by MAC is significantly slower than searching by IP.
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]
        $MacAddress,

        # Specifies whether all devices should be returned, or only enabled or disabled devices. Default is to return all matching devices.
        [Parameter()]
        [ValidateSet('All', 'Disabled', 'Enabled')]
        [string]
        $EnableFilter = 'All',

        # Specifies an optional hash table of key/value pairs matching properties on the items you're searching for.
        [Parameter()]
        [hashtable]
        $Properties = @{},

        [Parameter(ParameterSetName = 'ShowDialog')]
        [switch]
        $ShowDialog
    )

    begin {
        $loginSettings = Get-LoginSettings
        if ([version]'20.2' -gt [version]$loginSettings.ServerProductInfo.ProductVersion) {
            throw "The QueryItems feature was added to Milestone XProtect VMS versions starting with version 2020 R2 (v20.2). The current site is running $($loginSettings.ServerProductInfo.ProductVersion). Please upgrade to 2020 R2 or later for access to this feature."
        }
    }

    process {
        if ($ShowDialog) {
            Find-XProtectDeviceDialog
            return
        }
        if ($MyInvocation.BoundParameters.ContainsKey('Address')) {
            $ItemType = 'Hardware'
            $Properties.Address = $Address
        }

        if ($MyInvocation.BoundParameters.ContainsKey('MacAddress')) {
            $ItemType = 'Hardware'
            $MacAddress = $MacAddress.Replace(':', '').Replace('-', '')
        }
        # When many results are returned, this hashtable helps avoid unnecessary configuration api queries by caching parent items and indexing by their Path property
        $pathToItemMap = @{}

        Find-ConfigurationItem -ItemType $ItemType -EnableFilter $EnableFilter -Name $Name -Properties $Properties | Foreach-Object {
            $item = $_
            if (![string]::IsNullOrWhiteSpace($MacAddress)) {
                $hwid = ($item.Properties | Where-Object Key -eq 'Id').Value
                $mac = ((Get-ConfigurationItem -Path "HardwareDriverSettings[$hwid]").Children[0].Properties | Where-Object Key -like '*/MacAddress/*' | Select-Object -ExpandProperty Value).Replace(':', '').Replace('-', '')
                if ($mac -notlike "*$MacAddress*") {
                    return
                }
            }
            $deviceInfo = [ordered]@{}
            while ($true) {
                $deviceInfo.($item.ItemType) = $item.DisplayName
                if ($item.ItemType -eq 'RecordingServer') {
                    break
                }
                $parentItemPath = $item.ParentPath -split '/' | Select-Object -First 1

                # Set $item to the cached copy of that parent item if available. If not, retrieve it using configuration api and cache it.
                if ($pathToItemMap.ContainsKey($parentItemPath)) {
                    $item = $pathToItemMap.$parentItemPath
                } else {
                    $item = Get-ConfigurationItem -Path $parentItemPath
                    $pathToItemMap.$parentItemPath = $item
                }
            }
            [pscustomobject]$deviceInfo
        }
    }
}
function Get-ManagementServerConfig {
    [CmdletBinding()]
    param()

    begin {
        $configXml = Join-Path ([system.environment]::GetFolderPath([System.Environment+SpecialFolder]::CommonApplicationData)) 'milestone\xprotect management server\serverconfig.xml'
        if (-not (Test-Path $configXml)) {
            throw [io.filenotfoundexception]::new('Management Server configuration file not found', $configXml)
        }
    }

    process {
        $xml = [xml](Get-Content -Path $configXml)
        
        $versionNode = $xml.SelectSingleNode('/server/version')
        $clientRegistrationIdNode = $xml.SelectSingleNode('/server/ClientRegistrationId')
        $webApiPortNode = $xml.SelectSingleNode('/server/WebApiConfig/Port')
        $authServerAddressNode = $xml.SelectSingleNode('/server/WebApiConfig/AuthorizationServerUri')


        $serviceProperties = 'Name', 'PathName', 'StartName', 'ProcessId', 'StartMode', 'State', 'Status'
        $serviceInfo = Get-CimInstance -ClassName 'Win32_Service' -Property $serviceProperties -Filter "name = 'Milestone XProtect Management Server'"

        $config = @{
            Version = if ($null -ne $versionNode) { [version]::Parse($versionNode.InnerText) } else { [version]::new(0, 0) }
            ClientRegistrationId = if ($null -ne $clientRegistrationIdNode) { [guid]$clientRegistrationIdNode.InnerText } else { [guid]::Empty }
            WebApiPort = if ($null -ne $webApiPortNode) { [int]$webApiPortNode.InnerText } else { 0 }
            AuthServerAddress = if ($null -ne $authServerAddressNode) { [uri]$authServerAddressNode.InnerText } else { $null }
            ServerCertHash = $null
            InstallationPath = $serviceInfo.PathName.Trim('"')
            ServiceInfo = $serviceInfo
        }

        $netshResult = Get-ProcessOutput -FilePath 'netsh.exe' -ArgumentList "http show sslcert ipport=0.0.0.0:$($config.WebApiPort)"
        if ($netshResult.StandardOutput -match 'Certificate Hash\s+:\s+(\w+)\s+') {
            $config.ServerCertHash = $Matches.1
        }

        Write-Output ([pscustomobject]$config)
    }
}

function Get-PlaybackInfo {
    [CmdletBinding(DefaultParameterSetName = 'FromPath')]
    param (
        # Accepts a Milestone Configuration Item path string like Camera[A64740CF-5511-4957-9356-2922A25FF752]
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName = 'FromPath')]
        [ValidateScript( {
                if ($_ -notmatch '^(?<ItemType>\w+)\[(?<Id>[a-fA-F0-9\-]{36})\]$') {
                    throw "$_ does not a valid Milestone Configuration API Item path"
                }
                if ($Matches.ItemType -notin @('Camera', 'Microphone', 'Speaker', 'Metadata')) {
                    throw "$_ represents an item of type '$($Matches.ItemType)'. Only camera, microphone, speaker, or metadata item types are allowed."
                }
                return $true
            })]
        [string[]]
        $Path,

        # Accepts a Camera, Microphone, Speaker, or Metadata object
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'FromDevice')]
        [VideoOS.Platform.ConfigurationItems.IConfigurationItem[]]
        $Device,

        [Parameter()]
        [ValidateSet('MotionSequence', 'RecordingSequence', 'TimelineMotionDetected', 'TimelineRecording')]
        [string]
        $SequenceType = 'RecordingSequence',

        [Parameter()]
        [switch]
        $Parallel,

        [Parameter(ParameterSetName = 'DeprecatedParameterSet')]
        [VideoOS.Platform.ConfigurationItems.Camera]
        $Camera,

        [Parameter(ParameterSetName = 'DeprecatedParameterSet')]
        [guid]
        $CameraId,

        [Parameter(ParameterSetName = 'DeprecatedParameterSet')]
        [switch]
        $UseLocalTime
    )

    process {
        if ($PSCmdlet.ParameterSetName -eq 'DeprecatedParameterSet') {
            Write-Warning 'The Camera, CameraId, and UseLocalTime parameters are deprecated. See "Get-Help Get-PlaybackInfo -Full" for more information.'
            if ($null -ne $Camera) {
                $Path = $Camera.Path
            }
            else{
                $Path = "Camera[$CameraId]"
            }
        }
        if ($PSCmdlet.ParameterSetName -eq 'FromDevice') {
            $Path = $Device.Path
        }
        if ($Path.Count -le 60 -and $Parallel) {
            Write-Warning "Ignoring the Parallel switch since there are only $($Path.Count) devices to query."
            $Parallel = $false
        }

        if ($Parallel) {
            $jobRunner = [LocalJobRunner]::new()
        }


        $script = {
            param([string]$Path, [string]$SequenceType)
            if ($Path -notmatch '^(?<ItemType>\w+)\[(?<Id>[a-fA-F0-9\-]{36})\]$') {
                Write-Error "Path '$Path' is not a valid Milestone Configuration API item path."
                return
            }
            try {
                $site = Get-Site
                $epoch = [datetime]::SpecifyKind([datetimeoffset]::FromUnixTimeSeconds(0).DateTime, [datetimekind]::utc)
                $item = [videoos.platform.Configuration]::Instance.GetItem($site.FQID.ServerId, $Matches.Id, [VideoOS.Platform.Kind]::($Matches.ItemType))
                $sds = [VideoOS.Platform.Data.SequenceDataSource]::new($item)
                $sequenceTypeGuid = [VideoOS.Platform.Data.DataType+SequenceTypeGuids]::$SequenceType
                $first = $sds.GetData($epoch, [timespan]::zero, 0, ([datetime]::utcnow - $epoch), 1, $sequenceTypeGuid) | Select-Object -First 1
                $last = $sds.GetData([datetime]::utcnow, ([datetime]::utcnow - $epoch), 1, [timespan]::zero, 0, $sequenceTypeGuid) | Select-Object -First 1
                if ($first.EventSequence -and $last.EventSequence) {
                    [PSCustomObject]@{
                        Begin = $first.EventSequence.StartDateTime
                        End   = $last.EventSequence.EndDateTime
                        Retention = $last.EventSequence.EndDateTime - $first.EventSequence.StartDateTime
                        Path = $Path
                    }
                }
                else {
                    Write-Warning "No sequences of type '$SequenceType' found for $(($Matches.ItemType).ToLower()) $($item.Name) ($($item.FQID.ObjectId))"
                }
            } finally {
                if ($sds) {
                    $sds.Close()
                }
            }
        }

        try {
            foreach ($p in $Path) {
                if ($Parallel) {
                    $null = $jobRunner.AddJob($script, @{Path = $p; SequenceType = $SequenceType})
                }
                else {
                    $script.Invoke($p, $SequenceType) | Foreach-Object {
                        if ($UseLocalTime) {
                            $_.Begin = $_.Begin.ToLocalTime()
                            $_.End = $_.End.ToLocalTime()
                        }
                        $_
                    }
                }
            }

            if ($Parallel) {
                while ($jobRunner.HasPendingJobs()) {
                    $jobRunner.ReceiveJobs() | Foreach-Object {
                        if ($_.Output) {
                            if ($UseLocalTime) {
                                $_.Output.Begin = $_.Output.Begin.ToLocalTime()
                                $_.Output.End = $_.Output.End.ToLocalTime()
                            }
                            Write-Output $_.Output
                        }
                        if ($_.Errors) {
                            $_.Errors | Foreach-Object {
                                Write-Error $_
                            }
                        }
                    }
                    Start-Sleep -Milliseconds 200
                }
            }
        }
        finally {
            if ($jobRunner) {
                $jobRunner.Dispose()
            }
        }
    }
}
function Get-RecorderConfig {
    [CmdletBinding()]
    param()

    begin {
        $configXml = Join-Path ([system.environment]::GetFolderPath([System.Environment+SpecialFolder]::CommonApplicationData)) 'milestone\xprotect recording server\recorderconfig.xml'
        if (-not (Test-Path $configXml)) {
            throw [io.filenotfoundexception]::new('Recording Server configuration file not found', $configXml)
        }
    }

    process {
        $xml = [xml](Get-Content -Path $configXml)
        
        $versionNode = $xml.SelectSingleNode('/recorderconfig/version')
        $recorderIdNode = $xml.SelectSingleNode('/recorderconfig/recorder/id')
        $clientRegistrationIdNode = $xml.SelectSingleNode('/recorderconfig/recorder/ClientRegistrationId')
        $webServerPortNode = $xml.SelectSingleNode('/recorderconfig/webserver/port')        
        $alertServerPortNode = $xml.SelectSingleNode('/recorderconfig/driverservices/alert/port')
        $serverAddressNode = $xml.SelectSingleNode('/recorderconfig/server/address')        
        $serverPortNode = $xml.SelectSingleNode('/recorderconfig/server/webapiport')        
        $localServerPortNode = $xml.SelectSingleNode('/recorderconfig/webapi/port')
        $webApiPortNode = $xml.SelectSingleNode('/server/WebApiConfig/Port')
        $authServerAddressNode = $xml.SelectSingleNode('/recorderconfig/server/authorizationserveraddress')
        $clientCertHash = $xml.SelectSingleNode('/recorderconfig/webserver/encryption').Attributes['certificateHash'].Value

        $serviceProperties = 'Name', 'PathName', 'StartName', 'ProcessId', 'StartMode', 'State', 'Status'
        $serviceInfo = Get-CimInstance -ClassName 'Win32_Service' -Property $serviceProperties -Filter "name = 'Milestone XProtect Recording Server'"

        $config = @{
            Version = if ($null -ne $versionNode) { [version]::Parse($versionNode.InnerText) } else { [version]::new(0, 0) }
            RecorderId = if ($null -ne $recorderIdNode) { [guid]$recorderIdNode.InnerText } else { [guid]::Empty }
            ClientRegistrationId = if ($null -ne $clientRegistrationIdNode) { [guid]$clientRegistrationIdNode.InnerText } else { [guid]::Empty }
            WebServerPort = if ($null -ne $webServerPortNode) { [int]$webServerPortNode.InnerText } else { 0 }
            AlertServerPort = if ($null -ne $alertServerPortNode) { [int]$alertServerPortNode.InnerText } else { 0 }
            ServerAddress = $serverAddressNode.InnerText
            ServerPort = if ($null -ne $serverPortNode) { [int]$serverPortNode.InnerText } else { 0 }
            LocalServerPort = if ($null -ne $localServerPortNode) { [int]$localServerPortNode.InnerText } else { 0 }
            AuthServerAddress = if ($null -ne $authServerAddressNode) { [uri]$authServerAddressNode.InnerText } else { $null }
            ServerCertHash = $null
            InstallationPath = $serviceInfo.PathName.Trim('"')
            DevicePackPath = Get-ItemPropertyValue -Path HKLM:\SOFTWARE\WOW6432Node\VideoOS\DeviceDrivers -Name InstallPath
            ServiceInfo = $serviceInfo
        }

        $netshResult = Get-ProcessOutput -FilePath 'netsh.exe' -ArgumentList "http show sslcert ipport=0.0.0.0:$($config.LocalServerPort)"
        if ($netshResult.StandardOutput -match 'Certificate Hash\s+:\s+(\w+)\s+') {
            $config.ServerCertHash = $Matches.1
        }

        Write-Output ([pscustomobject]$config)
    }
}
function Get-VmsCamera {
    [CmdletBinding(DefaultParameterSetName = 'BySearch')]
    [OutputType([VideoOS.Platform.ConfigurationItems.Camera])]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName = 'ById')]
        [guid[]]
        $Id,

        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'ByHardware')]
        [VideoOS.Platform.ConfigurationItems.Hardware[]]
        $Hardware,

        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'ByHardware')]
        [ValidateRange(0, [int]::MaxValue)]
        [int[]]
        $Channel,

        [Parameter(ParameterSetName = 'ByHardware')]
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'BySearch')]
        [string]
        $Name,

        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'BySearch')]
        [string]
        $Description,

        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'BySearch')]
        [ValidateSet(
            [VideoOS.ConfigurationApi.ClientService.Operator]::Equals,
            [VideoOS.ConfigurationApi.ClientService.Operator]::NotEquals,
            [VideoOS.ConfigurationApi.ClientService.Operator]::Contains,
            [VideoOS.ConfigurationApi.ClientService.Operator]::BeginsWith
        )]
        [VideoOS.ConfigurationApi.ClientService.Operator]
        $Comparison = [VideoOS.ConfigurationApi.ClientService.Operator]::Contains,

        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'BySearch')]
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'ByHardware')]
        [VideoOS.ConfigurationApi.ClientService.EnableFilter]
        $EnableFilter = [VideoOS.ConfigurationApi.ClientService.EnableFilter]::Enabled,

        [Parameter(ParameterSetName = 'BySearch')]
        [timespan]
        $Timeout = [timespan]::FromSeconds(15),

        [Parameter(ParameterSetName = 'BySearch')]
        [ValidateRange(0, [int]::MaxValue)]
        [int]
        $MaxResults = [int]::MaxValue
    )

    begin {
        Assert-VmsConnected -ErrorAction Stop
        $aliasWarningsDelivered = 0
    }

    process {
        # TODO: Remove the alias and this block eventually.
        if ($PSCmdlet.MyInvocation.InvocationName -eq 'Get-Camera' -and $aliasWarningsDelivered -eq 0) {
            Write-Warning "The Get-Camera command is deprecated. For compatibility purposes it is temporarily aliased to Get-VmsCamera."
            if (-not $MyInvocation.BoundParameters.ContainsKey('EnableFilter')) {
                Write-Warning "The default behavior of Get-VmsCamera is to return only enabled devices, but while using the Get-Camera alias, the behavior matches the original Get-Camera command and returns all cameras, including disabled cameras."
                $EnableFilter = [VideoOS.ConfigurationApi.ClientService.EnableFilter]::All
            }
            $aliasWarningsDelivered++
        }

        switch ($PSCmdlet.ParameterSetName) {
            'ByHardware' {
                $filterChannels = $MyInvocation.BoundParameters.ContainsKey('Channel')
                foreach ($h in $Hardware) {
                    if ($EnableFilter -eq 'Enabled' -and $h.Disabled) {
                        continue
                    }
                    foreach ($camera in $h.CameraFolder.Cameras | Sort-Object Channel) {
                        if ($filterChannels -and $camera.Channel -notin $Channel) {
                            continue
                        }
                        if ($camera.Enabled -and $EnableFilter -eq 'Disabled') {
                            continue
                        }
                        if (-not $camera.Enabled -and $EnableFilter -eq 'Enabled') {
                            continue
                        }
                        if ($MyInvocation.BoundParameters.ContainsKey('Name') -and -not [string]::IsNullOrWhiteSpace($Name)) {
                            if ($camera.Name -like $Name) {
                                Write-Output $camera
                                break
                            }
                        } else {
                            Write-Output $camera
                        }
                    }
                }
            }

            'ById' {
                $site = Get-Site
                $serverId = $site.FQID.ServerId
                foreach ($guid in $Id) {
                    try {
                        $path = "Camera[$guid]"
                        $camera = [VideoOS.Platform.ConfigurationItems.Camera]::new($serverId, $path)
                        Write-Output $camera
                    } catch [VideoOS.Platform.PathNotFoundMIPException] {
                        if ($script:Messages) {
                            $message = $script:Messages.CameraOnSiteWithIdNotFound -f $site.Name, $guid
                            Write-Error -Message $message -Exception $_.Exception
                        } else {
                            Write-Error $_
                        }
                    }
                }
            }

            'BySearch' {
                Assert-VmsVersion -MinimumVersion 20.2
                $nameFilter = [VideoOS.ConfigurationApi.ClientService.PropertyFilter]::new('Name', $Comparison, $Name)
                $descriptionFilter = [VideoOS.ConfigurationApi.ClientService.PropertyFilter]::new('Description', $Comparison, $Description)
                $cameraFilter = [VideoOS.ConfigurationApi.ClientService.ItemFilter]::new('Camera', @($nameFilter, $descriptionFilter), $EnableFilter)
                $queryService = [VideoOS.ConfigurationApi.ClientService.QueryItems]::new((Get-Site).FQID.ServerId)
                foreach ($result in $queryService.Query($cameraFilter, $MaxResults)) {
                    Write-Output $result
                }
            }
            Default {}
        }
    }
}
function Get-VmsCameraGeneralSetting {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.Camera[]]
        $Camera,

        [Parameter()]
        [switch]
        $RawValues,

        [Parameter()]
        [switch]
        $ValueTypeInfo
    )

    process {
        foreach ($cam in $Camera) {
            $generalSettings = $cam.DeviceDriverSettingsFolder.DeviceDriverSettings[0].DeviceDriverSettingsChildItem
            $parsedSettings = $generalSettings | ConvertFrom-ConfigChildItem -RawValues:$RawValues
            if ($ValueTypeInfo) {
                Write-Output $parsedSettings.ValueTypeInfo.Clone()
            } else {
                Write-Output $parsedSettings.Properties.Clone()
            }
        }
    }
}
function Get-VmsCameraStream {
    [CmdletBinding(DefaultParameterSetName = 'ByName')]
    [OutputType([VmsCameraStreamConfig])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.Camera[]]
        $Camera,

        [Parameter(ParameterSetName = 'ByName')]
        [string]
        $Name,

        [Parameter(Mandatory, ParameterSetName = 'Enabled')]
        [switch]
        $Enabled,

        [Parameter(Mandatory, ParameterSetName = 'LiveDefault')]
        [switch]
        $LiveDefault,

        [Parameter(Mandatory, ParameterSetName = 'Recorded')]
        [switch]
        $Recorded,

        [Parameter()]
        [switch]
        $RawValues
    )

    process {
        foreach ($cam in $Camera) {
            $streamUsages = ($cam.StreamFolder.Streams | Select-Object -First 1).StreamUsageChildItems
            if ($null -eq $streamUsages) {
                $message = 'Camera "{0}" does not support simultaneous use of multiple streams. The following properties should be ignored for streams on this camera: DisplayName, Enabled, LiveMode, LiveDefault, Recorded.' -f $cam.Name
                Write-Warning $message
            }
            foreach ($stream in $cam.DeviceDriverSettingsFolder.DeviceDriverSettings[0].StreamChildItems) {
                $streamUsage = if ($streamUsages)  { $streamUsages | Where-Object {$_.StreamReferenceId -eq $_.StreamReferenceIdValues[$stream.DisplayName] } }
                if ($LiveDefault -and -not $streamUsage.LiveDefault) {
                    continue
                }
                if ($Recorded -and -not $streamUsage.Record) {
                    continue
                }
                if ($Enabled -and $null -eq $streamUsage -and -not $null -eq $streamUsages) {
                    # Added "-not $null -eq $streamUsages" so that old cameras without multi-stream
                    # support still provide stream settings when the user uses the Enabled switch.
                    # Otherwise it might look like there are no settings available on those camera's
                    # streams.
                    continue
                }
                if ($MyInvocation.BoundParameters.ContainsKey('Name') -and $stream.DisplayName -notlike $Name) {
                    continue
                }
                $parsedSettings = $stream | ConvertFrom-ConfigChildItem -RawValues:$RawValues
                [VmsCameraStreamConfig]@{
                    Name          = $stream.DisplayName
                    DisplayName   = $streamUsage.Name
                    Enabled       = $null -ne $streamUsage
                    LiveDefault   = $streamUsage.LiveDefault
                    LiveMode      = $streamUsage.LiveMode
                    Recorded      = $streamUsage.Record
                    Settings      = $parsedSettings.Properties.Clone()
                    ValueTypeInfo = $parsedSettings.ValueTypeInfo.Clone()
                    Camera        = $cam
                }
            }
        }
    }
}
function Get-VmsDeviceStatus {
    <#
    .SYNOPSIS
        Gets the current device status for any streaming device types directly from the recording server(s).
    .DESCRIPTION
        Recording Servers offer a status interface called
        RecorderStatusService2. This service has a method called
        GetCurrentDeviceStatus which can return the current state of any
        streaming device type including cameras, microphones, speakers, and
        metadata, as well as IO device types including inputs and outputs.

        This cmdlet will return status for one or more of the streaming device
        types, and the results will include all devices of the specified
        type(s) that are active on the recording server(s).
    .EXAMPLE
        Get-VmsDeviceStatus -DeviceType Camera, Microphone

        Returns the status of all cameras and microphones on all recording
        servers
    .EXAMPLE
        Get-VmsDeviceStatus

        Returns the status of all cameras on all recording servers. The
        default DeviceType value is 'Cameras', so if that is all you need, you
        may omit the DeviceType parameter like this.
    .EXAMPLE
        Get-RecordingServer -Name 'Recorder1' | Get-VmsDeviceStatus

        Returns the status of all cameras on recording server named
        "Recorder1".
    .EXAMPLE
        Get-VmsDeviceStatus -Parallel

        Returns the status of all cameras on all recording servers, but runs
        status requests in parallel by recording server using the PoshRSJobs
        module.
    .OUTPUTS
        [VmsStreamDeviceStatus]
    #>
    [CmdletBinding()]
    [OutputType([VmsStreamDeviceStatus])]
    param(
        # Specifies one or more Recording Server ID's. Omit this parameter and
        # all recording servers will be queried for status.
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Id')]
        [guid[]]
        $RecordingServerId,

        # Specifies one or more streaming device types to retrieve status for.
        # Default is 'Camera'.
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Camera', 'Microphone', 'Speaker', 'Metadata', IgnoreCase = $false)]
        [string[]]
        $DeviceType = 'Camera'
    )

    begin {
        $scriptBlock = {
            param([guid]$RecorderId, [VideoOS.Platform.Item[]]$Devices, [type]$VmsStreamDeviceStatusClass)
            $recorderItem = [VideoOS.Platform.Configuration]::Instance.GetItem($RecorderId, [VideoOS.Platform.Kind]::Server)
            $svc = [VideoOS.Platform.SDK.Proxy.Status2.RecorderStatusService2]::new($recorderItem.FQID.ServerId.Uri)
            $status = @{}
            $currentStatus = $svc.GetCurrentDeviceStatus((Get-Token), $Devices.FQID.ObjectId)
            foreach ($kind in 'Camera', 'Microphone', 'Speaker', 'Metadata') {
                foreach ($entry in $currentStatus."$($kind)DeviceStatusArray") {
                    $status[$entry.DeviceId] = $entry
                }
            }
            foreach ($item in $Devices) {
                $obj = $VmsStreamDeviceStatusClass::new($status[$item.FQID.ObjectId])
                $obj.DeviceName = $item.Name
                $obj.DeviceType = [VideoOS.Platform.Kind]::DefaultTypeToNameTable[$item.FQID.Kind]
                $obj.RecorderName = $recorderItem.Name
                $obj.RecorderId = $RecorderItem.FQID.ObjectId
                Write-Output $obj
            }
        }
    }

    process {
        <# TODO: Once a decision is made on how to handle the PoshRSJob
           dependency, uncomment the bits below and remove the line right
           after the opening foreach curly brace as it's already handled
           in the else block.
        #>
        $recorderCameraMap = Get-DevicesByRecorder -Id $RecordingServerId -DeviceType $DeviceType
        # $jobs = [system.collections.generic.list[RSJob]]::new()
        foreach ($recorderId in $recorderCameraMap.Keys) {
            $scriptBlock.Invoke($recorderId, $recorderCameraMap.$recorderId, ([VmsStreamDeviceStatus]))
            # if ($Parallel -and $RecordingServerId.Count -gt 1) {
            #     $job = Start-RSJob -ScriptBlock $scriptBlock -ArgumentList $recorderId, $recorderCameraMap.$recorderId, ([VmsStreamDeviceStatus])
            #     $jobs.Add($job)
            # } else {
            #     $scriptBlock.Invoke($recorderId, $recorderCameraMap.$recorderId, ([VmsStreamDeviceStatus]))
            # }
        }
        # if ($jobs.Count -gt 0) {
        #     $jobs | Wait-RSJob -ShowProgress:($ProgressPreference -eq 'Continue') | Receive-RSJob
        #     $jobs | Remove-RSJob
        # }
    }
}
function Get-VmsFailoverGroup {
    [CmdletBinding()]
    [OutputType([VideoOS.Platform.ConfigurationItems.FailoverGroup])]
    param(
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [guid]
        $Id
    )

    begin {
        Assert-VmsVersion -MinimumVersion 21.2 -ErrorAction Stop
    }

    process {
        if ($Id) {
            $group = [VideoOS.Platform.ConfigurationItems.FailoverGroup]::new((Get-Site).FQID.ServerId, "FailoverGroup[$Id]")
            Write-Output $group
        }
        else {
            $ms = Get-ManagementServer
            Write-Output $ms.FailoverGroupFolder.FailoverGroups
        }
    }
}
function Get-VmsStorageRetention {
    [CmdletBinding()]
    [OutputType([timespan])]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.Storage]
        $Storage
    )

    process {
        $retention = [int]$Storage.RetainMinutes
        foreach ($archive in $Storage.ArchiveStorageFolder.ArchiveStorages) {
            if ($archive.RetainMinutes -gt $retention) {
                $retention = $archive.RetainMinutes
            }
        }
        Write-Output ([timespan]::FromMinutes($retention))
    }
}
function Import-VmsHardware {
    [CmdletBinding(DefaultParameterSetName = 'ImportHardware', SupportsShouldProcess)]
    param (
        [Parameter(ValueFromPipeline, ParameterSetName = 'ImportHardware')]
        [VideoOS.Platform.ConfigurationItems.RecordingServer]
        $RecordingServer,

        [Parameter(Mandatory, ParameterSetName = 'ImportHardware')]
        [Parameter(Mandatory, ParameterSetName = 'SaveTemplate')]
        [string]
        $Path,

        [Parameter(Mandatory, ParameterSetName = 'SaveTemplate')]
        [switch]
        $SaveTemplate,

        [Parameter(ParameterSetName = 'SaveTemplate')]
        [switch]
        $Minimal
    )

    process {
        if ($SaveTemplate) {
            if (Test-Path -Path $Path) {
                Write-Error "There is already a file at $Path"
                return
            }
            if ($Minimal) {
                $rows = @(
                    [pscustomobject]@{
                        Address = 'http://192.168.1.100'
                        UserName = 'root'
                        Password = 'pass'
                        Description = 'This description column is optional and only present in this template as a place give you additional information. Feel free to delete this column.'
                    },
                    [pscustomobject]@{
                        Address = 'https://camera2.milestone.internal'
                        UserName = 'root'
                        Password = 'pass'
                        Description = 'In this row we use HTTPS instead of HTTP. If your camera is setup for HTTPS, this is how you tell Milestone to add the camera over a secure connection.'
                    },
                    [pscustomobject]@{
                        Address = '192.168.1.101'
                        UserName = 'root'
                        Password = 'pass'
                        Description = 'In this row we provide only the IP address. When you do this, the server assumes you want to use HTTP port 80.'
                    }
                )

                $rows | Export-Csv -Path $Path -NoTypeInformation
            }
            else {
                $rows = @(
                    [pscustomobject]@{
                        Address = 'http://192.168.1.100'
                        UserName = 'root'
                        Password = 'pass'
                        UserName2 = 'admin'
                        Password2 = 'admin'
                        UserName3 = 'service'
                        Password3 = '123456'
                        DriverNumber = ''
                        DriverFamily = ''
                        StorageName = ''
                        HardwareName = ''
                        Coordinates = ''

                        CameraName = ''
                        MicrophoneName = ''
                        SpeakerName = ''
                        MetadataName = ''
                        InputName = ''
                        OutputName = ''

                        EnabledCameraChannels = ''
                        EnabledMicrophoneChannels = ''
                        EnabledSpeakerChannels = ''
                        EnabledMetadataChannels = ''
                        EnabledInputChannels = ''
                        EnabledOutputChannels = ''

                        CameraGroup = $null
                        MicrophoneGroup = $null
                        SpeakerGroup = $null
                        MetadataGroup = $null
                        InputGroup = $null
                        OutputGroup = $null

                        RecordingServer = $null
                        UseDefaultCredentials = $false
                        Description = 'This camera will be scanned using three different sets of credentials and since no driver, or driver family name was provided, the camera will be scanned using all drivers. All values are left empty so only the first camera channel will be enabled, and the hardware name and child device names will be left at defaults.'
                    },
                    [pscustomobject]@{
                        Address = 'https://192.168.1.101'
                        UserName = 'root'
                        Password = 'pass'
                        UserName2 = 'admin'
                        Password2 = 'admin'
                        UserName3 = 'service'
                        Password3 = '123456'
                        DriverNumber = ''
                        DriverFamily = ''
                        StorageName = '90 Day Storage'
                        HardwareName = ''
                        Coordinates = ''

                        CameraName = ''
                        MicrophoneName = ''
                        SpeakerName = ''
                        MetadataName = ''
                        InputName = ''
                        OutputName = ''

                        EnabledCameraChannels = ''
                        EnabledMicrophoneChannels = ''
                        EnabledSpeakerChannels = ''
                        EnabledMetadataChannels = ''
                        EnabledInputChannels = ''
                        EnabledOutputChannels = ''

                        CameraGroup = $null
                        MicrophoneGroup = $null
                        SpeakerGroup = $null
                        MetadataGroup = $null
                        InputGroup = $null
                        OutputGroup = $null

                        RecordingServer = $null
                        UseDefaultCredentials = $false
                        Description = 'Same as the previous row, except this camera will be added over HTTPS port 443 and recorded to a storage configuration named "90 Day Storage"'
                    },
                    [pscustomobject]@{
                        Address = 'https://camera3.milestone.internal:8443'
                        UserName = 'root'
                        Password = 'pass'
                        UserName2 = 'admin'
                        Password2 = 'admin'
                        UserName3 = 'service'
                        Password3 = '123456'
                        DriverNumber = ''
                        DriverFamily = ''
                        StorageName = 'Invalid Storage Name'
                        HardwareName = ''
                        Coordinates = ''

                        CameraName = ''
                        MicrophoneName = ''
                        SpeakerName = ''
                        MetadataName = ''
                        InputName = ''
                        OutputName = ''

                        EnabledCameraChannels = ''
                        EnabledMicrophoneChannels = ''
                        EnabledSpeakerChannels = ''
                        EnabledMetadataChannels = ''
                        EnabledInputChannels = ''
                        EnabledOutputChannels = ''

                        CameraGroup = $null
                        MicrophoneGroup = $null
                        SpeakerGroup = $null
                        MetadataGroup = $null
                        InputGroup = $null
                        OutputGroup = $null

                        RecordingServer = $null
                        UseDefaultCredentials = $false
                        Description = 'Same as the previous row, except this camera will be added over HTTPS port 8443 and a dns name is used instead of an IP address. Since the StorageName value doesn''t match the name of a storage configuration on the recording server, the camera will record to default storage.'
                    },
                    [pscustomobject]@{
                        Address = '192.168.1.103'
                        UserName = 'root'
                        Password = 'pass'
                        UserName2 = ''
                        Password2 = ''
                        UserName3 = ''
                        Password3 = ''
                        DriverNumber = '713;710'
                        DriverFamily = ''
                        StorageName = ''
                        HardwareName = ''
                        Coordinates = ''

                        CameraName = ''
                        MicrophoneName = ''
                        SpeakerName = ''
                        MetadataName = ''
                        InputName = ''
                        OutputName = ''

                        EnabledCameraChannels = ''
                        EnabledMicrophoneChannels = ''
                        EnabledSpeakerChannels = ''
                        EnabledMetadataChannels = ''
                        EnabledInputChannels = ''
                        EnabledOutputChannels = ''

                        CameraGroup = $null
                        MicrophoneGroup = $null
                        SpeakerGroup = $null
                        MetadataGroup = $null
                        InputGroup = $null
                        OutputGroup = $null

                        RecordingServer = $null
                        UseDefaultCredentials = $true
                        Description = 'This camera will be added using HTTP port 80 since the address was not provided in the form of a URI. Since two DriverNumber values are present, the camera will be scanned against two drivers to see which one is best for the camera. Also since UseDefaultCredentials is true, the driver default credentials will be tried in addition to the user-supplied credentials.'
                    },
                    [pscustomobject]@{
                        Address = '192.168.1.104'
                        UserName = 'root'
                        Password = 'pass'
                        UserName2 = ''
                        Password2 = ''
                        UserName3 = ''
                        Password3 = ''
                        DriverNumber = ''
                        DriverFamily = 'Axis;Bosch'
                        StorageName = ''
                        HardwareName = ''
                        Coordinates = ''

                        CameraName = ''
                        MicrophoneName = ''
                        SpeakerName = ''
                        MetadataName = ''
                        InputName = ''
                        OutputName = ''

                        EnabledCameraChannels = ''
                        EnabledMicrophoneChannels = ''
                        EnabledSpeakerChannels = ''
                        EnabledMetadataChannels = ''
                        EnabledInputChannels = ''
                        EnabledOutputChannels = ''

                        CameraGroup = $null
                        MicrophoneGroup = $null
                        SpeakerGroup = $null
                        MetadataGroup = $null
                        InputGroup = $null
                        OutputGroup = $null

                        RecordingServer = $null
                        UseDefaultCredentials = $false
                        Description = 'This camera will be scanned against all Axis and Bosch device driver to find the best match.'
                    },
                    [pscustomobject]@{
                        Address = '192.168.1.105'
                        UserName = 'root'
                        Password = 'pass'
                        UserName2 = ''
                        Password2 = ''
                        UserName3 = ''
                        Password3 = ''
                        DriverNumber = '5000'
                        DriverFamily = ''
                        StorageName = ''
                        HardwareName = 'Parking (192.168.1.105)'
                        Coordinates = ''

                        CameraName = 'Parking East;Parking West'
                        MicrophoneName = ''
                        SpeakerName = ''
                        MetadataName = ''
                        InputName = ''
                        OutputName = ''

                        EnabledCameraChannels = '0;1'
                        EnabledMicrophoneChannels = ''
                        EnabledSpeakerChannels = ''
                        EnabledMetadataChannels = ''
                        EnabledInputChannels = ''
                        EnabledOutputChannels = ''

                        CameraGroup = $null
                        MicrophoneGroup = $null
                        SpeakerGroup = $null
                        MetadataGroup = $null
                        InputGroup = $null
                        OutputGroup = $null

                        RecordingServer = $null
                        UseDefaultCredentials = $false
                        Description = 'This camera will be added using the StableFPS driver using driver ID 5000, and the first two camera channels will be enabled. The hardware and two first camera channels will have user-supplied names from the CSV while the rest of the devices will have default names and will remain disabled.'
                    },
                    [pscustomobject]@{
                        Address = '192.168.1.106'
                        UserName = 'root'
                        Password = 'pass'
                        UserName2 = ''
                        Password2 = ''
                        UserName3 = ''
                        Password3 = ''
                        DriverNumber = '5000'
                        DriverFamily = ''
                        StorageName = ''
                        HardwareName = 'Reception'
                        Coordinates = ''

                        CameraName = 'Reception - Front Desk'
                        MicrophoneName = ''
                        SpeakerName = ''
                        MetadataName = ''
                        InputName = ''
                        OutputName = ''

                        EnabledCameraChannels = 'All'
                        EnabledMicrophoneChannels = ''
                        EnabledSpeakerChannels = ''
                        EnabledMetadataChannels = ''
                        EnabledInputChannels = ''
                        EnabledOutputChannels = ''

                        CameraGroup = '/Main Office/Reception'
                        MicrophoneGroup = $null
                        SpeakerGroup = $null
                        MetadataGroup = $null
                        InputGroup = $null
                        OutputGroup = $null

                        RecordingServer = $null
                        UseDefaultCredentials = $false
                        Description = 'This camera will be added with all camera channels enabled, and all other child devices will be disabled. The first camera channel will have a custom name and any additional channels will have the default name. A camera group path was provided, so if the top level group "Main Office" or subgroup "Reception" do not exist, they will be created, and the enabled cameras will be placed into the Reception subgroup.'
                    },
                    [pscustomobject]@{
                        Address = '192.168.1.107'
                        UserName = 'root'
                        Password = 'pass'
                        UserName2 = ''
                        Password2 = ''
                        UserName3 = ''
                        Password3 = ''
                        DriverNumber = '5000'
                        DriverFamily = ''
                        StorageName = ''
                        HardwareName = 'Warehouse (192.168.1.107)'
                        Coordinates = ''

                        CameraName = 'Warehouse Overview;;Warehouse 180'
                        MicrophoneName = ''
                        SpeakerName = ''
                        MetadataName = ''
                        InputName = ''
                        OutputName = ''

                        EnabledCameraChannels = '0;2'
                        EnabledMicrophoneChannels = ''
                        EnabledSpeakerChannels = ''
                        EnabledMetadataChannels = ''
                        EnabledInputChannels = ''
                        EnabledOutputChannels = ''

                        CameraGroup = '/New cameras'
                        MicrophoneGroup = $null
                        SpeakerGroup = $null
                        MetadataGroup = $null
                        InputGroup = $null
                        OutputGroup = $null

                        RecordingServer = $null
                        UseDefaultCredentials = $false
                        Description = 'This camera will be added with the first and third channels enabled. Channels are counted from 0, so channel 0 represents "Camera 1". Note how in the CameraName column if you split by the semicolon symbol, the second entry is empty. This means the second camera channel will not be renamed from the default, but the first and third channels will be.'
                    },
                    [pscustomobject]@{
                        Address = '192.168.1.107'
                        UserName = 'root'
                        Password = 'pass'
                        UserName2 = ''
                        Password2 = ''
                        UserName3 = ''
                        Password3 = ''
                        DriverNumber = '5000'
                        DriverFamily = ''
                        StorageName = ''
                        HardwareName = 'Warehouse (192.168.1.107)'
                        Coordinates = '47.620, -122.349'

                        CameraName = ''
                        MicrophoneName = ''
                        SpeakerName = ''
                        MetadataName = ''
                        InputName = ''
                        OutputName = ''

                        EnabledCameraChannels = ''
                        EnabledMicrophoneChannels = ''
                        EnabledSpeakerChannels = ''
                        EnabledMetadataChannels = ''
                        EnabledInputChannels = ''
                        EnabledOutputChannels = ''

                        CameraGroup = '/New cameras'
                        MicrophoneGroup = $null
                        SpeakerGroup = $null
                        MetadataGroup = $null
                        InputGroup = $null
                        OutputGroup = $null

                        RecordingServer = 'Recorder-10'
                        UseDefaultCredentials = $false
                        Description = 'This row has a recording server display name specified which will override the recording server specified in the RecordingServer parameter of Import-VmsHardware. It also has GPS coordinates in lat, long format so all enabled devices will have the GisPoint property updated.'
                    }
                )
                $rows | Export-Csv -Path $Path -NoTypeInformation
            }
            return
        }
        try {
            Get-Site -ErrorAction Stop | Select-Site
        }
        catch {
            throw
        }
        $initialProgressPreference = $ProgressPreference
        try {
            $ProgressPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
            $tasks = New-Object System.Collections.Generic.List[object]
            $rows = Import-Csv -Path $Path
            $recorderCache = @{}
            $recorderPathMap = @{}
            $addressRowMap = @{}
            Get-RecordingServer | Foreach-Object { $recorderCache.($_.Name) = $_; $recorderPathMap.($_.Path) = $_ }
            foreach ($row in $rows) {
                $recorder = $RecordingServer
                if (-not [string]::IsNullOrWhiteSpace($row.RecordingServer)) {
                    if (-not $recorderCache.ContainsKey($row.RecordingServer)) {
                        Write-Error "Recording Server with display name '$($row.RecordingServer)' not found. Entry '$($row.HardwareName)' with address $($row.Address) will be skipped."
                        continue
                    }
                    $recorder = $recorderCache.($row.RecordingServer)
                }
                $credentials = New-Object System.Collections.Generic.List[pscredential]
                if (-not [string]::IsNullOrWhiteSpace($row.UserName) -and -not [string]::IsNullOrWhiteSpace($row.Password)) {
                    $credentials.Add([pscredential]::new($row.UserName, ($row.Password | ConvertTo-SecureString -AsPlainText -Force)))
                }
                if (-not [string]::IsNullOrWhiteSpace($row.UserName2) -and -not [string]::IsNullOrWhiteSpace($row.Password2)) {
                    $credentials.Add([pscredential]::new($row.UserName2, ($row.Password2 | ConvertTo-SecureString -AsPlainText -Force)))
                }
                if (-not [string]::IsNullOrWhiteSpace($row.UserName3) -and -not [string]::IsNullOrWhiteSpace($row.Password3)) {
                    $credentials.Add([pscredential]::new($row.UserName3, ($row.Password3 | ConvertTo-SecureString -AsPlainText -Force)))
                }
                $scanParams = @{
                    RecordingServer = $recorder
                    Address = $row.Address
                    Credential = $credentials
                    DriverNumber = $row.DriverNumber -split ';' | Where-Object { ![string]::IsNullOrWhiteSpace($_) }
                    DriverFamily = $row.DriverFamily -split ';' | Where-Object { ![string]::IsNullOrWhiteSpace($_) }
                    UseDefaultCredentials = if ($credentials.Count -eq 0) { 'True' } else { $row.UseDefaultCredentials -eq 'True' }
                    PassThru = $true
                }
                if ($PSCmdlet.ShouldProcess("$($scanParams.Address) from recording server $($scanParams.RecordingServer.Name)", "Running Start-VmsHardwareScan")) {
                    Start-VmsHardwareScan @scanParams | Foreach-Object {
                        $tasks.Add($_)
                        $uri = $row.Address -as [uri]
                        $ip = $row.Address -as [ipaddress]
                        if ($null -ne $ip) {
                            $uri = $ip | ConvertTo-Uri
                        }
                        $addressRowMap.($uri.ToString()) = $row
                    }
                }
            }
            $ProgressPreference = $initialProgressPreference

            $scans = New-Object System.Collections.Generic.List[VmsHardwareScanResult]
            if ($tasks.Count -gt 0) {
                Wait-VmsTask -Path $tasks.Path -Title "Scanning hardware" -Cleanup | Foreach-Object {
                    $results = if ($_.Children.Count -gt 0) { [VmsHardwareScanResult[]]$_.Children } else { [VmsHardwareScanResult]$_ }
                    foreach ($result in $results) {
                        $result.RecordingServer = $recorderPathMap.($_.ParentPath)
                        $scans.Add($result)
                    }
                }
            }

            $deviceGroupCache = @{
                Camera = @{}
                Microphone = @{}
                Speaker = @{}
                Input = @{}
                Output = @{}
                Metadata = @{}
            }

            $recorderByPath = @{}
            $storageCache = @{}
            if ($PSCmdlet.ShouldProcess("Milestone XProtect site '$((Get-Site).Name)'", "Add-VmsHardware")) {
                Add-VmsHardware -HardwareScan $scans -Force -SkipConfig | Foreach-Object {
                    $stopwatch = [diagnostics.stopwatch]::StartNew()
                    $hardware = $_
                    $hardware.Enabled = $true

                    $row = $addressRowMap.($hardware.Address)

                    # If user is assigning to non-default storage, then we need to discover the available storages
                    # on the recording server and cache the information so we don't query the same information repeatedly.
                    $recorder = $null
                    $storagePath = ''
                    if (-not [string]::IsNullOrWhiteSpace($row.StorageName)) {
                        if (-not $recorderByPath.ContainsKey($hardware.ParentItemPath)) {
                            $recorderByPath.($hardware.ParentItemPath) = Get-RecordingServer -Id ($hardware.ParentItemPath.Substring(16, 36))
                            $storageCache.($hardware.ParentItemPath) = @{}
                        }
                        $recorder = $recorderByPath.($hardware.ParentItemPath)
                        if (-not $storageCache.($hardware.ParentItemPath).ContainsKey($row.StorageName)) {
                            foreach ($storage in $recorder.StorageFolder.Storages) {
                                $storageCache.($hardware.ParentItemPath).($storage.Name) = $storage.Path
                            }
                        }
                        $storagePath = $storageCache.($hardware.ParentItemPath).($row.StorageName)
                        if ([string]::IsNullOrWhiteSpace($storagePath)) {
                            $storagePath = [string]::Empty
                            Write-Warning "Storage named '$($row.StorageName)' not found on Recording Server '$($recorder.Name)'. All recording devices on $($hardware.Name) will record to the default storage."
                        }
                    }

                    if (-not [string]::IsNullOrWhiteSpace($row.HardwareName)) {
                        $hardware.Name = $row.HardwareName
                    }
                    if ([string]::IsNullOrWhiteSpace($row.Description)) {
                        $hardware.Description = "Added using PowerShell at $(Get-Date)"
                    }
                    elseif ($row.Description -ne 'blank') {
                        $hardware.Description = $row.Description
                    }
                    $hardware.Save()

                    $enabledChannels = @{}
                    foreach ($deviceType in @('Camera', 'Microphone', 'Speaker', 'Input', 'Output', 'Metadata')) {
                        if ([string]::IsNullOrWhiteSpace($row."Enabled$($deviceType)Channels")) {
                            $enabledChannels[$deviceType] = @()
                        }
                        elseif ($row."Enabled$($deviceType)Channels" -eq 'All') {
                            $enabledChannels[$deviceType] = 0..511
                        }
                        else {
                            $enabledChannels[$deviceType] = @( $row."Enabled$($deviceType)Channels" -split ';' | Where-Object { ![string]::IsNullOrWhiteSpace($_) } | ForEach-Object { [int]($_.Trim())} )
                        }
                    }
                    if ($enabledChannels['Camera'].Count -eq 0) { $enabledChannels['Camera'] = @(0) }


                    $deviceNames = @{
                        Camera = @{}
                        Microphone = @{}
                        Speaker = @{}
                        Input = @{}
                        Output = @{}
                        Metadata = @{}
                    }
                    foreach ($key in $deviceNames.Keys) {
                        if (-not [string]::IsNullOrWhiteSpace($row."$($key)Name")) {
                            $names = @( $row."$($key)Name" -split ';' | ForEach-Object { $_.Trim() } )
                            for ($i = 0; $i -lt $names.Count; $i++) {
                                $deviceNames[$key][$i] = $names[$i]
                            }
                        }
                    }

                    $gisPoint = 'POINT EMPTY'
                    if (-not [string]::IsNullOrWhiteSpace($row.Coordinates)) {
                        $lat, $long = $row.Coordinates -split ',' | ForEach-Object { $_.Trim() -as [double] }
                        if ($null -ne $lat -and $null -ne $long) {
                            $gisPoint = "POINT ($long $lat)"
                        }
                    }


                    $hardware | Get-VmsCamera  | Set-NewDeviceConfig -HardwareName $hardware.Name -EnabledChannels $enabledChannels.Camera     -ChannelNames $deviceNames.Camera     -DeviceGroups $row.CameraGroup     -DeviceGroupCache $deviceGroupCache -StoragePath $storagePath -GisPoint $gisPoint
                    $hardware | Get-Microphone | Set-NewDeviceConfig -HardwareName $hardware.Name -EnabledChannels $enabledChannels.Microphone -ChannelNames $deviceNames.Microphone -DeviceGroups $row.MicrophoneGroup -DeviceGroupCache $deviceGroupCache -StoragePath $storagePath -GisPoint $gisPoint
                    $hardware | Get-Speaker    | Set-NewDeviceConfig -HardwareName $hardware.Name -EnabledChannels $enabledChannels.Speaker    -ChannelNames $deviceNames.Speaker    -DeviceGroups $row.SpeakerGroup    -DeviceGroupCache $deviceGroupCache -StoragePath $storagePath -GisPoint $gisPoint
                    $hardware | Get-Input      | Set-NewDeviceConfig -HardwareName $hardware.Name -EnabledChannels $enabledChannels.Input      -ChannelNames $deviceNames.Input      -DeviceGroups $row.InputGroup      -DeviceGroupCache $deviceGroupCache -StoragePath $storagePath -GisPoint $gisPoint
                    $hardware | Get-Output     | Set-NewDeviceConfig -HardwareName $hardware.Name -EnabledChannels $enabledChannels.Output     -ChannelNames $deviceNames.Output     -DeviceGroups $row.OutputGroup     -DeviceGroupCache $deviceGroupCache -StoragePath $storagePath -GisPoint $gisPoint
                    $hardware | Get-Metadata   | Set-NewDeviceConfig -HardwareName $hardware.Name -EnabledChannels $enabledChannels.Metadata   -ChannelNames $deviceNames.Metadata   -DeviceGroups $row.MetadataGroup   -DeviceGroupCache $deviceGroupCache -StoragePath $storagePath -GisPoint $gisPoint
                    $hardware.ClearChildrenCache()
                    Write-Verbose "Completed configuration of $($hardware.Name) ($($hardware.Address)) in $($stopwatch.ElapsedMilliseconds)ms"
                    $hardware
                }
            }
        }
        finally {
            $ProgressPreference = $initialProgressPreference
        }
    }
}

function Set-NewDeviceConfig {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [object]
        $Device,

        [Parameter(Mandatory)]
        [string]
        $HardwareName,

        [Parameter()]
        [int[]]
        $EnabledChannels,

        [Parameter(Mandatory)]
        [hashtable]
        $ChannelNames,

        # Semi-colon delimited list of device group paths
        [Parameter()]
        [string]
        $DeviceGroups,

        [Parameter()]
        [hashtable]
        $DeviceGroupCache,

        # Specifies the Configuration API path like Storage[guid], not disk path.
        [Parameter()]
        [string]
        $StoragePath,

        [Parameter()]
        [string]
        $GisPoint = 'POINT EMPTY'
    )

    process {
        try {
            $deviceType = ''
            if ($Device.Path -match '(?<itemtype>.+)\[[a-fA-F0-9\-]{36}\](?:/(?<folderType>.+))?') {
                $deviceType = $Matches.itemtype
                if ($deviceType -eq 'InputEvent') {
                    $deviceType = 'Input'
                }
            }
            else {
                Write-Error "Failed to parse item type from configuration api path '$($Device.Path)'"
                return
            }

            if ([string]::IsNullOrWhiteSpace($ChannelNames[$Device.Channel])) {
                $Device.Name = $HardwareName + " - $deviceType $($Device.Channel + 1)"
            }
            else {
                $Device.Name = $ChannelNames[$Device.Channel]
            }

            if ($Device.Channel -in $EnabledChannels[$Device.Channel]) {
                $Device.Enabled = $true
                if (-not [string]::IsNullOrWhiteSpace($DeviceGroups)) {
                    foreach ($groupName in @( $DeviceGroups -split ';' )) {
                        if (-not $DeviceGroupCache.$deviceType.ContainsKey($groupName)) {
                            $DeviceGroupCache.$deviceType.$groupName = Add-DeviceGroup -DeviceCategory $deviceType -Path $groupName
                        }
                        $DeviceGroupCache.$deviceType.$groupName | Add-DeviceGroupMember -DeviceCategory $deviceType -Device $Device
                    }
                }
            }

            $device.GisPoint = $GisPoint
            $Device.Save()
        }
        catch [VideoOS.Platform.Proxy.ConfigApi.ValidateResultException] {
            foreach ($errorResult in $_.Exception.ValidateResult.ErrorResults) {
                Write-Error "Failed to update settings on $($Device.Name) ($($Device.Id)) due to a $($errorResult.ErrorProperty) validation error. $($errorResult.ErrorText.Trim('.'))."
            }
        }

        try {
            if (-not [string]::IsNullOrWhiteSpace($StoragePath) -and $null -ne $Device.RecordingStorage) {
                if ($Device.RecordingStorage -ne $StoragePath) {
                    $moveData = $false
                    $null = $Device.ChangeDeviceRecordingStorage($StoragePath, $moveData) | Wait-VmsTask -Cleanup
                }
            }
        }
        catch [VideoOS.Platform.ServerFaultMIPException] {
            $errorText = $_.Exception.InnerException.Message
            Write-Error "Failed to update recording storage for $($Device.Name) ($($Device.Id). $($errorText.Trim('.'))."
        }
    }
}
function Import-VmsLicense {
    [CmdletBinding()]
    [OutputType([VideoOS.Platform.ConfigurationItems.LicenseInformation])]
    param (
        [Parameter(Mandatory)]
        [string]
        $Path
    )

    begin {
        $ms = Get-ManagementServer -ErrorAction Stop
        if ([Version]$ms.Version -lt '20.2') {
            throw "Management of Milestone XProtect VMS licensing using MIP SDK was introduced in version 2020 R2 (v20.2). This function is not compatible with the current Management Server version, v$($ms.Version)."
        }
    }

    process {
        try {
            $filePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
            if (-not (Test-Path $filePath)) {
                throw [System.IO.FileNotFoundException]::new('Import-VmsLicense could not find the file.', $filePath)
            }
            $bytes = [IO.File]::ReadAllBytes($filePath)
            $b64 = [Convert]::ToBase64String($bytes)
            $result = $ms.LicenseInformationFolder.LicenseInformations[0].UpdateLicense($b64)
            if ($result.State -eq 'Success') {
                $ms.ClearChildrenCache()
                Write-Output $ms.LicenseInformationFolder.LicenseInformations[0]
            }
            else {
                Write-Error "Failed to import updated license file. $($result.ErrorText.Trim('.'))."
            }
        }
        catch {
            Write-Error -Message $_.Message -Exception $_.Exception
        }
    }
}
function Invoke-VmsLicenseActivation {
    [CmdletBinding()]
    [OutputType([VideoOS.Platform.ConfigurationItems.LicenseInformation])]
    param (
        [Parameter(Mandatory)]
        [pscredential]
        $Credential,

        [Parameter()]
        [switch]
        $EnableAutoActivation
    )

    begin {
        $ms = Get-ManagementServer -ErrorAction Stop
        if ([Version]$ms.Version -lt '20.2') {
            throw "Management of Milestone XProtect VMS licensing using MIP SDK was introduced in version 2020 R2 (v20.2). This function is not compatible with the current Management Server version, v$($ms.Version)."
        }
    }

    process {
        try {
            $result = $ms.LicenseInformationFolder.LicenseInformations[0].ActivateLicense($Credential.UserName, $Credential.Password, $EnableAutoActivation) | Wait-VmsTask -Title 'Performing online license activation' -Cleanup
            $state = ($result.Properties | Where-Object Key -eq 'State').Value
            if ($state -eq 'Success') {
                $ms.ClearChildrenCache()
                Write-Output $ms.LicenseInformationFolder.LicenseInformations[0]
            }
            else {
                $errorText = ($result.Properties | Where-Object Key -eq 'ErrorText').Value
                if ([string]::IsNullOrWhiteSpace($errorText)) {
                    $errorText = "Unknown error."
                }
                Write-Error "Call to ActivateLicense failed. $($errorText.Trim('.'))."
            }
        }
        catch {
            Write-Error -Message $_.Message -Exception $_.Exception
        }
    }
}
function Remove-Hardware {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact='High')]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.Hardware[]]
        $Hardware
    )

    process {
        try {
            $folders = @{}
            $Hardware | ForEach-Object {
                if (-not $folders.ContainsKey($_.ParentPath)) {
                    $folders[$_.ParentPath] = [VideoOS.Platform.ConfigurationItems.HardwareFolder]::new((Get-Site).FQID.ServerId, $_.ParentPath)
                }
            }
            $action = 'Permanently delete hardware and all associated video, audio and metadata from the VMS'
            foreach ($hw in $Hardware) {
                try {
                    $target = "$($hw.Name) with ID $($hw.Id)"
                    if ($PSCmdlet.ShouldProcess($target, $action)) {
                        $folder = $folders[$hw.ParentPath]
                        $result = $folder.DeleteHardware($hw.Path) | Wait-VmsTask -Title "Removing hardware $($hw.Name)" -Cleanup
                        $properties = @{}
                        $result.Properties | Foreach-Object { $properties[$_.Key] = $_.Value}
                        if ($properties.State -ne 'Success') {
                            Write-Error "An error occurred while deleting the hardware. $($properties.ErrorText.Trim('.'))."
                        }
                    }
                }
                catch [VideoOS.Platform.PathNotFoundMIPException] {
                    Write-Error "The hardware named $($hw.Name) with ID $($hw.Id) was not found."
                }
            }
        }
        catch [VideoOS.Platform.PathNotFoundMIPException] {
            Write-Error "One or more recording servers for the provided hardware values do not exist."
        }
    }
}
function Remove-VmsFailoverGroup {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact="High")]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.FailoverGroup]
        $FailoverGroup
    )

    begin {
        Assert-VmsVersion -MinimumVersion 21.2 -ErrorAction Stop
    }

    process {
        if ($PSCmdlet.ShouldProcess($FailoverGroup.Name, "Remove failover group")) {
            $ms = Get-ManagementServer
            $task = $ms.FailoverGroupFolder.RemoveFailoverGroup($FailoverGroup.Path)
            if ($task.State -ne 'Success') {
                Write-Error "Remove-VmsFailoverGroup encounted an error. $($task.ErrorText.Trim('.'))."
            }
        }
    }
}
function Resolve-VmsDeviceGroupPath {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.IConfigurationItem]
        $DeviceGroup
    )

    process {
        $itemType = $DeviceGroup.Path -replace '(\w+)\[.+\]', '$1'
        $item = Get-ConfigurationItem -Path $DeviceGroup.Path
        $path = ""
        while ($true) {
            $path = "/$($item.DisplayName)" + $path
            if ($item.ParentPath -eq "/$($ItemType)Folder") {
                break;
            }
            $item = Get-ConfigurationItem -Path $item.ParentPath.Replace("/$($ItemType)Folder", "")
        }
        Write-Output $path
    }
}
function Set-VmsCamera {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    [OutputType([VideoOS.Platform.ConfigurationItems.Camera])]
    param(
        # Specifies a camera object such as is returned by Get-VmsCamera. You cannot pass a camera name here.
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.Camera[]]
        $Camera,

        # Specifies a new camera name.
        [Parameter(ValueFromPipelineByPropertyName)]
        [string]
        $Name,

        # Specifies a new camera name.
        [Parameter(ValueFromPipelineByPropertyName)]
        [string]
        $ShortName,

        [Parameter(ValueFromPipelineByPropertyName)]
        [string]
        $Description,

        # Specifies whether the camera is enabled in the VMS. To disable, use "-Enabled $false".
        [Parameter(ValueFromPipelineByPropertyName)]
        [bool]
        $Enabled,

        # Specifies new GPS coordinates in "latitude, longitude" format. This will automatically be converted to "POINT (X Y)" format used in the GisPoint property.
        [Parameter(ValueFromPipelineByPropertyName)]
        [string]
        $Coordinates,

        # Specifies new GPS coordinates in GisPoint format "POINT (longitude latitude)".
        [Parameter(ValueFromPipelineByPropertyName)]
        [string]
        $GisPoint,

        # Specifies the direction the camera is facing as a value from 0 to 360 degrees where 0 is North, 90 is East, 180 is South, and 270 is West.
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(0, 360)]
        [double]
        $Direction,

        # Specifies the direction the camera is facing as a value from 0 to 1. 0 is North, 0.25 is East, 0.5 is South, and 0.75 is West. If using a value from 0 to 360, you can divide the value by 360 to get a value in the correct range.
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(0, 1)]
        [double]
        $CoverageDirection,

        # Specifies the field of view of the camera as a number between 0 and 360 degrees.
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(0, 360)]
        [double]
        $FieldOfView,

        # Specifies the field of view of the camera as a number between 0 and 1, representing degrees from 0 to 360. For example, if the field of view is 54 degrees, you should set the value to (54 / 360) or 0.15.
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(0, 1)]
        [double]
        $CoverageFieldOfView,

        # Specifies the depth or distance of the camera view. The unit of measure will either be "feet" or "meters" depending on the PC's region settings, and the unit can be overridden using the Units parameter. To specify meters or feet explicitly, include "-Units 'Metric'" for meters, or "-Units 'Imperial'" for feet.
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(0, [double]::MaxValue)]
        [double]
        $Depth,

        # Specifies the depth or distance of the camera view in meters, exactly as it is stored in Milestone, with no consideration for the PC's region settings.
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(0, [double]::MaxValue)]
        [double]
        $CoverageDepth,

        # Specifies whether the Depth value, if provided, should be interpreted as a metric unit (meters) or an imperial unit (feet). The default is set automatically based on the region setting of the PC.
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateSet('Metric', 'Imperial')]
        [string]
        $Units,

        # Specifies whether the prebuffer is enabled for the camera(s). To disable, use "-PrebufferEnabled $false".
        [Parameter(ValueFromPipelineByPropertyName)]
        [bool]
        $PrebufferEnabled,

        # Specifies whether the prebuffer should be done in memory. When set to false, prebuffering will be done to disk. To prebuffer to disk, use "-PrebufferInMemory $false".
        [Parameter(ValueFromPipelineByPropertyName)]
        [bool]
        $PrebufferInMemory,

        # Specifies the size of the prebuffer in seconds. Note that the server will not accept a value greater than 15 seconds when prebuffering to memory, or a value greater than 10000 seconds when prebuffering to disk.
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(1, 10000)]
        [int]
        $PrebufferSeconds,

        [Parameter(ValueFromPipelineByPropertyName)]
        [bool]
        $RecordingEnabled,

        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(0, 999)]
        [int]
        $RecordingFramerate,

        [Parameter(ValueFromPipelineByPropertyName)]
        [bool]
        $RecordKeyframesOnly,

        [Parameter(ValueFromPipelineByPropertyName)]
        [bool]
        $RecordOnRelatedDevices,

        [Parameter(ValueFromPipelineByPropertyName)]
        [bool]
        $EdgeStorageEnabled,

        [Parameter(ValueFromPipelineByPropertyName)]
        [bool]
        $EdgeStoragePlaybackEnabled,

        [Parameter(ValueFromPipelineByPropertyName)]
        [bool]
        $ManualRecordingTimeoutEnabled,

        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(0, [int]::MaxValue)]
        [int]
        $ManualRecordingTimeoutMinutes,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        $standardProperties = @(
            'Name',
            'ShortName',
            'Description',
            'Enabled',
            'PrebufferEnabled',
            'PrebufferSeconds',
            'PrebufferInMemory',
            'GisPoint',
            'CoverageDirection',
            'CoverageFieldOfView',
            'CoverageDepth',
            'RecordingEnabled',
            'RecordingFramerate',
            'RecordKeyframesOnly',
            'RecordOnRelatedDevices',
            'EdgeStorageEnabled',
            'EdgeStoragePlaybackEnabled',
            'ManualRecordingTimeoutEnabled',
            'ManualRecordingTimeoutMinutes'
        )
        $IsMetric = [System.Globalization.RegionInfo]::CurrentRegion.IsMetric
        $ConversionFactor = 1
        $metersInAFoot = 0.3048000000012192
    }

    process {
        if (-not [string]::IsNullOrWhiteSpace($Units)) {
            $IsMetric = $Units -eq 'Metric'
        }
        if (-not $IsMetric) {
            $ConversionFactor = $metersInAFoot
        }

        # The $settings hashtable will be loaded with keys matching actual
        # Camera object properties to be changed, and their values.
        $settings = @{}
        foreach ($key in $standardProperties) {
            if ($MyInvocation.BoundParameters.ContainsKey($key)) {
                $settings[$key] = $MyInvocation.BoundParameters[$key]
            }
        }

        # The following section handles the special parameters that don't match
        # up directly to Camera property names. For example we accept a
        # Direction value in degrees from 0-360 because that is user-friendly,
        # and we map that value to the CoverageDirection property which takes a
        # value from 0 to 1.
        if ($MyInvocation.BoundParameters.ContainsKey('Coordinates')) {
            if ([string]::IsNullOrWhiteSpace($Coordinates)) {
                $settings['GisPoint'] = 'POINT EMPTY'
            } else {
                $settings['GisPoint'] = ConvertTo-GisPoint -Coordinates $Coordinates
            }
        }
        if ($MyInvocation.BoundParameters.ContainsKey('Direction')) {
            $settings['CoverageDirection'] = $Direction / 360
        }
        if ($MyInvocation.BoundParameters.ContainsKey('FieldOfView')) {
            $settings['CoverageFieldOfView'] = $FieldOfView / 360
        }
        if ($MyInvocation.BoundParameters.ContainsKey('Depth')) {
            $settings['CoverageDepth'] = $Depth * $ConversionFactor
        }

        # The $settings hashtable is now loaded with property names and values
        # so we will enumerate over the Cameras and if the new value is
        # different than the old value, we'll update it
        foreach ($cam in $Camera) {
            $dirty = $false
            if ($MyInvocation.BoundParameters.ContainsKey('WhatIf')) {
                # This enables us to perform validation against a different camera
                # object reference without changing properties on the original camera
                # object. Otherwise the -WhatIf parameter would still change the state
                # of the source camera object in local memory which could have side effects.
                $cam = [VideoOS.Platform.ConfigurationItems.Camera]::new((Get-Site).FQID.ServerId, $cam.Path)
            }
            foreach ($key in $settings.Keys) {
                if ($cam.$key -ne $settings[$key] -and $PSCmdlet.ShouldProcess($cam.Name, ('Changing {0} from {1} to {2}' -f $key, $cam.$key, $settings[$key]))) {
                    $dirty = $true
                    $cam.$key = $settings[$key]
                }
            }
            if ($cam.PrebufferSeconds -gt 15 -and $cam.PrebufferInMemory) {
                # The validation error for invalid PreBufferSeconds is not informative
                # So we'll handle that here and ensure the value can't be set to something
                # invalid.
                $message = 'PrebufferSeconds exceeds the maximum of value for in-memory buffering. The value will be updated to 15 seconds.'
                if ($script:Messages) {
                    $message = $script:Messages.PrebufferSecondsExceedsMaximumValue
                }
                Write-Warning $message
                $dirty = $true
                $cam.PrebufferSeconds = 15
            }

            try {
                if ($dirty -and $PSCmdlet.ShouldProcess($cam.Name, 'Saving changes')) {
                    # Only save changes and make the API call if we actually changed something.
                    $cam.Save()
                } else {
                    $validation = $cam.ValidateItem()
                    if (-not $validation.ValidatedOk) {
                        foreach ($errorResult in $validation.ErrorResults) {
                            $message = $errorResult.ErrorText
                            if ($script:Messages) {
                                $message = $script:Messages.ClientServiceValidateResult -f $errorResult.ErrorProperty, $cam.($errorResult.ErrorProperty), $errorResult.ErrorText
                            }
                            Write-Warning $message
                        }
                    }
                }
                if ($PassThru) {
                    Write-Output $cam
                }
            } catch {
                Write-Error -Exception $_.Exception -Message $_.Message -TargetObject $cam -Category InvalidOperation
            }
        }
    }
}
function Set-VmsCameraGeneralSetting {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.Camera[]]
        $Camera,

        [Parameter(Mandatory, Position = 0)]
        [hashtable]
        $Settings
    )

    process {
        foreach ($cam in $Camera) {
            $target = $cam.Name
            $deviceDriverSettings = $cam.DeviceDriverSettingsFolder.DeviceDriverSettings[0]
            $generalSettings = $deviceDriverSettings.DeviceDriverSettingsChildItem
            if ($Settings.Keys.Count -gt 0) {
                $dirty = $false
                foreach ($key in $Settings.Keys) {
                    if ($key -notin $generalSettings.Properties.Keys) {
                        Write-Warning "A general setting named '$key' was not found on $($cam.Name)."
                    }

                    $currentValue = $generalSettings.Properties.GetValue($key)
                    if ($null -eq $currentValue -or $currentValue -eq $Settings.$key) {
                        continue
                    }

                    if ($PSCmdlet.ShouldProcess($target, "Changing $key from $currentValue to $($Settings.$key)")) {
                        $generalSettings.Properties.SetValue($key, $Settings.$key)
                        $dirty = $true
                    }
                }
                if ($dirty -and $PSCmdlet.ShouldProcess($target, "Save changes")) {
                    try {
                        $deviceDriverSettings.Save()
                    } catch [VideoOS.Platform.Proxy.ConfigApi.ValidateResultException] {
                        $mipException = $_.Exception -as [VideoOS.Platform.MIPException]
                        foreach ($errorResult in $mipException.ValidateResult.ErrorResults) {
                            $message = $errorResult.ErrorText
                            $null, $key, $null = $errorResult.ErrorProperty -split '/', 3
                            if ($script:Messages -and -not [string]::IsNullOrWhiteSpace($key)) {
                                $message = $script:Messages.ClientServiceValidateResult -f $key, $Settings.$key, $errorResult.ErrorText
                            }
                            Write-Error -Message $message -Exception $mipException
                        }
                        $cam.DeviceDriverSettingsFolder.ClearChildrenCache()
                    }
                }
            }
        }
    }
}
function Set-VmsCameraStream {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param (
        [Parameter(Mandatory, ParameterSetName = 'RemoveStream')]
        [switch]
        $Disabled,

        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'AddOrUpdateStream')]
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'RemoveStream')]
        [VmsCameraStreamConfig[]]
        $Stream,

        [Parameter(ParameterSetName = 'AddOrUpdateStream')]
        [string]
        $DisplayName,

        [Parameter(ParameterSetName = 'AddOrUpdateStream')]
        [switch]
        $Recorded,

        [Parameter(ParameterSetName = 'AddOrUpdateStream')]
        [switch]
        $LiveDefault,

        [Parameter(ParameterSetName = 'AddOrUpdateStream')]
        [ValidateSet('Always', 'Never', 'WhenNeeded', '', $null)]
        [string]
        $LiveMode,

        [Parameter(ParameterSetName = 'AddOrUpdateStream')]
        [hashtable]
        $Settings
    )

    process {
        foreach ($s in $Stream) {
            $target = "$($s.Name) on $($s.Camera.Name)"
            $deviceDriverSettings = $s.Camera.DeviceDriverSettingsFolder.DeviceDriverSettings[0]
            $streamUsages = $s.Camera.StreamFolder.Streams | Select-Object -First 1
            $streamRefToName = @{}
            if ($streamUsages.StreamUsageChildItems.Count -gt 0) {
                $streamNameToRef = $streamUsages.StreamUsageChildItems[0].StreamReferenceIdValues
                foreach ($key in $streamNameToRef.Keys) {
                    $streamRefToName[$streamNameToRef.$key] = $key
                }
                $streamUsageChildItem = $streamUsages.StreamUsageChildItems | Where-Object StreamReferenceId -eq $streamNameToRef[$Stream.Name]
            }

            if ($PSCmdlet.ParameterSetName -eq 'RemoveStream' -and $s.Enabled -and $PSCmdlet.ShouldProcess($s.Camera.Name, "Disabling stream '$($s.Name)'")) {
                if ($streamUsageChildItem.Record -or $streamUsageChildItem.LiveDefault) {
                    Write-Error "Stream $($s.Name) cannot be removed while it is either the LiveDefault or Record stream."
                } else {
                    $result = $streamUsages.RemoveStream($streamUsageChildItem.StreamReferenceId)
                    if ($result.State -eq 'Success') {
                        $s.Camera.StreamFolder.ClearChildrenCache()
                    } else {
                        Write-Error $result.ErrorText
                    }
                }
            } elseif ($PSCmdlet.ParameterSetName -eq 'AddOrUpdateStream') {
                $dirtyStreamUsages = $false
                if ($null -eq $streamUsageChildItem -and ($DisplayName -or $Recorded -or $LiveDefault -or $LiveMode) -and $PSCmdlet.ShouldProcess($s.Camera.Name, 'Adding a new stream usage')) {
                    try {
                        $result = $streamUsages.AddStream()
                        if ($result.State -ne 'Success') {
                            throw $result.ErrorText
                        }
                        $s.Camera.StreamFolder.ClearChildrenCache()
                        $streamUsages = $s.Camera.StreamFolder.Streams[0]
                        $streamUsageChildItem = $streamUsages.StreamUsageChildItems | Where-Object StreamReferenceId -eq $result.GetProperty('StreamReferenceId')
                        $streamUsageChildItem.StreamReferenceId = $streamNameToRef[$s.Name]
                        $streamUsageChildItem.Name = $s.Name
                        $dirtyStreamUsages = $true
                    } catch {
                        Write-Error $_
                    }
                }

                if ($MyInvocation.BoundParameters.ContainsKey('DisplayName') -and $DisplayName -ne $streamUsageChildItem.Name) {
                    if ($PSCmdlet.ShouldProcess($s.Camera.Name, "Setting DisplayName on $($streamUsageChildItem.Name)")) {
                        $streamUsageChildItem.Name = $DisplayName
                    }
                    $dirtyStreamUsages = $true
                }

                if ($MyInvocation.BoundParameters.ContainsKey('Recorded') -and $Recorded -and $Recorded -ne $streamUsageChildItem.Record) {
                    # Find and disable recording on the current recorded stream
                    $recordedStream = $streamUsages.StreamUsageChildItems | Where-Object Record
                    if ($PSCmdlet.ShouldProcess($s.Camera.Name, "Disabling recording on $($recordedStream.Name)")) {
                        $recordedStream.Record = $false
                        if ($recordedStream.LiveMode -eq 'Never' -and $PSCmdlet.ShouldProcess($s.Camera.Name, "Changing LiveMode from Never to WhenNeeded on $($recordedStream.Name)")) {
                            # This avoids a validation exception error.
                            $recordedStream.LiveMode = 'WhenNeeded'
                        }
                    }


                    # Turn recording on the new stream
                    if ($PSCmdlet.ShouldProcess($s.Camera.Name, "Enabling recording on $($streamUsageChildItem.Name)")) {
                        $streamUsageChildItem.Record = $true
                        $dirtyStreamUsages = $true
                    }
                }

                if ($MyInvocation.BoundParameters.ContainsKey('LiveDefault') -and $LiveDefault -and $LiveDefault -ne $streamUsageChildItem.LiveDefault) {
                    # Find and disable recording on the current recorded stream
                    $liveStream = $streamUsages.StreamUsageChildItems | Where-Object LiveDefault
                    if ($PSCmdlet.ShouldProcess($s.Camera.Name, "Disabling LiveDefault on $($liveStream.Name)")) {
                        $liveStream.LiveDefault = $false
                    }

                    # Turn recording on the new stream
                    if ($PSCmdlet.ShouldProcess($s.Camera.Name, "Enabling LiveDefault on $($streamUsageChildItem.Name)")) {
                        $streamUsageChildItem.LiveDefault = $true
                        $dirtyStreamUsages = $true
                    }
                }

                if ($MyInvocation.BoundParameters.ContainsKey('LiveMode') -and $LiveMode -ne $streamUsageChildItem.LiveMode  -and -not [string]::IsNullOrWhiteSpace($LiveMode)) {
                    if ($LiveMode -eq 'Never' -and (-not $streamUsageChildItem.Record -or $streamUsageChildItem.LiveDefault)) {
                        Write-Warning 'The LiveMode property can only be set to "Never" the recorded stream, and only when that stream is not used as the LiveDefault stream.'
                    } elseif ($PSCmdlet.ShouldProcess($s.Camera.Name, "Setting LiveMode on $($streamUsageChildItem.Name)")) {
                        $streamUsageChildItem.LiveMode = $LiveMode
                        $dirtyStreamUsages = $true
                    }
                }

                if ($dirtyStreamUsages -and $PSCmdlet.ShouldProcess($s.Camera.Name, "Saving StreamUsages")) {
                    try {
                        $streamUsages.Save()
                    } catch [VideoOS.Platform.Proxy.ConfigApi.ValidateResultException] {
                        $mipException = $_.Exception -as [VideoOS.Platform.MIPException]
                        foreach ($errorResult in $mipException.ValidateResult.ErrorResults) {
                            $message = $errorResult.ErrorText
                            $null, $errorProperty, $null = $errorResult.ErrorProperty -split '/', 3
                            if ($script:Messages -and -not [string]::IsNullOrWhiteSpace($errorProperty)) {
                                $message = $script:Messages.ClientServiceValidateResult -f $errorProperty, $cam.$errorProperty, $errorResult.ErrorText
                            }
                            Write-Error -Message $message -Exception $mipException
                        }
                        $streamUsages.ClearChildrenCache()
                    }
                }

                $streamChildItem = $deviceDriverSettings.StreamChildItems.Where( { $_.DisplayName -eq $s.Name })
                if ($Settings.Keys.Count -gt 0) {
                    $dirty = $false
                    foreach ($key in $Settings.Keys) {
                        if ($key -notin $s.Settings.Keys) {
                            Write-Warning "A setting with the key '$key' was not found for stream $($streamChildItem.DisplayName) on $($s.Camera.Name)."
                        }

                        $currentValue = $streamChildItem.Properties.GetValue($key)
                        if ($currentValue -eq $Settings.$key) {
                            continue
                        }

                        if ($PSCmdlet.ShouldProcess($target, "Changing $key from $currentValue to $($Settings.$key)")) {
                            $streamChildItem.Properties.SetValue($key, $Settings.$key)
                            $dirty = $true
                        }
                    }
                    if ($dirty -and $PSCmdlet.ShouldProcess($target, "Save changes")) {
                        try {
                            $deviceDriverSettings.Save()
                        } catch [VideoOS.Platform.Proxy.ConfigApi.ValidateResultException] {
                            $mipException = $_.Exception -as [VideoOS.Platform.MIPException]
                            foreach ($errorResult in $mipException.ValidateResult.ErrorResults) {
                                $message = $errorResult.ErrorText
                                $null, $errorProperty, $null = $errorResult.ErrorProperty -split '/', 3
                                if ($script:Messages -and -not [string]::IsNullOrWhiteSpace($errorProperty)) {
                                    $message = $script:Messages.ClientServiceValidateResult -f $errorProperty, $cam.$errorProperty, $errorResult.ErrorText
                                }
                                Write-Error -Message $message -Exception $mipException
                            }
                            $cam.ClearChildrenCache()
                        }
                    }
                }
            }
        }
    }
}
function Set-VmsLicense {
    [CmdletBinding()]
    [OutputType([VideoOS.Platform.ConfigurationItems.LicenseInformation])]
    param (
        [Parameter(Mandatory)]
        [string]
        $Path
    )

    begin {
        $ms = Get-ManagementServer -ErrorAction Stop
        if ([Version]$ms.Version -lt '20.2') {
            throw "Management of Milestone XProtect VMS licensing using MIP SDK was introduced in version 2020 R2 (v20.2). This function is not compatible with the current Management Server version, v$($ms.Version)."
        }
    }

    process {
        try {
            $filePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
            if (-not (Test-Path $filePath)) {
                throw [System.IO.FileNotFoundException]::new('Set-VmsLicense could not find the file.', $filePath)
            }
            $bytes = [IO.File]::ReadAllBytes($filePath)
            $b64 = [Convert]::ToBase64String($bytes)
            $result = $ms.LicenseInformationFolder.LicenseInformations[0].ChangeLicense($b64)
            if ($result.State -eq 'Success') {
                $oldSlc = $ms.LicenseInformationFolder.LicenseInformations[0].Slc
                $ms.ClearChildrenCache()
                $newSlc = $ms.LicenseInformationFolder.LicenseInformations[0].Slc
                if ($oldSlc -eq $newSlc) {
                    Write-Verbose "The software license code in the license file passed to Set-VmsLicense is the same as the existing software license code."
                }
                else {
                    Write-Verbose "Set-VmsLicense changed the software license code from $oldSlc to $newSlc."
                }
                Write-Output $ms.LicenseInformationFolder.LicenseInformations[0]
            }
            else {
                Write-Error "Call to ChangeLicense failed. $($result.ErrorText.Trim('.'))."
            }
        }
        catch {
            Write-Error -Message $_.Message -Exception $_.Exception
        }
    }
}
function Start-VmsHardwareScan {
    [CmdletBinding()]
    [OutputType([VmsHardwareScanResult])]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.RecordingServer[]]
        $RecordingServer,

        [Parameter(Mandatory, ParameterSetName = 'Express')]
        [switch]
        $Express,

        [Parameter(ParameterSetName = 'Manual')]
        [uri[]]
        $Address = @(),

        [Parameter(ParameterSetName = 'Manual')]
        [ipaddress]
        $Start,

        [Parameter(ParameterSetName = 'Manual')]
        [ipaddress]
        $End,

        [Parameter(ParameterSetName = 'Manual')]
        [string]
        $Cidr,

        [Parameter(ParameterSetName = 'Manual')]
        [int]
        $HttpPort = 80,

        [Parameter(ParameterSetName = 'Manual')]
        [int[]]
        $DriverNumber = @(),

        [Parameter(ParameterSetName = 'Manual')]
        [string[]]
        $DriverFamily,

        [Parameter()]
        [pscredential[]]
        $Credential,

        [Parameter()]
        [switch]
        $UseDefaultCredentials,

        [Parameter()]
        [switch]
        $UseHttps,

        [Parameter()]
        [switch]
        $PassThru
    )

    process {
        $tasks = @()
        $recorderPathMap = @{}
        $progressParams = @{
            Activity        = 'Initiating VMS hardware scan'
            PercentComplete = 0
        }
        try {
            switch ($PSCmdlet.ParameterSetName) {
                'Express' {
                    $logins = @()
                    foreach ($c in $Credential) {
                        $logins += [pscustomobject]@{
                            User = $c.UserName
                            Pass = $c.GetNetworkCredential().Password
                        }
                    }
                    try {
                        foreach ($recorder in $RecordingServer) {
                            $progressParams.PercentComplete = [int]($tasks.Count / $RecordingServer.Count * 100)
                            Write-Progress @progressParams
                            $recorderPathMap.($recorder.Path) = $recorder
                            $tasks += $recorder.HardwareScanExpress($logins[0].User, $logins[0].Pass, $logins[1].User, $logins[1].Pass, $logins[2].User, $logins[2].Pass, ($null -eq $Credential -or $UseDefaultCredentials), $UseHttps)
                        }
                    } catch {
                        throw
                    }
                }

                'Manual' {
                    $rangeParameters = ($MyInvocation.BoundParameters.Keys | Where-Object { $_ -in @('Start', 'End') }).Count
                    if ($rangeParameters -eq 1) {
                        Write-Error 'When using the Start or End parameters, you must provide both Start and End parameter values'
                        return
                    }
                    if ($Credential.Count -gt 1) {
                        Write-Warning "Manual address/range scanning supports the use of default credentials and only one user-supplied credential. Only the first of the $($Credential.Count) credentials provided in the Credential parameter will be used."
                    }
                    $Address = $Address | ForEach-Object {
                        if ($_.IsAbsoluteUri) {
                            $_
                        } else {
                            [uri]"http://$($_.OriginalString)"
                        }
                    }
                    if ($MyInvocation.BoundParameters.ContainsKey('UseHttps') -or $MyInvocation.BoundParameters.ContainsKey('HttpPort')) {
                        $Address = $Address | Foreach-Object {
                            $a = [uribuilder]$_
                            if ($MyInvocation.BoundParameters.ContainsKey('UseHttps')) {
                                $a.Scheme = if ($UseHttps) { 'https' } else { 'http' }
                            }
                            if ($MyInvocation.BoundParameters.ContainsKey('HttpPort')) {
                                $a.Port = $HttpPort
                            }
                            $a.Uri
                        }
                    }
                    if ($MyInvocation.BoundParameters.ContainsKey('Start')) {
                        $Address += Expand-IPRange -Start $Start -End $End | ConvertTo-Uri -UseHttps:$UseHttps -HttpPort $HttpPort
                    }
                    if ($MyInvocation.BoundParameters.ContainsKey('Cidr')) {
                        $Address += Expand-IPRange -Cidr $Cidr | Select-Object -Skip 1 | Select-Object -SkipLast 1 | ConvertTo-Uri -UseHttps:$UseHttps -HttpPort $HttpPort
                    }

                    foreach ($entry in $Address) {
                        try {
                            $user, $pass = $null
                            if ($Credential.Count -gt 0) {
                                $user = $Credential[0].UserName
                                $pass = $Credential[0].Password
                            }
                            foreach ($recorder in $RecordingServer) {
                                $progressParams.PercentComplete = [int]($tasks.Count / ($Address.Count * $RecordingServer.Count) * 100)
                                Write-Progress @progressParams
                                if ($MyInvocation.BoundParameters.ContainsKey('DriverFamily')) {
                                    $DriverNumber += $recorder | Get-HardwareDriver | Where-Object { $_.GroupName -in $DriverFamily -and $_.Number -notin $DriverNumber } | Select-Object -ExpandProperty Number
                                }
                                if ($DriverNumber.Count -eq 0) {
                                    Write-Warning "Start-VmsHardwareScan is about to scan $($Address.Count) addresses from $($recorder.Name) without specifying one or more hardware device drivers. This can take a very long time."
                                }
                                $driverNumbers = $DriverNumber -join ';'
                                Write-Verbose "Adding HardwareScan task for $($entry) using driver numbers $driverNumbers"
                                $recorderPathMap.($recorder.Path) = $recorder
                                $tasks += $RecordingServer.HardwareScan($entry.ToString(), $driverNumbers, $user, $pass, ($null -eq $Credential -or $UseDefaultCredentials))
                            }
                        } catch {
                            throw
                        }
                    }
                }
            }
        } finally {
            $progressParams.Completed = $true
            $progressParams.PercentComplete = 100
            Write-Progress @progressParams
        }

        if ($PassThru) {
            Write-Output $tasks
        } else {
            Wait-VmsTask -Path $tasks.Path -Title "Running $(($PSCmdlet.ParameterSetName).ToLower()) hardware scan" -Cleanup | Foreach-Object {
                $state = $_.Properties | Where-Object Key -eq 'State'
                if ($state.Value -eq 'Error') {
                    $errorText = $_.Properties | Where-Object Key -eq 'ErrorText'
                    Write-Error $errorText.Value
                } else {
                    $results = if ($_.Children.Count -gt 0) { [VmsHardwareScanResult[]]$_.Children } else {
                        [VmsHardwareScanResult]$_
                    }
                    foreach ($result in $results) {
                        $result.RecordingServer = $recorderPathMap.($_.ParentPath)
                        Write-Output $result
                    }
                }
            }
        }
    }
}
function Wait-VmsTask {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [ValidateScript({
            if ($_ -notmatch '^Task\[\d+\]$') {
                throw "Path must match the Configuration API Task path format 'Task[number]' where number is an integer.'"
            }
            return $true
        })]
        [string[]]
        $Path,

        [Parameter()]
        [string]
        $Title,

        [Parameter()]
        [switch]
        $Cleanup
    )

    process {
        $tasks = New-Object 'System.Collections.Generic.Queue[VideoOS.ConfigurationApi.ClientService.ConfigurationItem]'
        $Path | Foreach-Object {
            $item = $null
            $errorCount = 0
            while ($null -eq $item) {
                try {
                    $item = Get-ConfigurationItem -Path $_
                }
                catch {
                    $errorCount++
                    if ($errorCount -ge 5) {
                        throw
                    }
                    else {
                        Write-Verbose 'Wait-VmsTask received an error when communicating with Configuration API. The communication channel will be re-established and the connection will be attempted up to 5 times.'
                        Start-Sleep -Seconds 2
                        Get-Site | Select-Site
                    }
                }
            }

            if ($item.ItemType -ne 'Task') {
                Write-Error "Configuration Item with path '$($item.Path)' is incompatible with Wait-VmsTask. Expected an ItemType of 'Task' and received a '$($item.ItemType)'."
            }
            else {
                $tasks.Enqueue($item)
            }
        }
        $completedStates = 'Error', 'Success', 'Completed'
        $totalTasks = $tasks.Count
        $progressParams = @{
            Activity = if ([string]::IsNullOrWhiteSpace($Title)) { 'Waiting for VMS Task(s) to complete' } else { $Title }
            PercentComplete = 0
            Status = 'Processing'
        }
        try {
            Write-Progress @progressParams
            $stopwatch = [diagnostics.stopwatch]::StartNew()
            while ($tasks.Count -gt 0) {
                Start-Sleep -Milliseconds 500
                $taskInfo = $tasks.Dequeue()
                $completedTaskCount = $totalTasks - ($tasks.Count + 1)
                $tasksRemaining = $totalTasks - $completedTaskCount
                $percentComplete = [int]($taskInfo.Properties | Where-Object Key -eq 'Progress' | Select-Object -ExpandProperty Value)

                if ($completedTaskCount -gt 0) {
                    $timePerTask = $stopwatch.ElapsedMilliseconds / $completedTaskCount
                    $remainingTime = [timespan]::FromMilliseconds($tasksRemaining * $timePerTask)
                    $progressParams.SecondsRemaining = [int]$remainingTime.TotalSeconds
                }
                elseif ($percentComplete -gt 0){
                    $pointsRemaining = 100 - $percentComplete
                    $timePerPoint = $stopwatch.ElapsedMilliseconds / $percentComplete
                    $remainingTime = [timespan]::FromMilliseconds($pointsRemaining * $timePerPoint)
                    $progressParams.SecondsRemaining = [int]$remainingTime.TotalSeconds
                }

                if ($tasks.Count -eq 0) {
                    $progressParams.Status = "$($taskInfo.Path) - $($taskInfo.DisplayName)."
                    $progressParams.PercentComplete = $percentComplete
                    Write-Progress @progressParams
                }
                else {
                    $progressParams.Status = "Completed $completedTaskCount of $totalTasks tasks. Remaining tasks: $tasksRemaining"
                    $progressParams.PercentComplete = [int]($completedTaskCount / $totalTasks * 100)
                    Write-Progress @progressParams
                }
                $errorCount = 0
                while ($null -eq $taskInfo) {
                    try {
                        $taskInfo = $taskInfo | Get-ConfigurationItem
                        break
                    }
                    catch {
                        $errorCount++
                        if ($errorCount -ge 5) {
                            throw
                        }
                        else {
                            Write-Verbose 'Wait-VmsTask received an error when communicating with Configuration API. The communication channel will be re-established and the connection will be attempted up to 5 times.'
                            Start-Sleep -Seconds 2
                            Get-Site | Select-Site
                        }
                    }
                }
                $taskInfo = $taskInfo | Get-ConfigurationItem
                if (($taskInfo | Get-ConfigurationItemProperty -Key State) -notin $completedStates) {
                    $tasks.Enqueue($taskInfo)
                    continue
                }
                Write-Output $taskInfo
                if ($Cleanup -and $taskInfo.MethodIds -contains 'TaskCleanup') {
                    $null = $taskInfo | Invoke-Method -MethodId 'TaskCleanup'
                }
            }
        }
        finally {
            $progressParams.Completed = $true
            Write-Progress @progressParams
        }
    }
}
function Export-HardwareCsv {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.Hardware[]]
        $InputObject,
        [Parameter(Mandatory, Position = 1)]
        [string]
        $Path,
        [Parameter()]
        [switch]
        $Full
    )

    begin {
        $exportDirectory = Split-Path -Path $Path -Parent
        if (!(Test-Path $exportDirectory)) {
            $null = New-Item -Path $exportDirectory -ItemType Directory
        }
        $recorderMap = @{}
        Write-Verbose "Caching Recording Server names and IDs"
        foreach ($recorder in Get-ConfigurationItem -Path /RecordingServerFolder -ChildItems) {
            $id = $recorder.Properties | Where-Object { $_.Key -eq "Id" } | Select-Object -ExpandProperty Value -First 1
            $name = $recorder.Properties | Where-Object { $_.Key -eq "Name" } | Select-Object -ExpandProperty Value -First 1
            $recorderMap.Add($recorder.Path, [pscustomobject]@{Id=$id; Name=$name})
        }

        $rows = New-Object System.Collections.ArrayList
    }

    process {
        Write-Warning "Export-HardwareCsv is now deprecated. Please consider using Export-VmsHardware instead."
        foreach ($hardware in $InputObject) {
            Write-Verbose "Retrieving info for $($hardware.Name)"
            try {
                $hardwareSettings = $hardware | Get-HardwareSetting -ErrorAction Ignore
                $mac = if ($hardwareSettings) { $hardwareSettings.MacAddress } else { 'error' }
                $driver = $hardware | Get-HardwareDriver
                $row = New-Object System.Management.Automation.PSObject
                $row | Add-Member -MemberType NoteProperty -Name HardwareName -Value $hardware.Name
                $row | Add-Member -MemberType NoteProperty -Name HardwareAddress -Value $hardware.Address
                $row | Add-Member -MemberType NoteProperty -Name MacAddress -Value $mac
                $row | Add-Member -MemberType NoteProperty -Name UserName -Value $hardware.UserName
                $row | Add-Member -MemberType NoteProperty -Name Password -Value ($hardware | Get-HardwarePassword)
                $row | Add-Member -MemberType NoteProperty -Name DriverNumber -Value $driver.Number
                $row | Add-Member -MemberType NoteProperty -Name DriverDisplayName -Value $driver.DisplayName
                $row | Add-Member -MemberType NoteProperty -Name RecordingServerName -Value $recorderMap[$hardware.ParentItemPath].Name
                $row | Add-Member -MemberType NoteProperty -Name RecordingServerId -Value $recorderMap[$hardware.ParentItemPath].Id

                if ($Full) {
                    $row | Add-Member -MemberType NoteProperty -Name ConfigurationId -Value $hardware.Id
                    $content = $hardware | Get-ConfigurationItem -Recurse -Sort | ConvertTo-Json -Depth 100 -Compress
                    $configPath = Join-Path -Path $exportDirectory -ChildPath "$([System.IO.Path]::GetFileNameWithoutExtension($Path))_$($hardware.Id).json"
                    $content | Set-Content $configPath -Force
                }
                $null = $rows.Add($row)
            } catch {
                Write-Error "Failed to retrieve info for $($hardware.Name). Error: $_"
            }
        }
    }

    end {
        $rows | Export-Csv -Path $Path -NoTypeInformation
    }
}
function Import-HardwareCsv {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, Position = 1)]
        [string]
        $Path,
        [Parameter()]
        [switch]
        $Full,
        [Parameter()]
        [VideoOS.Platform.ConfigurationItems.RecordingServer]
        $RecordingServer
    )

    process {
        Write-Warning "Import-HardwareCsv is now deprecated. Please consider using Import-VmsHardware instead."
        $exportDirectory = Split-Path -Path $Path -Parent
        $rows = @(Import-Csv -Path $Path)
        $recorderMap = @{}
        for ($i = 0; $i -lt $rows.Count; $i++) {
            try {
                Write-Verbose "Processing row $($i + 1) of $($rows.Count)"
                $recorder = if($null -ne $RecordingServer) { $RecordingServer } else {
                    if ($recorderMap.ContainsKey($rows[$i].RecordingServerName)) {
                        $recorderMap[$rows[$i].RecordingServerName]
                    } else {
                        $rec = Get-RecordingServer -Name $rows[$i].RecordingServerName
                        $recorderMap.Add($rec.Name, $rec)
                        $rec
                    }
                }
                Write-Verbose "Adding $($rows[$i].HardwareAddress) to $($recorder.HostName)"

                $hardwareArgs = @{
                    Name = $rows[$i].HardwareName
                    Address = $rows[$i].HardwareAddress
                    UserName = $rows[$i].UserName
                    Password = $rows[$i].Password
                    GroupPath = if ($rows[$i].GroupPath) { $rows[$i].GroupPath } else { '/New Cameras' }
                }

                # Only add DriverId property if DriverNumber is present in this row
                # Rows where DriverNumber is not present will result in the Recording
                # Server scanning to discover the right driver to use.
                if ($rows[$i].DriverNumber) {
                    $hardwareArgs.Add('DriverId', $rows[$i].DriverNumber)
                }

                $hw = $null
                try {
                    $hw = $recorder | Add-Hardware @hardwareArgs -ErrorAction Stop
                    Write-Verbose "Successfully added $($hw.Name) with ID $($hw.Id)"
                    if ($Full -and $null -ne $hw) {
                        $configId = $rows[$i].ConfigurationId
                        $configPath = Join-Path -Path $exportDirectory -ChildPath "$([System.IO.Path]::GetFileNameWithoutExtension($Path))_$configId.json"
                        Get-Content -Path $configPath -Raw | ConvertFrom-Json |
                            Copy-ConfigurationItem -DestinationItem ($hw | Get-ConfigurationItem -Recurse -Sort) -Verbose:$VerbosePreference
                    }
                    Write-Output $hw
                } catch {
                    Write-Error $_
                }
            } catch {
                Write-Error $_
            }
        }
    }
}
function Get-LicenseDetails {
    [CmdletBinding()]
    [OutputType([VideoOS.Platform.ConfigurationItems.LicenseDetailChildItem])]
    param ()
    
    process {
        $site = Get-Site
        $licenseInfo = Get-LicenseInfo
        $licenseInfo.LicenseDetailFolder.LicenseDetailChildItems
    }
}
function Get-LicensedProducts {
    [CmdletBinding()]
    [OutputType([VideoOS.Platform.ConfigurationItems.LicenseInstalledProductChildItem])]
    param ()
    
    process {
        $site = Get-Site
        $licenseInfo = Get-LicenseInfo
        $licenseInfo.LicenseInstalledProductFolder.LicenseInstalledProductChildItems
    }
}
function Get-LicenseInfo {
    [CmdletBinding()]
    [OutputType([VideoOS.Platform.ConfigurationItems.LicenseInformation])]
    param ()
    
    process {
        $site = Get-Site
        [VideoOS.Platform.ConfigurationItems.LicenseInformation]::new($site.FQID.ServerId, "LicenseInformation[$($site.FQID.ObjectId)]")
    }
}
function Get-LicenseOverview {
    [CmdletBinding()]
    [OutputType([VideoOS.Platform.ConfigurationItems.LicenseOverviewAllChildItem])]
    param ()
    
    process {
        $site = Get-Site
        $licenseInfo = Get-LicenseInfo
        $licenseInfo.LicenseOverviewAllFolder.LicenseOverviewAllChildItems
    }
}
function Invoke-LicenseActivation {
    [CmdletBinding()]
    param (
        # Specifies the My Milestone credentials to use for the license activation request
        [Parameter(mandatory)]
        [pscredential]
        $Credential,

        # Specifies whether the provided credentials should be saved and re-used for automatic license activation
        [Parameter()]
        [switch]
        $EnableAutomaticActivation,

        # Specifies that the result of Get-LicenseDetails should be passed into the pipeline after activatino
        [Parameter()]
        [switch]
        $Passthru
    )
    
    process {
        $licenseInfo = Get-LicenseInfo
        $invokeResult = $licenseInfo.ActivateLicense($Credential.UserName, $Credential.Password, $EnableAutomaticActivation)
        do {
            $task = $invokeResult | Get-ConfigurationItem
            $state = $task | Get-ConfigurationItemProperty -Key State
            Write-Verbose ([string]::Join(', ', $task.Properties.Key))
            Start-Sleep -Seconds 1
        } while ($state -ne 'Error' -and $state -ne 'Success')
        if ($state -ne 'Success') {
            Write-Error ($task | Get-ConfigurationItemProperty -Key 'ErrorText')
        }

        if ($Passthru) {
            Get-LicenseDetails
        }
    }
}
function Get-MobileServerInfo {
    [CmdletBinding()]
    param ( )
    process {
        try {
            $mobServerPath = Get-ItemPropertyValue -Path 'HKLM:\SOFTWARE\WOW6432Node\Milestone\XProtect Mobile Server' -Name INSTALLATIONFOLDER
            [Xml]$doc = Get-Content "$mobServerPath.config" -ErrorAction Stop

            $xpath = "/configuration/ManagementServer/Address/add[@key='Ip']"
            $msIp = $doc.SelectSingleNode($xpath).Attributes['value'].Value
            $xpath = "/configuration/ManagementServer/Address/add[@key='Port']"
            $msPort = $doc.SelectSingleNode($xpath).Attributes['value'].Value

            $xpath = "/configuration/HttpMetaChannel/Address/add[@key='Port']"
            $httpPort = [int]::Parse($doc.SelectSingleNode($xpath).Attributes['value'].Value)
            $xpath = "/configuration/HttpMetaChannel/Address/add[@key='Ip']"
            $httpIp = $doc.SelectSingleNode($xpath).Attributes['value'].Value
            if ($httpIp -eq '+') { $httpIp = '0.0.0.0'}

            $xpath = "/configuration/HttpSecureMetaChannel/Address/add[@key='Port']"
            $httpsPort = [int]::Parse($doc.SelectSingleNode($xpath).Attributes['value'].Value)
            $xpath = "/configuration/HttpSecureMetaChannel/Address/add[@key='Ip']"
            $httpsIp = $doc.SelectSingleNode($xpath).Attributes['value'].Value
            if ($httpsIp -eq '+') { $httpsIp = '0.0.0.0'}
            try {
                $hash = Get-HttpSslCertThumbprint -IPPort "$($httpsIp):$($httpsPort)" -ErrorAction Stop
            } catch {
                $hash = $null
            }
            $info = [PSCustomObject]@{
                Version = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($mobServerPath).FileVersion;
                ExePath = $mobServerPath;
                ConfigPath = "$mobServerPath.config";
                ManagementServerIp = $msIp;
                ManagementServerPort = $msPort;
                HttpIp = $httpIp;
                HttpPort = $httpPort;
                HttpsIp = $httpsIp;
                HttpsPort = $httpsPort;
                CertHash = $hash
            }
            $info
        } catch {
            Write-Error $_
        }
    }
}
function Remove-MobileServerCertificate {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact='Medium')]
    param ()

    begin {
        $Elevated = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if (!$Elevated) {
            throw "Elevation is required for Remove-MobileServerCertificate to work properly. Consider re-launching PowerShell by right-clicking and running as Administrator."
        }
    }

    process {
        Write-Warning "This function is deprecated and will be removed in a future version. Use Set-XProtectCertificate which takes advantage of the Milestone Server Configurator CLI."
        try {
            $mosInfo = Get-MobileServerInfo -Verbose:$VerbosePreference
            $ipPort = "$($mosInfo.HttpsIp):$($mosInfo.HttpsPort)"
            if ($mosInfo.CertHash) {
                if ($PSCmdlet.ShouldProcess($ipPort, "Remove SSL certificate binding and restart Milestone XProtect Mobile Server")) {
                    $result = netsh http delete sslcert ipport=$ipPort
                    if ($result -notcontains 'SSL Certificate successfully deleted') {
                        Write-Warning "Unexpected result from netsh http delete sslcert: $result"
                    }
                    Restart-Service -Name 'Milestone XProtect Mobile Server' -Verbose:$VerbosePreference
                }
            } else {
                Write-Warning "No sslcert binding present for $ipPort"
            }
        } catch {
            Write-Error $_
        }
    }
}
function Set-MobileServerCertificate {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact='Medium')]
    param (
        [parameter(ValueFromPipeline=$true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]
        $X509Certificate,

        [parameter(Position = 1, ValueFromPipelineByPropertyName=$true)]
        [string]
        $Thumbprint
    )

    begin {
        $Elevated = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if (!$Elevated) {
            throw "Elevation is required for Set-MobileServerCertificate to work properly. Consider re-launching PowerShell by right-clicking and running as Administrator."
        }
    }

    process {
        Write-Warning "This function is deprecated and will be removed in a future version. Use Set-XProtectCertificate which takes advantage of the Milestone Server Configurator CLI."
        try {
            $mosInfo = Get-MobileServerInfo -Verbose:$VerbosePreference
            $ipPort = "$($mosInfo.HttpsIp):$($mosInfo.HttpsPort)"
            $appId = "{00000000-0000-0000-0000-000000000000}"
            $certHash = if ($null -eq $X509Certificate) { $Thumbprint } else { $X509Certificate.Thumbprint }
            if ($PSCmdlet.ShouldProcess($ipPort, "Add/update SSL certificate binding and restart Milestone XProtect Mobile Server")) {
                if ($null -ne $mosInfo.CertHash) {
                    Remove-MobileServerCertificate -Verbose:$VerbosePreference
                }

                $result = netsh http add sslcert ipport=$ipPort appid="$appId" certhash=$certHash
                if ($result -notcontains 'SSL Certificate successfully added') {
                    Write-Error "Failed to add certificate binding. $result"
                    return
                }
                else {
                    Write-Verbose [string]$result
                }

                Restart-Service -Name 'Milestone XProtect Mobile Server' -Verbose:$VerbosePreference
            }
        } catch {
            Write-Error $_
        }
    }
}
function Set-XProtectCertificate {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        # Specifies the Milestone component on which to update the certificate
        # - Server: Applies to communication between Management Server and Recording Server, as well as client connections to the HTTPS port for the Management Server.
        # - StreamingMedia: Applies to all connections to Recording Servers. Typically on port 7563.
        # - MobileServer: Applies to HTTPS connections to the Milestone Mobile Server.
        [Parameter(Mandatory)]
        [ValidateSet('Server', 'StreamingMedia', 'MobileServer')]
        [string]
        $VmsComponent,

        # Specifies that encryption for the specified Milestone XProtect service should be disabled
        [Parameter(ParameterSetName = 'Disable')]
        [switch]
        $Disable,

        # Specifies the thumbprint of the certificate to apply to Milestone XProtect service
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'Enable')]
        [string]
        $Thumbprint,

        # Specifies the Windows user account for which read access to the private key is required
        [Parameter(ParameterSetName = 'Enable')]
        [string]
        $UserName,

        # Specifies the path to the Milestone Server Configurator executable. The default location is C:\Program Files\Milestone\Server Configurator\ServerConfigurator.exe
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]
        $ServerConfiguratorPath = 'C:\Program Files\Milestone\Server Configurator\ServerConfigurator.exe',

        # Specifies that all certificates issued to
        [Parameter(ParameterSetName = 'Enable')]
        [switch]
        $RemoveOldCert,

        # Specifies that the Server Configurator process should be terminated if it's currently running
        [switch]
        $Force
    )

    begin {
        $principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
        $adminRole = [Security.Principal.WindowsBuiltInRole]::Administrator
        if (-not $principal.IsInRole($adminRole)) {
            throw "Elevation is required. Consider re-launching PowerShell by right-clicking and running as Administrator."
        }

        $certGroups = @{
            Server         = '84430eb7-847c-422d-aa00-7915cd0d7a65'
            StreamingMedia = '549df21d-047c-456b-958e-99e65dd8b3ec'
            MobileServer   = '76cfc719-a852-4210-913e-703eadab139a'
        }

        $knownExitCodes = @{
            0  = 'Success'
            -1 = 'Unknown error'
            -2 = 'Invalid arguments'
            -3 = 'Invalid argument value'
            -4 = 'Another instance is running'
        }
    }

    process {
        $utility = [IO.FileInfo]$ServerConfiguratorPath
        if (-not $utility.Exists) {
            $exception = [System.IO.FileNotFoundException]::new("Milestone Server Configurator not found at $ServerConfiguratorPath", $utility.FullName)
            Write-Error -Message $exception.Message -Exception $exception
            return
        }
        if ($utility.VersionInfo.FileVersion -lt [version]'20.3') {
            Write-Error "Server Configurator version 20.3 is required as the command-line interface for Server Configurator was introduced in version 2020 R3. The current version appears to be $($utility.VersionInfo.FileVersion). Please upgrade to version 2020 R3 or greater."
            return
        }
        Write-Verbose "Verified Server Configurator version $($utility.VersionInfo.FileVersion) is available at $ServerConfiguratorPath"

        $newCert = Get-ChildItem -Path "Cert:\LocalMachine\My\$Thumbprint" -ErrorAction Ignore
        if ($null -eq $newCert -and -not $Disable) {
            Write-Error "Certificate not found in Cert:\LocalMachine\My with thumbprint '$Thumbprint'. Please make sure the certificate is installed in the correct certificate store."
            return
        } elseif ($Thumbprint) {
            Write-Verbose "Located certificate in Cert:\LocalMachine\My with thumbprint $Thumbprint"
        }

        # Add read access to the private key for the specified certificate if UserName was specified
        if (-not [string]::IsNullOrWhiteSpace($UserName)) {
            try {
                Write-Verbose "Ensuring $UserName has the right to read the private key for the specified certificate"
                $newCert | Set-CertKeyPermission -UserName $UserName
            } catch {
                Write-Error -Message "Error granting user '$UserName' read access to the private key for certificate with thumbprint $Thumbprint" -Exception $_.Exception
            }
        }

        if ($Force) {
            if ($PSCmdlet.ShouldProcess("ServerConfigurator", "Kill process if running")) {
                Get-Process -Name ServerConfigurator -ErrorAction Ignore | Foreach-Object {
                    Write-Verbose 'Server Configurator is currently running. The Force switch was provided so it will be terminated.'
                    $_ | Stop-Process
                }
            }
        }

        $procParams = @{
            FilePath               = $utility.FullName
            Wait                   = $true
            PassThru               = $true
            RedirectStandardOutput = Join-Path -Path ([system.environment]::GetFolderPath([system.environment+specialfolder]::ApplicationData)) -ChildPath ([io.path]::GetRandomFileName())
        }
        if ($Disable) {
            $procParams.ArgumentList = '/quiet', '/disableencryption', "/certificategroup=$($certGroups.$VmsComponent)"
        } else {
            $procParams.ArgumentList = '/quiet', '/enableencryption', "/certificategroup=$($certGroups.$VmsComponent)", "/thumbprint=$Thumbprint"
        }
        $argumentString = [string]::Join(' ', $procParams.ArgumentList)
        Write-Verbose "Running Server Configurator with the following arguments: $argumentString"

        if ($PSCmdlet.ShouldProcess("ServerConfigurator", "Start process with arguments '$argumentString'")) {
            $result = Start-Process @procParams
            if ($result.ExitCode -ne 0) {
                Write-Error "Server Configurator exited with code $($result.ExitCode). $($knownExitCodes.$($result.ExitCode))"
                return
            }
        }

        if ($RemoveOldCert) {
            $oldCerts = Get-ChildItem -Path "Cert:\LocalMachine\My" | Where-Object { $_.Subject -eq $newCert.Subject -and $_.Thumbprint -ne $newCert.Thumbprint }
            if ($null -eq $oldCerts) {
                Write-Verbose "No other certificates found matching the subject name $($newCert.Subject)"
                return
            }
            foreach ($cert in $oldCerts) {
                if ($PSCmdlet.ShouldProcess($cert.Thumbprint, "Remove certificate from certificate store")) {
                    Write-Verbose "Removing certificate with thumbprint $($cert.Thumbprint)"
                    $cert | Remove-Item
                }
            }
        }
    }
}
function Get-CameraRecordingStats {
    [CmdletBinding()]
    param(
        # Specifies the Id's of cameras for which to retrieve recording statistics
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [guid[]]
        $Id,

        # Specifies the timestamp from which to start retrieving recording statistics. Default is 7 days prior to 12:00am of the current day.
        [Parameter()]
        [datetime]
        $StartTime = (Get-Date).Date.AddDays(-7),

        # Specifies the timestamp marking the end of the time period for which to retrieve recording statistics. The default is 12:00am of the current day.
        [Parameter()]
        [datetime]
        $EndTime = (Get-Date).Date,

        # Specifies the type of sequence to get statistics on. Default is RecordingSequence.
        [Parameter()]
        [ValidateSet('RecordingSequence', 'MotionSequence')]
        [string]
        $SequenceType = 'RecordingSequence',

        # Specifies that the output should be provided in a complete hashtable instead of one pscustomobject value at a time
        [Parameter()]
        [switch]
        $AsHashTable,

        # Specifies the runspacepool to use. If no runspacepool is provided, one will be created.
        [Parameter()]
        [System.Management.Automation.Runspaces.RunspacePool]
        $RunspacePool
    )

    process {
        if ($EndTime -le $StartTime) {
            throw "EndTime must be greater than StartTime"
        }

        $disposeRunspacePool = $true
        if ($PSBoundParameters.ContainsKey('RunspacePool')) {
            $disposeRunspacePool = $false
        }
        $pool = $RunspacePool
        if ($null -eq $pool) {
            Write-Verbose "Creating a runspace pool"
            $pool = [runspacefactory]::CreateRunspacePool(1, ([int]$env:NUMBER_OF_PROCESSORS + 1))
            $pool.Open()
        }

        $scriptBlock = {
            param(
                [guid]$Id,
                [datetime]$StartTime,
                [datetime]$EndTime,
                [string]$SequenceType
            )

            $sequences = Get-SequenceData -Path "Camera[$Id]" -SequenceType $SequenceType -StartTime $StartTime -EndTime $EndTime -CropToTimeSpan
            $recordedMinutes = $sequences | Foreach-Object {
                ($_.EventSequence.EndDateTime - $_.EventSequence.StartDateTime).TotalMinutes
                } | Measure-Object -Sum | Select-Object -ExpandProperty Sum
            [pscustomobject]@{
                DeviceId = $Id
                StartTime = $StartTime
                EndTime = $EndTime
                SequenceCount = $sequences.Count
                TimeRecorded = [timespan]::FromMinutes($recordedMinutes)
                PercentRecorded = [math]::Round(($recordedMinutes / ($EndTime - $StartTime).TotalMinutes * 100), 1)
            }
        }

        try {
            $threads = New-Object System.Collections.Generic.List[pscustomobject]
            foreach ($cameraId in $Id) {
                $ps = [powershell]::Create()
                $ps.RunspacePool = $pool
                $asyncResult = $ps.AddScript($scriptBlock).AddParameters(@{
                    Id = $cameraId
                    StartTime = $StartTime
                    EndTime = $EndTime
                    SequenceType = $SequenceType
                }).BeginInvoke()
                $threads.Add([pscustomobject]@{
                    DeviceId = $cameraId
                    PowerShell = $ps
                    Result = $asyncResult
                })
            }

            if ($threads.Count -eq 0) {
                return
            }

            $hashTable = @{}
            $completedThreads = New-Object System.Collections.Generic.List[pscustomobject]
            while ($threads.Count -gt 0) {
                foreach ($thread in $threads) {
                    if ($thread.Result.IsCompleted) {
                        if ($AsHashTable) {
                            $hashTable.$($thread.DeviceId.ToString()) = $null
                        }
                        else {
                            $obj = [ordered]@{
                                DeviceId = $thread.DeviceId.ToString()
                                RecordingStats = $null
                            }
                        }
                        try {
                            $result = $thread.PowerShell.EndInvoke($thread.Result) | ForEach-Object { Write-Output $_ }
                            if ($AsHashTable) {
                                $hashTable.$($thread.DeviceId.ToString()) = $result
                            }
                            else {
                                $obj.RecordingStats = $result
                            }
                        }
                        catch {
                            Write-Error $_
                        }
                        finally {
                            $thread.PowerShell.Dispose()
                            $completedThreads.Add($thread)
                            if (!$AsHashTable) {
                                Write-Output ([pscustomobject]$obj)
                            }
                        }
                    }
                }
                $completedThreads | Foreach-Object { [void]$threads.Remove($_)}
                $completedThreads.Clear()
                if ($threads.Count -eq 0) {
                    break;
                }
                Start-Sleep -Milliseconds 250
            }
            if ($AsHashTable) {
                Write-Output $hashTable
            }
        }
        finally {
            if ($threads.Count -gt 0) {
                Write-Warning "Stopping $($threads.Count) running PowerShell instances. This may take a minute. . ."
                foreach ($thread in $threads) {
                    $thread.PowerShell.Dispose()
                }
            }
            if ($disposeRunspacePool) {
                Write-Verbose "Closing runspace pool in $($MyInvocation.MyCommand.Name)"
                $pool.Close()
                $pool.Dispose()
            }
        }
    }
}
function Get-CameraReport {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.RecordingServer[]]
        $RecordingServer,

        [Parameter()]
        [switch]
        $IncludePlainTextPasswords,

        [Parameter()]
        [switch]
        $IncludeDisabled,

        [Parameter()]
        [switch]
        $IncludeSnapshots,

        [Parameter()]
        [ValidateRange(0, [int]::MaxValue)]
        [int]
        $SnapshotHeight = 300,

        [Parameter()]
        [switch]
        $IncludeRecordingStats
    )

    begin {
        $null = Get-ManagementServer -ErrorAction Stop

        $initialSessionState = [initialsessionstate]::CreateDefault()
        foreach ($functionName in @('Get-StreamProperties', 'ConvertFrom-StreamUsage', 'Get-ValueDisplayName', 'ConvertFrom-Snapshot', 'ConvertFrom-GisPoint')) {
            $definition = Get-Content Function:\$functionName -ErrorAction Stop
            $sessionStateFunction = [System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new($functionName, $definition)
            $initialSessionState.Commands.Add($sessionStateFunction)
        }
        $poolSize = [int]$env:NUMBER_OF_PROCESSORS
        $runspacepool = [runspacefactory]::CreateRunspacePool(1, $poolSize, $initialSessionState, $Host)
        $runspacepool.Open()
        $shellPool = New-Object System.Collections.Generic.Queue[System.Management.Automation.PowerShell]
        1..$runspacepool.GetMaxRunspaces() | Foreach-Object {
            $shell = [powershell]::Create()
            $shell.RunspacePool = $runspacepool
            $shellPool.Enqueue($shell)
        }
        $threads = New-Object System.Collections.Generic.List[pscustomobject]
        $processDevice = {
            param(
                [VideoOS.Platform.Messaging.ItemState[]]$States,
                [VideoOS.Platform.ConfigurationItems.RecordingServer]$RecordingServer,
                [hashtable]$VideoDeviceStatistics,
                [hashtable]$CurrentDeviceStatus,
                [hashtable]$RecordingStats,
                [hashtable]$StorageTable,
                [VideoOS.Platform.ConfigurationItems.Hardware]$Hardware,
                [VideoOS.Platform.ConfigurationItems.Camera]$Camera,
                [bool]$IncludePasswords,
                [bool]$IncludeSnapshots,
                [int]$SnapshotHeight
            )
            $returnResult = [pscustomobject]@{
                Data = $null
                ErrorRecord = $null
            }
            try {
                $cameraEnabled = $Hardware.Enabled -and $Camera.Enabled
                $streamUsages = $Camera | Get-Stream -All
                $liveStreamName = $streamUsages | Where-Object LiveDefault | ConvertFrom-StreamUsage
                $recordStreamName = $streamUsages | Where-Object Record | ConvertFrom-StreamUsage
                $liveStreamSettings = $Camera | Get-StreamProperties -StreamName $liveStreamName
                $recordedStreamSettings = if ($liveStreamName -eq $recordStreamName) { $liveStreamSettings } else { $Camera | Get-StreamProperties -StreamName $recordStreamName }

                $motionDetection = $Camera.MotionDetectionFolder.MotionDetections[0]
                $hardwareSettings = $Hardware | Get-HardwareSetting
                $playbackInfo = @{ Begin = 'NotAvailable'; End = 'NotAvailable'}
                if ($cameraEnabled -and $camera.RecordingEnabled) {
                    $tempPlaybackInfo = $Camera | Get-PlaybackInfo -ErrorAction Ignore -WarningAction Ignore
                    if ($null -ne $tempPlaybackInfo) {
                        $playbackInfo = $tempPlaybackInfo
                    }
                }
                $driver = $Hardware | Get-HardwareDriver
                $password = ''
                if ($IncludePasswords) {
                    try {
                        $password = $Hardware | Get-HardwarePassword -ErrorAction Ignore
                    }
                    catch {
                        $password = $_.Message
                    }
                }
                $cameraState = if ($cameraEnabled -and $null -ne $States) { $States | Where-Object { $_.FQID.ObjectId -eq $Camera.Id } | Select-Object -ExpandProperty State } else { 'NotAvailable' }
                $cameraStatus = $CurrentDeviceStatus.$($RecordingServer.Id).CameraDeviceStatusArray | Where-Object DeviceId -eq $Camera.Id
                $statistics = $VideoDeviceStatistics.$($RecordingServer.Id) | Where-Object DeviceId -eq $Camera.Id
                $currentLiveFps = $statistics | Select-Object -ExpandProperty VideoStreamStatisticsArray | Where-Object LiveStreamDefault | Select-Object -ExpandProperty FPS -First 1
                $currentRecFps = $statistics | Select-Object -ExpandProperty VideoStreamStatisticsArray | Where-Object RecordingStream | Select-Object -ExpandProperty FPS -First 1
                $expectedRetention = New-Timespan -Minutes ($StorageTable.$($Camera.RecordingStorage) | ForEach-Object { $_; $_.ArchiveStorageFolder.ArchiveStorages } | Sort-Object RetainMinutes -Descending | Select-Object -First 1 -ExpandProperty RetainMinutes)

                $snapshot = $null
                if ($IncludeSnapshots -and $cameraEnabled -and $cameraStatus.Started -and $cameraState -eq 'Responding') {
                    $snapshot = $Camera | Get-Snapshot -Live -Quality 100 -ErrorAction Ignore | ConvertFrom-Snapshot
                    if ($SnapshotHeight -ne 0 -and $null -ne $snapshot) {
                        $snapshot = $snapshot | Resize-Image -Height $SnapshotHeight -DisposeSource
                    }
                }
                elseif (!$IncludeSnapshots) {
                    $snapshot = 'NotRequested'
                }

                $returnResult.Data = [pscustomobject]@{
                    Name = $Camera.Name
                    Channel = $Camera.Channel
                    Enabled = $cameraEnabled
                    State = $cameraState
                    MediaOverflow = if ($cameraEnabled)  { $cameraStatus.ErrorOverflow } else { 'NotAvailable' }
                    DbRepairInProgress = if ($cameraEnabled)  { $cameraStatus.DbRepairInProgress } else { 'NotAvailable' }
                    DbWriteError = if ($cameraEnabled)  { $cameraStatus.ErrorWritingGop } else { 'NotAvailable' }
                    GpsCoordinates = $Camera | ConvertFrom-GisPoint
                    MediaDatabaseBegin = $playbackInfo.Begin
                    MediaDatabaseEnd = $playbackInfo.End
                    UsedSpaceInBytes = if ($cameraEnabled) { $statistics | Select-Object -ExpandProperty UsedSpaceInBytes } else { 'NotAvailable' }
                    PercentRecordedOneWeek = if ($cameraEnabled -and $RecordingStats.$($Camera.Id).PercentRecorded -is [double]) { $RecordingStats.$($Camera.Id).PercentRecorded } else { 'NotAvailable' }

                    LastModified = $Camera.LastModified
                    Id = $Camera.Id
                    HardwareName = $Hardware.Name
                    Address = $Hardware.Address
                    Username = $Hardware.UserName
                    Password = $password
                    HTTPSEnabled = if ($null -ne $hardwareSettings.HTTPSEnabled) { $hardwareSettings.HTTPSEnabled.ToUpper() } else { 'NO' }
                    MAC = $hardwareSettings.MacAddress
                    Firmware = $hardwareSettings.FirmwareVersion
                    Model = $Hardware.Model
                    Driver = $driver.Name
                    DriverNumber = $driver.Number.ToString()
                    DriverRevision = $driver.DriverRevision
                    HardwareId = $Hardware.Id
                    RecorderName = $RecordingServer.Name
                    RecorderUri = $RecordingServer.WebServerUri
                    RecorderId = $RecordingServer.Id

                    ConfiguredLiveResolution = Get-ValueDisplayName -PropertyList $liveStreamSettings -PropertyName 'Resolution', 'StreamProperty'
                    ConfiguredLiveCodec = Get-ValueDisplayName -PropertyList $liveStreamSettings -PropertyName 'Codec'
                    ConfiguredLiveFPS = Get-ValueDisplayName -PropertyList $liveStreamSettings -PropertyName 'FPS', 'Framerate'
                    LiveMode = $streamUsages | Where-Object LiveDefault | Select-Object -ExpandProperty LiveMode
                    ConfiguredRecordResolution = Get-ValueDisplayName -PropertyList $recordedStreamSettings -PropertyName 'Resolution', 'StreamProperty' #GetResolution -PropertyList $recordedStreamSettings
                    ConfiguredRecordCodec = Get-ValueDisplayName -PropertyList $recordedStreamSettings -PropertyName 'Codec'
                    ConfiguredRecordFPS = Get-ValueDisplayName -PropertyList $recordedStreamSettings -PropertyName 'FPS', 'Framerate'

                    CurrentLiveResolution = if ($cameraEnabled) { $statistics | Select-Object -ExpandProperty VideoStreamStatisticsArray | Where-Object LiveStreamDefault | Select-Object -ExpandProperty ImageResolution -First 1 | Foreach-Object { "$($_.Width)x$($_.Height)" } } else { 'NotAvailable' }
                    CurrentLiveFPS = if ($cameraEnabled -and $currentLiveFps -is [double]) { [math]::Round($currentLiveFps, 1) } else { 'NotAvailable' }
                    CurrentLiveBPS = if ($cameraEnabled) { $statistics | Select-Object -ExpandProperty VideoStreamStatisticsArray | Where-Object LiveStreamDefault | Select-Object -ExpandProperty BPS -First 1 } else { 'NotAvailable' }
                    CurrentRecordedResolution = if ($cameraEnabled) { $statistics | Select-Object -ExpandProperty VideoStreamStatisticsArray | Where-Object RecordingStream | Select-Object -ExpandProperty ImageResolution -First 1 | Foreach-Object { "$($_.Width)x$($_.Height)" } } else { 'NotAvailable' }
                    CurrentRecordedFPS = if ($cameraEnabled -and $currentRecFps -is [double]) { [math]::Round($currentRecFps, 1) } else { 'NotAvailable' }
                    CurrentRecordedBPS = if ($cameraEnabled) { $statistics | Select-Object -ExpandProperty VideoStreamStatisticsArray | Where-Object RecordingStream | Select-Object -ExpandProperty BPS -First 1 } else { 'NotAvailable' }

                    RecordingEnabled = $Camera.RecordingEnabled
                    RecordKeyframesOnly = $Camera.RecordKeyframesOnly
                    RecordOnRelatedDevices = $Camera.RecordOnRelatedDevices
                    PrebufferEnabled = $Camera.PrebufferEnabled
                    PrebufferSeconds = $Camera.PrebufferSeconds
                    PrebufferInMemory = $Camera.PrebufferInMemory

                    RecordingStorageName = $StorageTable.$($Camera.RecordingStorage).Name
                    RecordingPath = [io.path]::Combine($StorageTable.$($Camera.RecordingStorage).DiskPath, $StorageTable.$($Camera.RecordingStorage).Id)
                    ExpectedRetention = $expectedRetention
                    ActualRetention = if ($playbackInfo.Begin -is [string]) { 'NotAvailable' } else { [datetime]::UtcNow - $playbackInfo.Begin }
                    MeetsRetentionPolicy = if ($playbackInfo.Begin -is [string]) { 'NotAvailable' } else { ([datetime]::UtcNow - $playbackInfo.Begin) -ge $expectedRetention }

                    MotionEnabled = $motionDetection.Enabled
                    MotionKeyframesOnly = $motionDetection.KeyframesOnly
                    MotionProcessTime = $motionDetection.ProcessTime
                    MotionSensitivityMode = if ($motionDetection.ManualSensitivityEnabled) { 'Manual' } else { 'Automatic' }
                    MotionManualSensitivity = $motionDetection.ManualSensitivity
                    MotionMetadataEnabled = $motionDetection.GenerateMotionMetadata
                    MotionExcludeRegions = if ($motionDetection.UseExcludeRegions) { 'Yes' } else { 'No' }
                    MotionHardwareAccelerationMode = $motionDetection.HardwareAccelerationMode

                    PrivacyMaskEnabled = $Camera.PrivacyProtectionFolder.PrivacyProtections[0].Enabled

                    Snapshot = $snapshot
                }
            }
            catch {
                $returnResult.ErrorRecord = $_
            }
            return $returnResult
        }
    }

    process {
        Write-Warning "Get-CameraReport is now deprecated. Please consider using Get-VmsCameraReport instead."
        $progressParams = @{
            Activity = 'Camera Report'
            CurrentOperation = ''
            Status = 'Preparing to run report'
            PercentComplete = 0
            Completed = $false
        }
        if ($null -eq $RecordingServer) {
            Write-Verbose "Getting a list of all recording servers on $((Get-ManagementServer).Name)"
            $progressParams.CurrentOperation = 'Getting Recording Servers'
            Write-Progress @progressParams
            $RecordingServer = Get-RecordingServer
        }

        Write-Verbose 'Getting the current state of all cameras'
        $progressParams.CurrentOperation = 'Calling Get-ItemState'
        Write-Progress @progressParams
        $itemState = Get-ItemState
        if ($null -eq $itemState) {
            Write-Warning 'Get-ItemState failed which indicates the Milestone Event Server service may not be 100% functional. It may take longer than normal to run this report if any servers or cameras are not responding.'
        }

        Write-Verbose 'Discovering all cameras and retrieving status and statistics'
        try {
            $progressParams.CurrentOperation = 'Calling Get-VideoDeviceStatistics on all responding recording servers'
            Write-Progress @progressParams
            $respondingRecordingServers = if ($null -eq $itemState) { $RecordingServer.Id } else { $RecordingServer.Id | Where-Object { $id = $_; $id -in $itemState.FQID.ObjectId -and ($itemState | Where-Object { $id -eq $_.FQID.ObjectId }).State -eq 'Server Responding' } }
            $respondingCameras = if ($null -eq $itemState) { (Get-PlatformItem -Kind ([videoos.platform.kind]::camera)).fqid.objectid } else { ($itemState | Where-Object { $_.FQID.Kind -eq [videoos.platform.kind]::Camera -and $_.State -eq 'Responding' }).FQID.ObjectId }
            $videoDeviceStatistics = Get-VideoDeviceStatistics -AsHashtable -RecordingServerId $respondingRecordingServers -RunspacePool $runspacepool

            $progressParams.CurrentOperation = 'Calling Get-CurrentDeviceStatus on all responding recording servers'
            Write-Progress @progressParams
            $currentDeviceStatus = Get-CurrentDeviceStatus -AsHashtable -RecordingServerId $respondingRecordingServers -RunspacePool $runspacepool
            $recordingStats = @{}
            if ($IncludeRecordingStats -and $respondingCameras.Count -gt 0) {
                $progressParams.CurrentOperation = "Retrieving 7 days of recording stats for $($respondingCameras.Count) cameras using Get-CameraRecordingStats"
                Write-Progress @progressParams
                $recordingStats = Get-CameraRecordingStats -Id $respondingCameras -AsHashTable -RunspacePool $runspacepool
            }

            $progressParams.CurrentOperation = 'Adding camera information requests to the queue'
            Write-Progress @progressParams
            $storageTable = @{}
            foreach ($rs in $RecordingServer) {
                $rs.StorageFolder.Storages | Foreach-Object {
                    $_.FillChildren('StorageArchive')
                    $storageTable.$($_.Path) = $_
                }
                foreach ($hw in $rs | Get-Hardware) {
                    foreach ($cam in $hw | Get-VmsCamera) {
                        if (!$IncludeDisabled -and -not ($cam.Enabled -and $hw.Enabled)) {
                            continue
                        }
                        while ($shellPool.Count -eq 0) {
                            ReceiveJobs -Jobs $threads -ShellQueue $shellPool
                            if ($shellPool.Count -eq 0) {
                                Start-Sleep -Milliseconds 100
                            }
                        }
                        $ps = [powershell]::Create()
                        $ps.RunspacePool = $runspacepool
                        $asyncResult = $ps.AddScript($processDevice).AddParameters(@{
                            States = $itemState
                            RecordingServer = $rs
                            VideoDeviceStatistics = $videoDeviceStatistics
                            CurrentDeviceStatus = $currentDeviceStatus
                            RecordingStats = $recordingStats
                            StorageTable = $storageTable
                            Hardware = $hw
                            Camera = $cam
                            IncludePasswords = $IncludePlainTextPasswords
                            IncludeSnapshots = $IncludeSnapshots
                            SnapshotHeight = $SnapshotHeight
                        }).BeginInvoke()
                        $threads.Add([pscustomobject]@{
                            PowerShell = $ps
                            Result = $asyncResult
                            Camera = $cam.Name
                        })
                    }
                }
            }

            if ($threads.Count -eq 0) {
                return
            }
            $progressParams.CurrentOperation = 'Processing'
            $completedThreads = New-Object System.Collections.Generic.List[pscustomobject]
            $totalDevices = $threads.Count
            while ($threads.Count -gt 0) {
                $progressParams.PercentComplete = ($totalDevices - $threads.Count) / $totalDevices * 100
                $progressParams.Status = "Processed $($totalDevices - $threads.Count) out of $totalDevices cameras"
                Write-Progress @progressParams
                ReceiveJobs -Jobs $threads -ShellQueue $shellPool
                Start-Sleep -Milliseconds 100
            }
        }
        finally {
            if ($threads.Count -gt 0) {
                Write-Warning "Stopping $($threads.Count) running PowerShell instances. This may take a minute. . ."
                foreach ($thread in $threads) {
                    $thread.PowerShell.Dispose()
                }
            }
            $runspacepool.Close()
            $runspacepool.Dispose()
            $progressParams.Completed = $true
            Write-Progress @progressParams
        }
    }
}

function ReceiveJobs {
    [CmdletBinding()]
    param(
        [Parameter()]
        [System.Collections.Generic.List[pscustomobject]]$Jobs,
        [Parameter()]
        [System.Collections.Generic.Queue[System.Management.Automation.PowerShell]]$ShellQueue
    )

    process {
        $completedJobs = New-Object System.Collections.Generic.List[pscustomobject]
        $Jobs | Where-Object { $_.Result.IsCompleted } | ForEach-Object {
            $_.PowerShell.EndInvoke($_.Result).Data | Foreach-Object {
                Write-Output $_
            }
            if ($_.PowerShell.HadErrors) {
                $_.PowerShell.Streams.Error | Foreach-Object {
                    Write-Error $_
                }
            }
            $ShellQueue.Enqueue($_.PowerShell)
            $completedJobs.Add($_)
        }
        $completedJobs | Foreach-Object { $null = $Jobs.Remove($_) }
    }
}
function Get-CurrentDeviceStatus {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param (
        # Specifies one or more Recording Server ID's to which the results will be limited. Omit this parameter if you want device status from all Recording Servers
        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('Id')]
        [guid[]]
        $RecordingServerId,

        # Specifies the type of devices to include in the results. By default only cameras will be included and you can expand this to include all device types
        [Parameter()]
        [ValidateSet('Camera', 'Microphone', 'Speaker', 'Metadata', 'Input event', 'Output', 'Event', 'Hardware', 'All')]
        [string[]]
        $DeviceType = 'Camera',

        # Specifies that the output should be provided in a complete hashtable instead of one pscustomobject value at a time
        [Parameter()]
        [switch]
        $AsHashTable,

        # Specifies the runspacepool to use. If no runspacepool is provided, one will be created.
        [Parameter()]
        [System.Management.Automation.Runspaces.RunspacePool]
        $RunspacePool
    )

    process {
        if ($DeviceType -contains 'All') {
            $DeviceType = @('Camera', 'Microphone', 'Speaker', 'Metadata', 'Input event', 'Output', 'Event', 'Hardware')
        }
        $includedDeviceTypes = $DeviceType | Foreach-Object { [videoos.platform.kind]::$_ }

        $disposeRunspacePool = $true
        if ($PSBoundParameters.ContainsKey('RunspacePool')) {
            $disposeRunspacePool = $false
        }
        $pool = $RunspacePool
        if ($null -eq $pool) {
            Write-Verbose "Creating a runspace pool"
            $pool = [runspacefactory]::CreateRunspacePool(1, ([int]$env:NUMBER_OF_PROCESSORS + 1))
            $pool.Open()
        }

        $scriptBlock = {
            param(
                [uri]$Uri,
                [guid[]]$DeviceIds
            )
            try {
                $client = [VideoOS.Platform.SDK.Proxy.Status2.RecorderStatusService2]::new($Uri)
                $client.GetCurrentDeviceStatus((Get-Token), $deviceIds)
            }
            catch {
                throw "Unable to get current device status from $Uri"
            }
        }

        Write-Verbose 'Retrieving recording server information'
        $managementServer = [videoos.platform.configuration]::Instance.GetItems([videoos.platform.itemhierarchy]::SystemDefined) | Where-Object { $_.FQID.Kind -eq [videoos.platform.kind]::Server -and $_.FQID.ObjectId -eq (Get-ManagementServer).Id }
        $recorders = $managementServer.GetChildren() | Where-Object { $_.FQID.ServerId.ServerType -eq 'XPCORS' -and ($null -eq $RecordingServerId -or $_.FQID.ObjectId -in $RecordingServerId) }
        Write-Verbose "Retrieving video device statistics from $($recorders.Count) recording servers"
        try {
            $threads = New-Object System.Collections.Generic.List[pscustomobject]
            foreach ($recorder in $recorders) {
                Write-Verbose "Requesting device status from $($recorder.Name) at $($recorder.FQID.ServerId.Uri)"
                $folders = $recorder.GetChildren() | Where-Object { $_.FQID.Kind -in $includedDeviceTypes -and $_.FQID.FolderType -eq [videoos.platform.foldertype]::SystemDefined}
                $deviceIds = [guid[]]($folders | Foreach-Object {
                    $children = $_.GetChildren()
                    if ($null -ne $children -and $children.Count -gt 0) {
                        $children.FQID.ObjectId
                    }
                })

                $ps = [powershell]::Create()
                $ps.RunspacePool = $pool
                $asyncResult = $ps.AddScript($scriptBlock).AddParameters(@{
                    Uri = $recorder.FQID.ServerId.Uri
                    DeviceIds = $deviceIds
                }).BeginInvoke()
                $threads.Add([pscustomobject]@{
                    RecordingServerId = $recorder.FQID.ObjectId
                    RecordingServerName = $recorder.Name
                    PowerShell = $ps
                    Result = $asyncResult
                })
            }

            if ($threads.Count -eq 0) {
                return
            }

            $hashTable = @{}
            $completedThreads = New-Object System.Collections.Generic.List[pscustomobject]
            while ($threads.Count -gt 0) {
                foreach ($thread in $threads) {
                    if ($thread.Result.IsCompleted) {
                        Write-Verbose "Receiving results from recording server $($thread.RecordingServerName)"
                        if ($AsHashTable) {
                            $hashTable.$($thread.RecordingServerId.ToString()) = $null
                        }
                        else {
                            $obj = @{
                                RecordingServerId = $thread.RecordingServerId.ToString()
                                CurrentDeviceStatus = $null
                            }
                        }
                        try {
                            $result = $thread.PowerShell.EndInvoke($thread.Result) | ForEach-Object { Write-Output $_ }
                            if ($AsHashTable) {
                                $hashTable.$($thread.RecordingServerId.ToString()) = $result
                            }
                            else {
                                $obj.CurrentDeviceStatus = $result
                            }
                        }
                        catch {
                            Write-Error $_
                        }
                        finally {
                            $thread.PowerShell.Dispose()
                            $completedThreads.Add($thread)
                            if (!$AsHashTable) {
                                Write-Output ([pscustomobject]$obj)
                            }
                        }
                    }
                }
                $completedThreads | Foreach-Object { [void]$threads.Remove($_)}
                $completedThreads.Clear()
                if ($threads.Count -eq 0) {
                    break;
                }
                Start-Sleep -Milliseconds 250
            }
            if ($AsHashTable) {
                Write-Output $hashTable
            }
        }
        finally {
            if ($threads.Count -gt 0) {
                Write-Warning "Stopping $($threads.Count) running PowerShell instances. This may take a minute. . ."
                foreach ($thread in $threads) {
                    $thread.PowerShell.Dispose()
                }
            }
            if ($disposeRunspacePool) {
                Write-Verbose "Closing runspace pool in $($MyInvocation.MyCommand.Name)"
                $pool.Close()
                $pool.Dispose()
            }
        }
    }
}
function Get-VideoDeviceStatistics {
    [CmdletBinding()]
    param (
        # Specifies one or more Recording Server ID's to which the results will be limited. Omit this parameter if you want device status from all Recording Servers
        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('Id')]
        [guid[]]
        $RecordingServerId,

        # Specifies that the output should be provided in a complete hashtable instead of one pscustomobject value at a time
        [Parameter()]
        [switch]
        $AsHashTable,

        # Specifies the runspacepool to use. If no runspacepool is provided, one will be created.
        [Parameter()]
        [System.Management.Automation.Runspaces.RunspacePool]
        $RunspacePool
    )

    process {
        $disposeRunspacePool = $true
        if ($PSBoundParameters.ContainsKey('RunspacePool')) {
            $disposeRunspacePool = $false
        }
        $pool = $RunspacePool
        if ($null -eq $pool) {
            Write-Verbose "Creating a runspace pool"
            $pool = [runspacefactory]::CreateRunspacePool(1, ([int]$env:NUMBER_OF_PROCESSORS + 1))
            $pool.Open()
        }

        $scriptBlock = {
            param(
                [uri]$Uri,
                [guid[]]$DeviceIds
            )
            try {
                $client = [VideoOS.Platform.SDK.Proxy.Status2.RecorderStatusService2]::new($Uri)
                $client.GetVideoDeviceStatistics((Get-Token), $deviceIds)
            }
            catch {
                throw "Unable to get video device statistics from $Uri"
            }

        }

        Write-Verbose 'Retrieving recording server information'
        $managementServer = [videoos.platform.configuration]::Instance.GetItems([videoos.platform.itemhierarchy]::SystemDefined) | Where-Object { $_.FQID.Kind -eq [videoos.platform.kind]::Server -and $_.FQID.ObjectId -eq (Get-ManagementServer).Id }
        $recorders = $managementServer.GetChildren() | Where-Object { $_.FQID.ServerId.ServerType -eq 'XPCORS' -and ($null -eq $RecordingServerId -or $_.FQID.ObjectId -in $RecordingServerId) }
        Write-Verbose "Retrieving video device statistics from $($recorders.Count) recording servers"
        try {
            $threads = New-Object System.Collections.Generic.List[pscustomobject]
            foreach ($recorder in $recorders) {
                Write-Verbose "Requesting video device statistics from $($recorder.Name) at $($recorder.FQID.ServerId.Uri)"
                $folders = $recorder.GetChildren() | Where-Object { $_.FQID.Kind -eq [videoos.platform.kind]::Camera -and $_.FQID.FolderType -eq [videoos.platform.foldertype]::SystemDefined}
                $deviceIds = [guid[]]($folders | Foreach-Object {
                    $children = $_.GetChildren()
                    if ($null -ne $children -and $children.Count -gt 0) {
                        $children.FQID.ObjectId
                    }
                })

                $ps = [powershell]::Create()
                $ps.RunspacePool = $pool
                $asyncResult = $ps.AddScript($scriptBlock).AddParameters(@{
                    Uri = $recorder.FQID.ServerId.Uri
                    DeviceIds = $deviceIds
                }).BeginInvoke()
                $threads.Add([pscustomobject]@{
                    RecordingServerId = $recorder.FQID.ObjectId
                    RecordingServerName = $recorder.Name
                    PowerShell = $ps
                    Result = $asyncResult
                })
            }

            if ($threads.Count -eq 0) {
                return
            }

            $hashTable = @{}
            $completedThreads = New-Object System.Collections.Generic.List[pscustomobject]
            while ($threads.Count -gt 0) {
                foreach ($thread in $threads) {
                    if ($thread.Result.IsCompleted) {
                        Write-Verbose "Receiving results from recording server $($thread.RecordingServerName)"
                        if ($AsHashTable) {
                            $hashTable.$($thread.RecordingServerId.ToString()) = $null
                        }
                        else {
                            $obj = @{
                                RecordingServerId = $thread.RecordingServerId.ToString()
                                VideoDeviceStatistics = $null
                            }
                        }
                        try {
                            $result = $thread.PowerShell.EndInvoke($thread.Result) | ForEach-Object { Write-Output $_ }
                            if ($AsHashTable) {
                                $hashTable.$($thread.RecordingServerId.ToString()) = $result
                            }
                            else {
                                $obj.VideoDeviceStatistics = $result
                            }
                        }
                        catch {
                            Write-Error $_
                        }
                        finally {
                            $thread.PowerShell.Dispose()
                            $completedThreads.Add($thread)
                            if (!$AsHashTable) {
                                Write-Output ([pscustomobject]$obj)
                            }
                        }
                    }
                }
                $completedThreads | Foreach-Object { [void]$threads.Remove($_)}
                $completedThreads.Clear()
                if ($threads.Count -eq 0) {
                    break;
                }
                Start-Sleep -Milliseconds 250
            }
            if ($AsHashTable) {
                Write-Output $hashTable
            }
        }
        finally {
            if ($threads.Count -gt 0) {
                Write-Warning "Stopping $($threads.Count) running PowerShell instances. This may take a minute. . ."
                foreach ($thread in $threads) {
                    $thread.PowerShell.Dispose()
                }
            }
            if ($disposeRunspacePool) {
                Write-Verbose "Closing runspace pool in $($MyInvocation.MyCommand.Name)"
                $pool.Close()
                $pool.Dispose()
            }
        }
    }
}
function Get-VmsCameraReport {
    [CmdletBinding()]
    param (
        [Parameter()]
        [VideoOS.Platform.ConfigurationItems.RecordingServer[]]
        $RecordingServer,

        [Parameter()]
        [switch]
        $IncludePlainTextPasswords,

        [Parameter()]
        [switch]
        $IncludeRetentionInfo,

        [Parameter()]
        [switch]
        $IncludeRecordingStats,

        [Parameter()]
        [switch]
        $IncludeSnapshots,

        [Parameter()]
        [ValidateRange(1, [int]::MaxValue)]
        [int]
        $SnapshotHeight = 300,

        [Parameter()]
        [ValidateSet('All', 'Disabled', 'Enabled')]
        [string]
        $EnableFilter = 'Enabled'
    )

    begin {
        $ms = Get-ManagementServer -ErrorAction Stop
        for ($attempt = 1; $attempt -le 2; $attempt++) {
            try {
                $supportsFillChildren = [version]$ms.Version -ge '20.2'
                $scs = Get-IServerCommandService -ErrorAction Stop
                $config = $scs.GetConfiguration((MilestonePSTools\Get-Token))
                $recorderCameraMap = @{}
                $config.Recorders | Foreach-Object {
                    $deviceList = New-Object System.Collections.Generic.List[guid]
                    $_.Cameras.DeviceId | Foreach-Object { if ($_) { $deviceList.Add($_) } }
                    $recorderCameraMap.($_.RecorderId) = $deviceList
                }
                break
            } catch {
                if ($attempt -ge 2) {
                    throw
                }
                # Typically if an error is thrown here, it's on $scs.GetConfiguration because the
                # IServerCommandService WCF channel is cached and reused, and might be timed out.
                # The Select-Site cmdlet has a side effect of flushing all cached WCF channels.
                Get-Site | Select-Site
            }
        }
        $isAdmin = $null -ne ((Get-LoginSettings | Where-Object Guid -eq (Get-Site).FQID.ObjectId).GroupMembership | Where-Object { $_ -as [guid] } | Foreach-Object { Get-Role -RoleId $_ } | Where-Object RoleType -eq Adminstrative)
        $jobRunner = [LocalJobRunner]@{
            JobPollingInterval = [timespan]::FromMilliseconds(500)
        }
    }

    process {
        try {
            if ($IncludePlainTextPasswords -and -not $isAdmin) {
                Write-Warning $script:Messages.MustBeAdminToReadPasswords
            }
            if (-not $RecordingServer) {
                Write-Verbose $script:Messages.ListingAllRecorders
                $RecordingServer = Get-RecordingServer
            }
            $cache = @{
                DeviceState    = @{}
                PlaybackInfo   = @{}
                Snapshots      = @{}
                Passwords      = @{}
                RecordingStats = @{}
            }

            $ids = @()
            $RecordingServer | Foreach-Object { $ids += $recorderCameraMap[[guid]$_.Id] }

            Write-Verbose $script:Messages.CallingGetItemState
            Get-ItemState -CamerasOnly -ErrorAction Ignore | Foreach-Object {
                $cache.DeviceState[$_.FQID.ObjectId] = @{
                    ItemState = $_.State
                }
            }

            Write-Verbose $script:Messages.StartingFillChildrenThreadJob
            $fillChildrenJobs = $RecordingServer | Foreach-Object {
                $jobRunner.AddJob(
                    {
                        param([bool]$supportsFillChildren, [object]$recorder, [string]$EnableFilter, [bool]$getPasswords, [hashtable]$cache)

                        $manualMethod = {
                            param([object]$recorder)
                            $null = $recorder.HardwareDriverFolder.HardwareDrivers
                            $null = $recorder.StorageFolder.Storages.ArchiveStorageFolder.ArchiveStorages
                            $null = $recorder.HardwareFolder.Hardwares.HardwareDriverSettingsFolder.HardwareDriverSettings
                            $null = $recorder.HardwareFolder.Hardwares.CameraFolder.Cameras.StreamFolder.Streams
                            $null = $recorder.HardwareFolder.Hardwares.CameraFolder.Cameras.DeviceDriverSettingsFolder.DeviceDriverSettings
                        }
                        if ($supportsFillChildren) {
                            try {
                                $itemTypes = 'Hardware', 'HardwareDriverFolder', 'HardwareDriver', 'HardwareDriverSettingsFolder', 'HardwareDriverSettings', 'StorageFolder', 'Storage', 'StorageInformation', 'ArchiveStorageFolder', 'ArchiveStorage', 'CameraFolder', 'Camera', 'DeviceDriverSettingsFolder', 'DeviceDriverSettings', 'MotionDetectionFolder', 'MotionDetection', 'StreamFolder', 'Stream', 'StreamSettings', 'StreamDefinition'
                                $alwaysIncludedItemTypes = @('MotionDetection', 'HardwareDriver', 'HardwareDriverSettings', 'Hardware', 'Storage', 'ArchiveStorage', 'DeviceDriverSettings')
                                $supportsPrivacyMask = (Get-IServerCommandService).GetConfiguration((MilestonePSTools\Get-Token)).ServerOptions | Where-Object Key -eq 'PrivacyMask' | Select-Object -ExpandProperty Value
                                if ($supportsPrivacyMask -eq 'True') {
                                    $alwaysIncludedItemTypes += 'PrivacyProtectionFolder', 'PrivacyProtection'
                                }
                                $itemFilters = $itemTypes | Foreach-Object {
                                    $enableFilterSelection = if ($_ -in $alwaysIncludedItemTypes) { 'All' } else { $EnableFilter }
                                    [VideoOS.ConfigurationApi.ClientService.ItemFilter]@{
                                        ItemType        = $_
                                        EnableFilter    = $enableFilterSelection
                                        PropertyFilters = @()
                                    }
                                }
                                $recorder.FillChildren($itemTypes, $itemFilters)

                                # TODO: Remove this after TFS 447559 is addressed. The StreamFolder.Streams collection is empty after using FillChildren
                                # So this entire foreach block is only necessary to flush the children of StreamFolder and force another query for every
                                # camera so we can fill the collection up in this background task before enumerating over everything at the end.
                                foreach ($hw in $recorder.hardwarefolder.hardwares) {
                                    if ($getPasswords) {
                                        $password = $hw.ReadPasswordHardware().GetProperty('Password')
                                        $cache.Passwords[[guid]$hw.Id] = $password
                                    }
                                    foreach ($cam in $hw.camerafolder.cameras) {
                                        try {
                                            if ($cam.StreamFolder.Streams.Count -eq 0) {
                                                $cam.StreamFolder.ClearChildrenCache()
                                                $null = $cam.StreamFolder.Streams
                                            }
                                        }
                                        catch {
                                            Write-Error $_
                                        }
                                    }
                                }
                            }
                            catch {
                                Write-Error $_
                                $manualMethod.Invoke($recorder)
                            }
                        }
                        else {
                            $manualMethod.Invoke($recorder)
                        }
                    },
                    @{ SupportsFillChildren = $supportsFillChildren; recorder = $_; EnableFilter = $EnableFilter; getPasswords = ($isAdmin -and $IncludePlainTextPasswords); cache = $cache }
                )
            }

            # Kick off snapshots early if requested. Pick up results at the end.
            $snapshotsById = @{}
            if ($IncludeSnapshots) {
                Write-Verbose "Starting Get-Snapshot threadjob"
                $snapshotScriptBlock = {
                    param([guid[]]$ids, [int]$snapshotHeight, [hashtable]$snapshotsById, [hashtable]$cache)
                    foreach ($id in $ids) {
                        $itemState = $cache.DeviceState[$id].ItemState
                        if (-not [string]::IsNullOrWhiteSpace($itemState) -and $itemState -ne 'Responding') {
                            # Do not attempt to get a live image if the event server says the camera is not responding. Saves time.
                            continue
                        }
                        $snapshot = Get-Snapshot -CameraId $id -Live -Quality 100
                        if ($null -ne $snapshot) {
                            $image = $snapshot | ConvertFrom-Snapshot | Resize-Image -Height $snapshotHeight -DisposeSource
                            $snapshotsById[$id] = $image
                        }
                    }
                }
                $snapshotsJob = $jobRunner.AddJob($snapshotScriptBlock, @{ids = $ids; snapshotHeight = $SnapshotHeight; snapshotsById = $snapshotsById; cache = $cache })
            }

            if ($IncludeRetentionInfo) {
                Write-Verbose 'Starting Get-PlaybackInfo threadjob'
                $playbackInfoScriptblock = {
                    param(
                        [guid]$id,
                        [hashtable]$cache
                    )

                    $info = Get-PlaybackInfo -Path "Camera[$id]"
                    if ($null -ne $info) {
                        $cache.PlaybackInfo[$id] = $info
                    }
                }
                $playbackInfoJobs = $ids | Foreach-Object {
                    $jobRunner.AddJob($playbackInfoScriptblock, @{ id = $_; cache = $cache } )
                }
            }

            if ($IncludeRecordingStats) {
                Write-Verbose 'Starting recording stats threadjob'
                $recordingStatsScript = {
                    param(
                        [guid]$Id,
                        [datetime]$StartTime,
                        [datetime]$EndTime,
                        [string]$SequenceType
                    )

                    $sequences = Get-SequenceData -Path "Camera[$Id]" -SequenceType $SequenceType -StartTime $StartTime -EndTime $EndTime -CropToTimeSpan
                    $recordedMinutes = $sequences | Foreach-Object {
                        ($_.EventSequence.EndDateTime - $_.EventSequence.StartDateTime).TotalMinutes
                        } | Measure-Object -Sum | Select-Object -ExpandProperty Sum
                    [pscustomobject]@{
                        DeviceId = $Id
                        StartTime = $StartTime
                        EndTime = $EndTime
                        SequenceCount = $sequences.Count
                        TimeRecorded = [timespan]::FromMinutes($recordedMinutes)
                        PercentRecorded = [math]::Round(($recordedMinutes / ($EndTime - $StartTime).TotalMinutes * 100), 1)
                    }
                }
                $recordingStatsJobs = $ids | Foreach-Object {
                    $jobRunner.AddJob($recordingStatsScript, @{Id = $_; StartTime = (Get-Date).Date.AddDays(-7); EndTime = (Get-Date).Date; SequenceType = 'RecordingSequence'})
                }
            }

            # Get VideoDeviceStatistics for all Recording Servers in the report
            Write-Verbose 'Starting GetVideoDeviceStatistics threadjob'
            $videoDeviceStatsScriptBlock = {
                param(
                    [VideoOS.Platform.SDK.Proxy.Status2.RecorderStatusService2]$svc,
                    [guid[]]$ids
                )
                $svc.GetVideoDeviceStatistics((MilestonePSTools\Get-Token), $ids)
            }
            $videoDeviceStatsJobs = $RecordingServer | ForEach-Object {
                $jobRunner.AddJob($videoDeviceStatsScriptBlock, @{ svc = ($_ | Get-RecorderStatusService2); ids = $recorderCameraMap[[guid]$_.Id] })
            }

            # Get Current Device Status for everything in the report
            Write-Verbose 'Starting GetCurrentDeviceStatus threadjob'
            $currentDeviceStatsJobsScriptBlock = {
                param(
                        [VideoOS.Platform.SDK.Proxy.Status2.RecorderStatusService2]$svc,
                        [guid[]]$ids
                    )
                    $svc.GetCurrentDeviceStatus((MilestonePSTools\Get-Token), $ids)
            }
            $currentDeviceStatsJobs = $RecordingServer | Foreach-Object {
                $svc = $_ | Get-RecorderStatusService2
                $jobRunner.AddJob($currentDeviceStatsJobsScriptBlock, @{svc = $svc; ids = $recorderCameraMap[[guid]$_.Id] })
            }

            Write-Verbose 'Receiving results of FillChildren threadjob'
            $jobRunner.Wait($fillChildrenJobs)
            $fillChildrenResults = $jobRunner.ReceiveJobs($fillChildrenJobs)
            foreach ($e in $fillChildrenResults.Errors) {
                Write-Error $e
            }

            if ($IncludeRetentionInfo) {
                Write-Verbose 'Receiving results of Get-PlaybackInfo threadjob'
                $jobRunner.Wait($playbackInfoJobs)
                $playbackInfoResult = $jobRunner.ReceiveJobs($playbackInfoJobs)
                foreach ($e in $playbackInfoResult.Errors) {
                    Write-Error $e
                }
            }

            if ($IncludeRecordingStats) {
                Write-Verbose 'Receiving results of recording stats threadjob'
                $jobRunner.Wait($recordingStatsJobs)
                foreach ($job in $jobRunner.ReceiveJobs($recordingStatsJobs)) {
                    if ($job.Output.DeviceId) {
                        $cache.RecordingStats[$job.Output.DeviceId] = $job.Output
                    }
                    foreach ($e in $job.Errors) {
                        Write-Error $e
                    }
                }
            }

            Write-Verbose 'Receiving results of GetVideoDeviceStatistics threadjobs'
            $jobRunner.Wait($videoDeviceStatsJobs)
            foreach ($job in $jobRunner.ReceiveJobs($videoDeviceStatsJobs)) {
                foreach ($result in $job.Output) {
                    if (-not $cache.DeviceState.ContainsKey($result.DeviceId)) {
                        $cache.DeviceState[$result.DeviceId] = @{}
                    }
                    $cache.DeviceState[$result.DeviceId].UsedSpaceInBytes = $result.UsedSpaceInBytes
                    $cache.DeviceState[$result.DeviceId].VideoStreamStatisticsArray = $result.VideoStreamStatisticsArray
                }
                foreach ($e in $job.Errors) {
                    Write-Error $e
                }
            }

            Write-Verbose 'Receiving results of GetCurrentDeviceStatus threadjobs'
            $jobRunner.Wait($currentDeviceStatsJobs)
            $currentDeviceStatsResult = $jobRunner.ReceiveJobs($currentDeviceStatsJobs)
            $currentDeviceStatsResult.Output | Foreach-Object {
                foreach ($row in $_.CameraDeviceStatusArray) {
                    if (-not $cache.DeviceState.ContainsKey($row.DeviceId)) {
                        $cache.DeviceState[$row.DeviceId] = @{}
                    }
                    $cache.DeviceState[$row.DeviceId].Status = $row
                }
            }
            foreach ($e in $currentDeviceStatsResult.Errors) {
                Write-Error $e
            }

            if ($null -ne $snapshotsJob) {
                Write-Verbose 'Receiving results of Get-Snapshot threadjob'
                $jobRunner.Wait($snapshotsJob)
                $snapshotsResult = $jobRunner.ReceiveJobs($snapshotsJob)
                $cache.Snapshots = $snapshotsById
                foreach ($e in $snapshotsResult.Errors) {
                    Write-Error $e
                }
            }

            foreach ($rec in $RecordingServer) {
                foreach ($hw in $rec.HardwareFolder.Hardwares | Where-Object { if ($EnableFilter -eq 'All') { $true } else { $_.Enabled } }) {
                    try {
                        $hwSettings = ConvertFrom-ConfigurationApiProperties -Properties $hw.HardwareDriverSettingsFolder.HardwareDriverSettings[0].HardwareDriverSettingsChildItems[0].Properties -UseDisplayNames
                        $driver = $rec.HardwareDriverFolder.HardwareDrivers | Where-Object Path -eq $hw.HardwareDriverPath
                        foreach ($cam in $hw.CameraFolder.Cameras | Where-Object { if ($EnableFilter -eq 'All') { $true } elseif ($EnableFilter -eq 'Enabled') { $_.Enabled -and $hw.Enabled } else { !$_.Enabled -or !$hw.Enabled } }) {
                            $id = [guid]$cam.Id
                            $state = $cache.DeviceState[$id]
                            $storage = $rec.StorageFolder.Storages | Where-Object Path -eq $cam.RecordingStorage
                            $motion = $cam.MotionDetectionFolder.MotionDetections[0]
                            if ($cam.StreamFolder.Streams.Count -gt 0) {
                                $liveStreamUsage = $cam.StreamFolder.Streams[0].StreamUsageChildItems | Where-Object LiveDefault
                                $liveStreamName = $liveStreamUsage.StreamReferenceIdValues.Keys | Where-Object { $liveStreamUsage.StreamReferenceIdValues.$_ -eq $liveStreamUsage.StreamReferenceId }
                                $liveStreamChildItem = $cam.DeviceDriverSettingsFolder.DeviceDriverSettings[0].StreamChildItems | Where-Object DisplayName -eq $liveStreamName
                                $liveStreamSettings = ConvertFrom-ConfigurationApiProperties -Properties $liveStreamChildItem.Properties -UseDisplayNames
                                $liveStreamStats = $state.VideoStreamStatisticsArray | Where-Object StreamId -eq $liveStreamUsage.StreamReferenceId
                                $recordedStreamUsage = $cam.StreamFolder.Streams[0].StreamUsageChildItems | Where-Object Record
                                $recordedStreamName = $recordedStreamUsage.StreamReferenceIdValues.Keys | Where-Object { $recordedStreamUsage.StreamReferenceIdValues.$_ -eq $recordedStreamUsage.StreamReferenceId }
                                $recordedStreamChildItem = $cam.DeviceDriverSettingsFolder.DeviceDriverSettings[0].StreamChildItems | Where-Object DisplayName -eq $recordedStreamName
                                $recordedStreamSettings = ConvertFrom-ConfigurationApiProperties -Properties $recordedStreamChildItem.Properties -UseDisplayNames
                                $recordedStreamStats = $state.VideoStreamStatisticsArray | Where-Object StreamId -eq $recordedStreamUsage.StreamReferenceId
                            }
                            else {
                                Write-Warning "Live & recorded stream properties unavailable for $($cam.Name) as the camera does not support multi-streaming."
                                $liveStreamName = ''
                                $recordedStreamName = ''
                            }
                            $obj = [ordered]@{
                                Name                         = $cam.Name
                                Channel                      = $cam.Channel
                                Enabled                      = $cam.Enabled -and $hw.Enabled
                                State                        = $state.ItemState
                                LastModified                 = $cam.LastModified
                                Id                           = $cam.Id
                                IsStarted                    = $state.Status.Started
                                IsMotionDetected             = $state.Status.Motion
                                IsRecording                  = $state.Status.Recording
                                IsInOverflow                 = $state.Status.ErrorOverflow
                                IsInDbRepair                 = $state.Status.DbRepairInProgress
                                ErrorWritingGOP              = $state.Status.ErrorWritingGop
                                ErrorNotLicensed             = $state.Status.ErrorNotLicensed
                                ErrorNoConnection            = $state.Status.ErrorNoConnection
                                StatusTime                   = $state.Status.Time
                                GpsCoordinates               = $cam.GisPoint | ConvertFrom-GisPoint

                                HardwareName                 = $hw.Name
                                HardwareId                   = $hw.Id
                                Model                        = $hw.Model
                                Address                      = $hw.Address
                                Username                     = $hw.UserName
                                Password                     = if ($cache.Passwords.ContainsKey([guid]$hw.Id)) { $cache.Passwords[[guid]$hw.Id] } else { 'NotIncluded' }
                                HTTPSEnabled                 = $hwSettings.HTTPSEnabled -eq 'yes'
                                MAC                          = $hwSettings.MacAddress
                                Firmware                     = $hwSettings.FirmwareVersion

                                DriverFamily                 = $driver.GroupName
                                Driver                       = $driver.Name
                                DriverNumber                 = $driver.Number
                                DriverVersion                = $driver.DriverVersion
                                DriverRevision               = $driver.DriverRevision

                                RecorderName                 = $rec.Name
                                RecorderUri                  = $rec.ActiveWebServerUri, $rec.WebServerUri | Where-Object { ![string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
                                RecorderId                   = $rec.Id

                                LiveStream                   = $liveStreamName
                                LiveStreamDescription        = $liveStreamUsage.Name
                                LiveStreamMode               = $liveStreamUsage.LiveMode
                                ConfiguredLiveResolution     = $liveStreamSettings.Resolution, $liveStreamSettings.StreamProperty | Where-Object { ![string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
                                ConfiguredLiveCodec          = $liveStreamSettings.Codec
                                ConfiguredLiveFPS            = $liveStreamSettings.FPS, $liveStreamSettings.FrameRate | Where-Object { ![string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
                                CurrentLiveResolution        = if ($null -eq $liveStreamStats) { 'Unavailable' } else { "{0}x{1}" -f $liveStreamStats.ImageResolution.Width, $liveStreamStats.ImageResolution.Height }
                                CurrentLiveCodec             = if ($null -eq $liveStreamStats) { 'Unavailable' } else { $liveStreamStats.VideoFormat }
                                CurrentLiveFPS               = if ($null -eq $liveStreamStats) { 'Unavailable' } else { $liveStreamStats.FPS -as [int] }
                                CurrentLiveBitrate           = if ($null -eq $liveStreamStats) { 'Unavailable' } else { (($liveStreamStats.BPS -as [int]) / 1MB).ToString('N1') }

                                RecordedStream               = $recordedStreamName
                                RecordedStreamDescription    = $recordedStreamUsage.Name
                                RecordedStreamMode           = $liveStreamUsage.LiveMode
                                ConfiguredRecordedResolution = $recordedStreamSettings.Resolution, $recordedStreamSettings.StreamProperty | Where-Object { ![string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
                                ConfiguredRecordedCodec      = $recordedStreamSettings.Codec
                                ConfiguredRecordedFPS        = $recordedStreamSettings.FPS, $recordedStreamSettings.FrameRate | Where-Object { ![string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
                                CurrentRecordedResolution    = if ($null -eq $recordedStreamStats) { 'Unavailable' } else { "{0}x{1}" -f $recordedStreamStats.ImageResolution.Width, $recordedStreamStats.ImageResolution.Height }
                                CurrentRecordedCodec         = if ($null -eq $recordedStreamStats) { 'Unavailable' } else { $recordedStreamStats.VideoFormat }
                                CurrentRecordedFPS           = if ($null -eq $recordedStreamStats) { 'Unavailable' } else { $recordedStreamStats.FPS -as [int] }
                                CurrentRecordedBitrate       = if ($null -eq $recordedStreamStats) { 'Unavailable' } else { (($recordedStreamStats.BPS -as [int]) / 1MB).ToString('N1') }

                                RecordingEnabled             = $cam.RecordingEnabled
                                RecordKeyframesOnly          = $cam.RecordKeyframesOnly
                                RecordOnRelatedDevices       = $cam.RecordOnRelatedDevices
                                PrebufferEnabled             = $cam.PrebufferEnabled
                                PrebufferSeconds             = $cam.PrebufferSeconds
                                PrebufferInMemory            = $cam.PrebufferInMemory

                                RecordingStorageName         = $storage.Name
                                RecordingPath                = [io.path]::Combine($storage.DiskPath, $storage.Id)
                                ExpectedRetentionDays        = ($storage | Get-VmsStorageRetention).TotalDays
                                PercentRecordedOneWeek       = if ($IncludeRecordingStats) { $cache.RecordingStats[$id].PercentRecorded -as [double] } else { 'NotIncluded' }

                                MediaDatabaseBegin           = if ($null -eq $cache.PlaybackInfo[$id].Begin) { if ($IncludeRetentionInfo) { 'Unavailable' } else { 'NotIncluded' } } else { $cache.PlaybackInfo[$id].Begin }
                                MediaDatabaseEnd             = if ($null -eq $cache.PlaybackInfo[$id].End) { if ($IncludeRetentionInfo) { 'Unavailable' } else { 'NotIncluded' } } else { $cache.PlaybackInfo[$id].End }
                                UsedSpaceInGB                = if ($null -eq $state.UsedSpaceInBytes) { 'Unavailable' } else { ($state.UsedSpaceInBytes / 1GB).ToString('N2') }

                            }
                            if ($IncludeRetentionInfo) {
                                $obj.ActualRetentionDays     = ($cache.PlaybackInfo[$id].End - $cache.PlaybackInfo[$id].Begin).TotalDays
                                $obj.MeetsRetentionPolicy    = $obj.ActualRetentionDays -gt $obj.ExpectedRetentionDays
                                $obj.MediaDatabaseBegin      = $cache.PlaybackInfo[$id].Begin
                                $obj.MediaDatabaseEnd        = $cache.PlaybackInfo[$id].End
                            }

                            $obj.MotionEnabled = $motion.Enabled
                            $obj.MotionKeyframesOnly = $motion.KeyframesOnly
                            $obj.MotionProcessTime = $motion.ProcessTime
                            $obj.MotionManualSensitivityEnabled = $motion.ManualSensitivityEnabled
                            $obj.MotionManualSensitivity = $motion.ManualSensitivity
                            $obj.MotionMetadataEnabled = $motion.GenerateMotionMetadata
                            $obj.MotionExcludeRegions = $motion.UseExcludeRegions
                            $obj.MotionHardwareAccelerationMode = $motion.HardwareAccelerationMode

                            $obj.PrivacyMaskEnabled = ($cam.PrivacyProtectionFolder.PrivacyProtections | Select-Object -First 1).Enabled -eq $true

                            if ($IncludeSnapshots) {
                                $obj.Snapshot = $cache.Snapshots[$id]
                            }
                            Write-Output ([pscustomobject]$obj)
                        }
                    }
                    catch {
                        Write-Error $_
                    }
                }
            }
        }
        finally {
            if ($jobRunner) {
                $jobRunner.Dispose()
            }
        }
    }
}
function Add-VmsArchiveStorage {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([VideoOS.Platform.ConfigurationItems.ArchiveStorage])]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.Storage]
        $Storage,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Name,

        [Parameter()]
        [string]
        $Description,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Path,

        [Parameter()]
        [ValidateScript({
            if ($_ -lt [timespan]::FromMinutes(60)) {
                throw "Retention must be greater than or equal to one hour"
            }
            if ($_ -gt [timespan]::FromMinutes([int]::MaxValue)) {
                throw "Retention must be less than or equal to $([int]::MaxValue) minutes."
            }
            $true
        })]
        [timespan]
        $Retention,

        [Parameter(Mandatory)]
        [ValidateRange(1, [int]::MaxValue)]
        [int]
        $MaximumSizeMB,

        [Parameter()]
        [switch]
        $ReduceFramerate,

        [Parameter()]
        [ValidateRange(0.00028, 100)]
        [double]
        $TargetFramerate = 5
    )

    process {
        $archiveFolder = $Storage.ArchiveStorageFolder
        if ($PSCmdlet.ShouldProcess("Recording storage '$($Storage.Name)'", "Add new archive storage named '$($Name)' with retention of $($Retention.TotalHours) hours and a maximum size of $($MaximumSizeMB) MB")) {
            try {
                $taskInfo = $archiveFolder.AddArchiveStorage($Name, $Description, $Path, $TargetFrameRate, $Retention.TotalMinutes, $MaximumSizeMB)
                if ($taskInfo.State -ne [videoos.platform.configurationitems.stateenum]::Success) {
                    Write-Error -Message $taskInfo.ErrorText
                    return
                }

                $archive = [VideoOS.Platform.ConfigurationItems.ArchiveStorage]::new((Get-ManagementServer).ServerId, $taskInfo.Path)

                if ($ReduceFramerate) {
                    $invokeInfo = $archive.SetFramerateReductionArchiveStorage()
                    $invokeInfo.SetProperty('FramerateReductionEnabled', 'True')
                    [void]$invokeInfo.ExecuteDefault()
                }

                $storage.ClearChildrenCache()
                Write-Output $archive
            }
            catch {
                Write-Error $_
                return
            }
        }
    }
}
function Add-VmsStorage {
    [CmdletBinding(DefaultParameterSetName = 'WithoutEncryption', SupportsShouldProcess)]
    [OutputType([VideoOS.Platform.ConfigurationItems.Storage])]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'WithoutEncryption')]
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'WithEncryption')]
        [VideoOS.Platform.ConfigurationItems.RecordingServer]
        $RecordingServer,

        [Parameter(Mandatory, ParameterSetName = 'WithoutEncryption')]
        [Parameter(Mandatory, ParameterSetName = 'WithEncryption')]
        [ValidateNotNullOrEmpty()]
        [string]
        $Name,

        [Parameter(ParameterSetName = 'WithoutEncryption')]
        [Parameter(ParameterSetName = 'WithEncryption')]
        [string]
        $Description,

        [Parameter(Mandatory, ParameterSetName = 'WithoutEncryption')]
        [Parameter(Mandatory, ParameterSetName = 'WithEncryption')]
        [ValidateNotNullOrEmpty()]
        [string]
        $Path,

        [Parameter(ParameterSetName = 'WithoutEncryption')]
        [Parameter(ParameterSetName = 'WithEncryption')]
        [ValidateScript({
            if ($_ -lt [timespan]::FromMinutes(60)) {
                throw "Retention must be greater than or equal to one hour"
            }
            if ($_ -gt [timespan]::FromMinutes([int]::MaxValue)) {
                throw "Retention must be less than or equal to $([int]::MaxValue) minutes."
            }
            $true
        })]
        [timespan]
        $Retention,

        [Parameter(Mandatory, ParameterSetName = 'WithoutEncryption')]
        [Parameter(Mandatory, ParameterSetName = 'WithEncryption')]
        [ValidateRange(1, [int]::MaxValue)]
        [int]
        $MaximumSizeMB,

        [Parameter(ParameterSetName = 'WithoutEncryption')]
        [Parameter(ParameterSetName = 'WithEncryption')]
        [switch]
        $Default,

        [Parameter(ParameterSetName = 'WithoutEncryption')]
        [Parameter(ParameterSetName = 'WithEncryption')]
        [switch]
        $EnableSigning,

        [Parameter(Mandatory, ParameterSetName = 'WithEncryption')]
        [ValidateSet('Light', 'Strong', IgnoreCase = $false)]
        [string]
        $EncryptionMethod,

        [Parameter(Mandatory, ParameterSetName = 'WithEncryption')]
        [securestring]
        $Password
    )

    process {
        $storageFolder = $RecordingServer.StorageFolder
        if ($PSCmdlet.ShouldProcess("Recording Server '$($RecordingServer.Name)' at $($RecordingServer.HostName)", "Add new storage named '$($Name)' with retention of $($Retention.TotalHours) hours and a maximum size of $($MaximumSizeMB) MB")) {
            try {
                $taskInfo = $storageFolder.AddStorage($Name, $Description, $Path, $EnableSigning, $Retention.TotalMinutes, $MaximumSizeMB)
                if ($taskInfo.State -ne [videoos.platform.configurationitems.stateenum]::Success) {
                    Write-Error -Message $taskInfo.ErrorText
                    return
                }
            }
            catch {
                Write-Error $_
                return
            }

            $storage = [VideoOS.Platform.ConfigurationItems.Storage]::new((Get-ManagementServer).ServerId, $taskInfo.Path)
        }

        if ($PSCmdlet.ParameterSetName -eq 'WithEncryption' -and $PSCmdlet.ShouldProcess("Recording Storage '$Name'", "Enable '$EncryptionMethod' Encryption")) {
            try {
                $invokeResult = $storage.EnableEncryption($Password, $EncryptionMethod)
                if ($invokeResult.State -ne [videoos.platform.configurationitems.stateenum]::Success) {
                    throw $invokeResult.ErrorText
                }

                $storage = [VideoOS.Platform.ConfigurationItems.Storage]::new((Get-ManagementServer).ServerId, $taskInfo.Path)
            }
            catch {
                [void]$storageFolder.RemoveStorage($taskInfo.Path)
                Write-Error $_
                return
            }
        }

        if ($Default -and $PSCmdlet.ShouldProcess("Recording Storage '$Name'", "Set as default storage configuration")) {
            try {
                $invokeResult = $storage.SetStorageAsDefault()
                if ($invokeResult.State -ne [videoos.platform.configurationitems.stateenum]::Success) {
                    throw $invokeResult.ErrorText
                }

                $storage = [VideoOS.Platform.ConfigurationItems.Storage]::new((Get-ManagementServer).ServerId, $taskInfo.Path)
            }
            catch {
                [void]$storageFolder.RemoveStorage($taskInfo.Path)
                Write-Error $_
                return
            }
        }

        if (!$PSBoundParameters.ContainsKey('WhatIf')) {
            Write-Output $storage
        }
    }
}
function Get-VmsArchiveStorage {
    [CmdletBinding()]
    [OutputType([VideoOS.Platform.ConfigurationItems.ArchiveStorage])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.Storage]
        $Storage,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [SupportsWildcards()]
        [string]
        $Name = '*'
    )

    process {
        $storagesMatched = 0
        $Storage.ArchiveStorageFolder.ArchiveStorages | ForEach-Object {
            if ($_.Name -like $Name) {
                $storagesMatched++
                Write-Output $_
            }
        }

        if ($storagesMatched -eq 0 -and -not [System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($Name)) {
            Write-Error "No recording storages found matching the name '$Name'"
        }
    }
}
function Get-VmsStorage {
    [CmdletBinding()]
    [OutputType([VideoOS.Platform.ConfigurationItems.Storage])]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'FromName')]
        [VideoOS.Platform.ConfigurationItems.RecordingServer]
        $RecordingServer,

        [Parameter(ParameterSetName = 'FromName')]
        [ValidateNotNullOrEmpty()]
        [SupportsWildcards()]
        [string]
        $Name = '*',

        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'FromPath')]
        [ValidateScript({
            if ($_ -match 'Storage\[.{36}\]') {
                $true
            }
            else {
                throw "Invalid storage item path. Expected format: Storage[$([guid]::NewGuid())]"
            }
        })]
        [Alias('RecordingStorage', 'Path')]
        [string]
        $ItemPath
    )

    process {
        switch ($PSCmdlet.ParameterSetName) {
            'FromName' {
                $storagesMatched = 0
                $RecordingServer.StorageFolder.Storages | ForEach-Object {
                    if ($_.Name -like $Name) {
                        $storagesMatched++
                        Write-Output $_
                    }
                }

                if ($storagesMatched -eq 0 -and -not [System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($Name)) {
                    Write-Error "No recording storages found matching the name '$Name'"
                }
            }
            'FromPath' {
                [VideoOS.Platform.ConfigurationItems.Storage]::new((Get-ManagementServer).ServerId, $ItemPath)
            }
            Default {
                throw "ParameterSetName $($PSCmdlet.ParameterSetName) not implemented"
            }
        }
    }
}
function Remove-VmsArchiveStorage {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'ByName')]
        [VideoOS.Platform.ConfigurationItems.Storage]
        $Storage,

        [Parameter(Mandatory, ParameterSetName = 'ByName')]
        [ValidateNotNullOrEmpty()]
        [string]
        $Name,

        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'ByStorage')]
        [VideoOS.Platform.ConfigurationItems.ArchiveStorage]
        $ArchiveStorage
    )

    process {
        switch ($PSCmdlet.ParameterSetName) {
            'ByName' {
                foreach ($archiveStorage in $Storage | Get-VmsArchiveStorage -Name $Name) {
                    $archiveStorage | Remove-VmsArchiveStorage
                }
            }

            'ByStorage' {
                $recorder = [VideoOS.Platform.ConfigurationItems.RecordingServer]::new((Get-ManagementServer).ServerId, $Storage.ParentItemPath)
                $storage = [VideoOS.Platform.ConfigurationItems.Storage]::new((Get-ManagementServer).ServerId, $ArchiveStorage.ParentItemPath)
                if ($PSCmdlet.ShouldProcess("Recording server $($recorder.Name)", "Delete archive $($ArchiveStorage.Name) from $($storage.Name)")) {
                    $folder = [VideoOS.Platform.ConfigurationItems.ArchiveStorageFolder]::new((Get-ManagementServer).ServerId, $ArchiveStorage.ParentPath)
                    [void]$folder.RemoveArchiveStorage($ArchiveStorage.Path)
                }
            }
            Default {
                throw 'Unknown parameter set'
            }
        }
    }
}
function Remove-VmsStorage {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'ByName')]
        [VideoOS.Platform.ConfigurationItems.RecordingServer]
        $RecordingServer,

        [Parameter(Mandatory, ParameterSetName = 'ByName')]
        [ValidateNotNullOrEmpty()]
        [string]
        $Name,

        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'ByStorage')]
        [VideoOS.Platform.ConfigurationItems.Storage]
        $Storage
    )

    process {
        switch ($PSCmdlet.ParameterSetName) {
            'ByName' {
                foreach ($vmsStorage in $RecordingServer | Get-VmsStorage -Name $Name) {
                    $vmsStorage | Remove-VmsStorage
                }
            }

            'ByStorage' {
                $recorder = [VideoOS.Platform.ConfigurationItems.RecordingServer]::new((Get-ManagementServer).ServerId, $Storage.ParentItemPath)
                if ($PSCmdlet.ShouldProcess("Recording server $($recorder.Name)", "Delete $($Storage.Name) and all archives")) {
                    $folder = [VideoOS.Platform.ConfigurationItems.StorageFolder]::new((Get-ManagementServer).ServerId, $Storage.ParentPath)
                    [void]$folder.RemoveStorage($Storage.Path)
                }
            }
            Default {
                throw 'Unknown parameter set'
            }
        }
    }
}
function ConvertFrom-ConfigurationApiProperties {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.ConfigurationApiProperties]
        $Properties,

        [Parameter()]
        [switch]
        $UseDisplayNames
    )

    process {
        $languageId = (Get-Culture).Name
        $result = @{}
        foreach ($key in $Properties.Keys) {
            if ($key -notmatch '^.+/(?<Key>.+)/(?:[0-9A-F\-]{36})$') {
                Write-Warning "Failed to parse property with key name '$key'"
                continue
            }
            $propertyInfo = $Properties.GetValueTypeInfoCollection($key)
            $propertyValue = $Properties.GetValue($key)

            if ($UseDisplayNames) {
                $valueTypeInfo = $propertyInfo | Where-Object Value -eq $propertyValue
                $displayName = $valueTypeInfo.Name
                if ($propertyInfo.Count -gt 0 -and $displayName -and $displayName -notin @('true', 'false', 'MinValue', 'MaxValue', 'StepValue')) {
                    if ($valueTypeInfo.TranslationId -and $languageId -and $languageId -ne 'en-US') {
                        $translatedName = (Get-Translations -LanguageId $languageId).($valueTypeInfo.TranslationId)
                        if (![string]::IsNullOrWhiteSpace($translatedName)) {
                            $displayName = $translatedName
                        }
                    }
                    $result[$Matches.Key] = $displayName
                }
                else {
                    $result[$Matches.Key] = $propertyValue
                }
            }
            else {
                $result[$Matches.Key] = $propertyValue
            }
        }

        Write-Output $result
    }
}
function ConvertFrom-GisPoint {
    [CmdletBinding()]
    [OutputType([system.device.location.geocoordinate])]
    param (
        # Specifies the GisPoint value to convert to a GeoCoordinate. Milestone stores GisPoint data in the format "POINT ([longitude] [latitude])" or "POINT EMPTY".
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ValueFromPipeline)]
        [string]
        $GisPoint
    )

    process {
        if ($GisPoint -eq 'POINT EMPTY') {
            Write-Output ([system.device.location.geocoordinate]::Unknown)
        }
        else {
            $temp = $GisPoint.Substring(7, $GisPoint.Length - 8)
            $long, $lat, $null = $temp -split ' '
            Write-Output ([system.device.location.geocoordinate]::new($lat, $long))
        }
    }
}
function ConvertFrom-Snapshot {
    [CmdletBinding()]
    [OutputType([system.drawing.image])]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [Alias('Bytes')]
        [byte[]]
        $Content
    )

    process {
        if ($null -eq $Content -or $Content.Length -eq 0) {
            return $null
        }
        $ms = [io.memorystream]::new($Content)
        Write-Output ([system.drawing.image]::FromStream($ms))
    }
}
function ConvertTo-GisPoint {
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'FromGeoCoordinate')]
        [system.device.location.geocoordinate]
        $Coordinate,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'FromValues')]
        [double]
        $Latitude,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'FromValues')]
        [double]
        $Longitude,

        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'FromValues')]
        [double]
        $Altitude = [double]::NaN,

        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'FromString')]
        [string]
        $Coordinates
    )

    process {

        switch ($PsCmdlet.ParameterSetName) {
            'FromValues' {
                # do nothing?
                break
            }

            'FromGeoCoordinate' {
                $Latitude = $Coordinate.Latitude
                $Longitude = $Coordinate.Longitude
                $Altitude = $Coordinate.Altitude
                break
            }

            'FromString' {
                $values = $Coordinates -split ',' | Foreach-Object {
                    [double]$_.Trim()
                }
                if ($values.Count -lt 2 -or $values.Count -gt 3) {
                    Write-Error "Failed to parse coordinates into latitude, longitude and optional altitude."
                    return
                }
                $Latitude = $values[0]
                $Longitude = $values[1]
                if ($values.Count -gt 2) {
                    $Altitude = $values[2]
                }
                break
            }
        }

        if ([double]::IsNan($Altitude)) {
            Write-Output ('POINT ({0} {1})' -f $Longitude, $Latitude)
        }
        else {
            Write-Output ('POINT ({0} {1} {2})' -f $Longitude, $Latitude, $Altitude)
        }
    }
}
function Get-BankTable {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $Path,
        [Parameter()]
        [string[]]
        $DeviceId,
        [Parameter()]
        [DateTime]
        $StartTime = [DateTime]::MinValue,
        [Parameter()]
        [DateTime]
        $EndTime = [DateTime]::MaxValue.AddHours(-1)

    )

    process {
        $di = [IO.DirectoryInfo]$Path
        foreach ($table in $di.EnumerateDirectories()) {
            if ($table.Name -match "^(?<id>[0-9a-fA-F\-]{36})(_(?<tag>\w+)_(?<endTime>\d\d\d\d-\d\d-\d\d_\d\d-\d\d-\d\d).*)?") {
                $tableTimestamp = if ($null -eq $Matches["endTime"]) { (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss") } else { $Matches["endTime"] }
                $timestamp = [DateTime]::ParseExact($tableTimestamp, "yyyy-MM-dd_HH-mm-ss", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeLocal)
                if ($timestamp -lt $StartTime -or $timestamp -gt $EndTime.AddHours(1)) {
                    # Timestamp of table is outside the requested timespan
                    continue
                }
                if ($null -ne $DeviceId -and [cultureinfo]::InvariantCulture.CompareInfo.IndexOf($DeviceId, $Matches["id"], [System.Globalization.CompareOptions]::IgnoreCase) -eq -1) {
                    # Device ID for table is not requested
                    continue
                }
                [pscustomobject]@{
                    DeviceId = [Guid]$Matches["id"]
                    EndTime = $timestamp
                    Tag = $Matches["tag"]
                    IsLiveTable = $null -eq $Matches["endTime"]
                    Path = $table.FullName
                }
            }
        }
    }
}
function Get-ConfigurationItemProperty {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.ConfigurationApi.ClientService.ConfigurationItem]
        [ValidateNotNullOrEmpty()]
        $InputObject,
        [Parameter(Mandatory)]
        [string]
        [ValidateNotNullOrEmpty()]
        $Key
    )

    process {
        $property = $InputObject.Properties | Where-Object Key -eq $Key
        if ($null -eq $property) {
            Write-Error -Message "Key '$Key' not found on configuration item $($InputObject.Path)" -TargetObject $InputObject -Category InvalidArgument
            return
        }
        $property.Value
    }
}
function Get-StreamProperties {
    [CmdletBinding()]
    [OutputType([VideoOS.ConfigurationApi.ClientService.Property[]])]
    param (
        # Specifies the camera to retrieve stream properties for
        [Parameter(ValueFromPipeline, Mandatory, ParameterSetName = 'ByName')]
        [Parameter(ValueFromPipeline, Mandatory, ParameterSetName = 'ByNumber')]
        [VideoOS.Platform.ConfigurationItems.Camera]
        $Camera,

        # Specifies a StreamUsageChildItem from Get-Stream
        [Parameter(ParameterSetName = 'ByName')]
        [ValidateNotNullOrEmpty()]
        [string]
        $StreamName,

        # Specifies the stream number starting from 0. For example, "Video stream 1" is usually in the 0'th position in the StreamChildItems collection.
        [Parameter(ParameterSetName = 'ByNumber')]
        [ValidateRange(0, [int]::MaxValue)]
        [int]
        $StreamNumber
    )

    process {
        switch ($PSCmdlet.ParameterSetName) {
            'ByName' {
                $stream = (Get-ConfigurationItem -Path "DeviceDriverSettings[$($Camera.Id)]").Children | Where-Object { $_.ItemType -eq 'Stream' -and $_.DisplayName -like $StreamName }
                if ($null -eq $stream -and ![system.management.automation.wildcardpattern]::ContainsWildcardCharacters($StreamName)) {
                    Write-Error "No streams found on $($Camera.Name) matching the name '$StreamName'"
                    return
                }
                foreach ($obj in $stream) {
                    Write-Output $obj.Properties
                }
            }
            'ByNumber' {
                $streams = (Get-ConfigurationItem -Path "DeviceDriverSettings[$($Camera.Id)]").Children | Where-Object { $_.ItemType -eq 'Stream' }
                if ($StreamNumber -lt $streams.Count) {
                    Write-Output ($streams[$StreamNumber].Properties)
                }
                else {
                    Write-Error "There are $($streams.Count) streams available on the camera and stream number $StreamNumber does not exist. Remember to index the streams from zero."
                }
            }
            Default {}
        }
    }
}
function Get-ValueDisplayName {
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory, ParameterSetName = 'ConfigurationApi')]
        [VideoOS.ConfigurationApi.ClientService.Property[]]
        $PropertyList,

        [Parameter(Mandatory, ParameterSetName = 'StrongTypes')]
        [VideoOS.Platform.ConfigurationItems.ConfigurationApiProperties]
        $Properties,

        [Parameter(Mandatory, ParameterSetName = 'ConfigurationApi')]
        [Parameter(Mandatory, ParameterSetName = 'StrongTypes')]
        [string[]]
        $PropertyName,

        [Parameter()]
        [string]
        $DefaultValue = 'NotAvailable'
    )

    process {
        $value = $DefaultValue
        if ($null -eq $PropertyList -or $PropertyList.Count -eq 0) {
            return $value
        }

        $selectedProperty = $null
        foreach ($property in $PropertyList) {
            foreach ($name in $PropertyName) {
                if ($property.Key -like "*/$name/*") {
                    $selectedProperty = $property
                    break
                }
            }
            if ($null -ne $selectedProperty) { break }
        }
        if ($null -ne $selectedProperty) {
            $value = $selectedProperty.Value
            if ($selectedProperty.ValueType -eq 'Enum') {
                $displayName = ($selectedProperty.ValueTypeInfos | Where-Object Value -eq $selectedProperty.Value).Name
                if (![string]::IsNullOrWhiteSpace($displayName)) {
                    $value = $displayName
                }
            }
        }
        Write-Output $value
    }
}
function Install-StableFPS {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $Source = "C:\Program Files\Milestone\MIPSDK\Tools\StableFPS",
        [Parameter()]
        [int]
        [ValidateRange(1, 200)]
        $Cameras = 32,
        [Parameter()]
        [int]
        [ValidateRange(1, 5)]
        $Streams = 1,
        [Parameter()]
        [string]
        $DevicePackPath
    )

    begin {
        $Elevated = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if (!$Elevated) {
            throw "Elevation is required for this command to work properly. Consider re-launching PowerShell by right-clicking and running as Administrator."
        }
        if (!(Test-Path (Join-Path $Source "StableFPS_DATA"))) {
            throw "Path not found: $((Join-Path $Source "StableFPS_DATA"))"
        }
        if (!(Test-Path (Join-Path $Source "vLatest"))) {
            throw "Path not found: $((Join-Path $Source "vLatest"))"
        }
    }

    process {
        $serviceStopped = $false
        try {
            $dpPath = if ([string]::IsNullOrWhiteSpace($DevicePackPath)) { (Get-RecorderConfig).DevicePackPath } else { $DevicePackPath }
            if (!(Test-Path $dpPath)) {
                throw "DevicePackPath not valid"
            }
            if ([string]::IsNullOrWhiteSpace($DevicePackPath)) {
                $service = Get-Service "Milestone XProtect Recording Server"
                if ($service.Status -eq [System.ServiceProcess.ServiceControllerStatus]::Running) {
                    $service | Stop-Service -Force
                    $serviceStopped = $true
                }
            }

            $srcData = Join-Path $Source "StableFPS_Data"
            $srcDriver = Join-Path $Source "vLatest"
            Copy-Item $srcData -Destination $dpPath -Container -Recurse -Force
            Copy-Item "$srcDriver\*" -Destination $dpPath -Recurse -Force

            $tempXml = Join-Path $dpPath "resources\StableFPS_TEMP.xml"
            $newXml = Join-Path $dpPath "resources\StableFPS.xml"
            $content = Get-Content $tempXml -Raw
            $content = $content.Replace("{CAM_NUM_REQUESTED}", $Cameras)
            $content = $content.Replace("{STREAM_NUM_REQUESTED}", $Streams)
            $content | Set-Content $newXml
            Remove-Item $tempXml
        }
        catch {
            throw
        }
        finally {
            if ($serviceStopped -and $null -ne $service) {
                $service.Refresh()
                $service.Start()
            }
        }
    }
}
function Invoke-ServerConfigurator {
    [CmdletBinding()]
    param(
        # Enable encryption for the CertificateGroup specified
        [Parameter(ParameterSetName = 'EnableEncryption', Mandatory)]
        [switch]
        $EnableEncryption,

        # Disable encryption for the CertificateGroup specified
        [Parameter(ParameterSetName = 'DisableEncryption', Mandatory)]
        [switch]
        $DisableEncryption,

        # Specifies the CertificateGroup [guid] identifying which component for which encryption
        # should be enabled or disabled
        [Parameter(ParameterSetName = 'EnableEncryption', Mandatory)]
        [Parameter(ParameterSetName = 'DisableEncryption', Mandatory)]
        [guid]
        $CertificateGroup,

        # Specifies the thumbprint of the certificate to be used to encrypt communications with the
        # component designated by the CertificateGroup id.
        [Parameter(ParameterSetName = 'EnableEncryption', Mandatory)]
        [string]
        $Thumbprint,

        # List the available certificate groups on the local machine. Output will be a [hashtable]
        # where the keys are the certificate group names (which may contain spaces) and the values
        # are the associated [guid] id's.
        [Parameter(ParameterSetName = 'ListCertificateGroups')]
        [switch]
        $ListCertificateGroups,

        # Register all local components with the optionally specified AuthAddress. If no
        # AuthAddress is provided, the last-known address will be used.
        [Parameter(ParameterSetName = 'Register', Mandatory)]
        [switch]
        $Register,

        # Specifies the address of the Authorization Server which is usually the Management Server
        # address. A [uri] value is expected, but only the URI host value will be used. The scheme
        # and port will be inferred based on whether encryption is enabled/disabled and is fixed to
        # port 80/443 as this is how Server Configurator is currently designed.
        [Parameter(ParameterSetName = 'Register')]
        [uri]
        $AuthAddress,

        # Specifies the path to the Server Configurator utility. Omit this path and the path will
        # be discovered using Get-RecorderConfig or Get-ManagementServerConfig by locating the
        # installation path of the Management Server or Recording Server and assuming the Server
        # Configurator is located in the same path.
        [Parameter()]
        [string]
        $Path,

        # Specifies that the standard output from the Server Configurator utility should be written
        # after the operation is completed. The output will include the following properties:
        # - StandardOutput
        # - StandardError
        # - ExitCode
        [Parameter(ParameterSetName = 'EnableEncryption')]
        [Parameter(ParameterSetName = 'DisableEncryption')]
        [Parameter(ParameterSetName = 'Register')]
        [switch]
        $PassThru
    )

    process {
        $exePath = $Path
        if ([string]::IsNullOrWhiteSpace($exePath)) {
            # Find ServerConfigurator.exe by locating either the Management Server or Recording Server installation path
            $configurationInfo = try {
                Get-ManagementServerConfig
            }
            catch {
                try {
                    Get-RecorderConfig
                }
                catch {
                    $null
                }
            }
            if ($null -eq $configurationInfo) {
                Write-Error "Could not find a Management Server or Recording Server installation"
                return
            }
            $fileInfo = [io.fileinfo]::new($configurationInfo.InstallationPath)
            $exePath = Join-Path $fileInfo.Directory.Parent.FullName "Server Configurator\serverconfigurator.exe"
            if (-not (Test-Path $exePath)) {
                Write-Error "Expected to find Server Configurator at '$exePath' but failed."
                return
            }
        }


        # Ensure version is 20.3 (2020 R3) or newer
        $fileInfo = [io.fileinfo]::new($exePath)
        if ($fileInfo.VersionInfo.FileVersion -lt [version]"20.3") {
            Write-Error "Invoke-ServerConfigurator requires Milestone version 2020 R3 or newer as this is when command-line options were introduced. Found Server Configurator version $($fileInfo.VersionInfo.FileVersion)"
            return
        }

        $exitCode = @{
            0 = 'Success'
            -1 = 'Unknown error'
            -2 = 'Invalid arguments'
            -3 = 'Invalid argument value'
            -4 = 'Another instance is running'
        }

        # Get Certificate Group list for either display to user or verification
        $output = Get-ProcessOutput -FilePath $exePath -ArgumentList /listcertificategroups
        if ($output.ExitCode -ne 0) {
            Write-Error "Server Configurator exited with code $($output.ExitCode). $($exitCode.($output.ExitCode))."
            Write-Error $output.StandardOutput
            return
        }
        Write-Information $output.StandardOutput
        $groups = @{}
        foreach ($line in $output.StandardOutput -split ([environment]::NewLine)) {
            if ($line -match "Found '(?<groupName>.+)' group with ID = (?<groupId>.{36})") {
                $groups.$($Matches.groupName) = [guid]::Parse($Matches.groupId)
            }
        }


        switch ($PSCmdlet.ParameterSetName) {
            'EnableEncryption' {
                if ($groups.Values -notcontains $CertificateGroup) {
                    Write-Error "CertificateGroup value '$CertificateGroup' not found. Use the ListCertificateGroups switch to discover valid CertificateGroup values"
                    return
                }

                $enableArgs = @('/enableencryption', "/certificategroup=$CertificateGroup", "/thumbprint=$Thumbprint", '/quiet')
                $output = Get-ProcessOutput -FilePath $exePath -ArgumentList $enableArgs
                if ($output.ExitCode -ne 0) {
                    Write-Error "EnableEncryption failed. Server Configurator exited with code $($output.ExitCode). $($exitCode.($output.ExitCode))."
                    Write-Error $output.StandardOutput
                }
            }

            'DisableEncryption' {
                if ($groups.Values -notcontains $CertificateGroup) {
                    Write-Error "CertificateGroup value '$CertificateGroup' not found. Use the ListCertificateGroups switch to discover valid CertificateGroup values"
                    return
                }
                $disableArgs = @('/disableencryption', "/certificategroup=$CertificateGroup", '/quiet')
                $output = Get-ProcessOutput -FilePath $exePath -ArgumentList $disableArgs
                if ($output.ExitCode -ne 0) {
                    Write-Error "EnableEncryption failed. Server Configurator exited with code $($output.ExitCode). $($exitCode.($output.ExitCode))."
                    Write-Error $output.StandardOutput
                }
            }

            'ListCertificateGroups' {
                Write-Output $groups
                return
            }

            'Register' {
                $registerArgs = @('/register', '/quiet')
                if ($PSCmdlet.MyInvocation.BoundParameters -contains 'AuthAddress') {
                    $registerArgs += $AuthAddress.ToString()
                }
                $output = Get-ProcessOutput -FilePath $exePath -ArgumentList $registerArgs
                if ($output.ExitCode -ne 0) {
                    Write-Error "Registration failed. Server Configurator exited with code $($output.ExitCode). $($exitCode.($output.ExitCode))."
                    Write-Error $output.StandardOutput
                }

            }

            Default {
            }
        }

        Write-Information $output.StandardOutput
        if ($PassThru) {
            Write-Output $output
        }
    }
}
function Resize-Image {
    [CmdletBinding()]
    [OutputType([System.Drawing.Image])]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Drawing.Image]
        $Image,

        [Parameter(Mandatory)]
        [int]
        $Height,

        [Parameter()]
        [long]
        $Quality = 95,

        [Parameter()]
        [ValidateSet('BMP', 'JPEG', 'GIF', 'TIFF', 'PNG')]
        [string]
        $OutputFormat,

        [Parameter()]
        [switch]
        $DisposeSource
    )

    process {
        if ($null -eq $Image -or $Image.Width -le 0 -or $Image.Height -le 0) {
            Write-Error 'Cannot resize an invalid image object.'
            return
        }

        [int]$width = $image.Width / $image.Height * $Height
        $bmp = [system.drawing.bitmap]::new($width, $Height)
        $graphics = [system.drawing.graphics]::FromImage($bmp)
        $graphics.InterpolationMode = [system.drawing.drawing2d.interpolationmode]::HighQualityBicubic
        $graphics.DrawImage($Image, 0, 0, $width, $Height)
        $graphics.Dispose()

        try {
            $formatId = if ([string]::IsNullOrWhiteSpace($OutputFormat)) {
                    $Image.RawFormat.Guid
                }
                else {
                    ([system.drawing.imaging.imagecodecinfo]::GetImageEncoders() | Where-Object FormatDescription -eq $OutputFormat).FormatID
                }
            $encoder = [system.drawing.imaging.imagecodecinfo]::GetImageEncoders() | Where-Object FormatID -eq $formatId
            $encoderParameters = [system.drawing.imaging.encoderparameters]::new(1)
            $qualityParameter = [system.drawing.imaging.encoderparameter]::new([system.drawing.imaging.encoder]::Quality, $Quality)
            $encoderParameters.Param[0] = $qualityParameter
            Write-Verbose "Saving resized image as $($encoder.FormatDescription) with $Quality% quality"
            $ms = [io.memorystream]::new()
            $bmp.Save($ms, $encoder, $encoderParameters)
            $resizedImage = [system.drawing.image]::FromStream($ms)
            Write-Output ($resizedImage)
        }
        finally {
            $qualityParameter.Dispose()
            $encoderParameters.Dispose()
            $bmp.Dispose()
            if ($DisposeSource) {
                $Image.Dispose()
            }
        }

    }
}
function Select-Camera {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]
        $Title = "Select Camera(s)",
        [Parameter()]
        [switch]
        $SingleSelect,
        [Parameter()]
        [switch]
        $AllowFolders,
        [Parameter()]
        [switch]
        $AllowServers,
        [Parameter()]
        [switch]
        $RemoveDuplicates,
        [Parameter()]
        [switch]
        $OutputAsItem
    )
    process {
        $items = Select-VideoOSItem -Title $Title -Kind ([VideoOS.Platform.Kind]::Camera) -AllowFolders:$AllowFolders -AllowServers:$AllowServers -SingleSelect:$SingleSelect -FlattenOutput
        $processed = @{}
        if ($RemoveDuplicates) {
            foreach ($item in $items) {
                if ($processed.ContainsKey($item.FQID.ObjectId)) {
                    continue
                }
                $processed.Add($item.FQID.ObjectId, $null)
                if ($OutputAsItem) {
                    Write-Output $item
                }
                else {
                    Get-VmsCamera -Id $item.FQID.ObjectId
                }
            }
        }
        else {
            if ($OutputAsItem) {
                Write-Output $items
            }
            else {
                Write-Output ($items | ForEach-Object { Get-VmsCamera -Id $_.FQID.ObjectId })
            }
        }
    }
}
function Select-VideoOSItem {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $Title = "Select Item(s)",
        [Parameter()]
        [guid[]]
        $Kind,
        [Parameter()]
        [VideoOS.Platform.Admin.Category[]]
        $Category,
        [Parameter()]
        [switch]
        $SingleSelect,
        [Parameter()]
        [switch]
        $AllowFolders,
        [Parameter()]
        [switch]
        $AllowServers,
        [Parameter()]
        [switch]
        $KindUserSelectable,
        [Parameter()]
        [switch]
        $CategoryUserSelectable,
        [Parameter()]
        [switch]
        $FlattenOutput,
        [Parameter()]
        [switch]
        $HideGroupsTab,
        [Parameter()]
        [switch]
        $HideServerTab
    )

    process {
        $form = [MilestonePSTools.UI.CustomItemPickerForm]::new();
        $form.KindFilter = $Kind
        $form.CategoryFilter = $Category
        $form.AllowFolders = $AllowFolders
        $form.AllowServers = $AllowServers
        $form.KindUserSelectable = $KindUserSelectable
        $form.CategoryUserSelectable = $CategoryUserSelectable
        $form.SingleSelect = $SingleSelect
        $form.GroupTabVisable = -not $HideGroupsTab
        $form.ServerTabVisable = -not $HideServerTab
        $form.Icon = [System.Drawing.Icon]::FromHandle([VideoOS.Platform.UI.Util]::ImageList.Images[[VideoOS.Platform.UI.Util]::SDK_GeneralIx].GetHicon())
        $form.Text = $Title
        $form.TopMost = $true
        $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $form.BringToFront()
        $form.Activate()

        if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            if ($FlattenOutput) {
                Write-Output $form.ItemsSelectedFlattened
            }
            else {
                Write-Output $form.ItemsSelected
            }
        }
    }
}
function Set-ConfigurationItemProperty {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.ConfigurationApi.ClientService.ConfigurationItem]
        [ValidateNotNullOrEmpty()]
        $InputObject,
        [Parameter(Mandatory)]
        [string]
        [ValidateNotNullOrEmpty()]
        $Key,
        [Parameter(Mandatory)]
        [string]
        [ValidateNotNullOrEmpty()]
        $Value,
        [Parameter()]
        [switch]
        $PassThru
    )

    process {
        $property = $InputObject.Properties | Where-Object Key -eq $Key
        if ($null -eq $property) {
            Write-Error -Message "Key '$Key' not found on configuration item $($InputObject.Path)" -TargetObject $InputObject -Category InvalidArgument
            return
        }
        $property.Value = $Value
        if ($PassThru) {
            $InputObject
        }
    }
}
<#
Functions in this module are written as independent PS1 files, and to improve module load time they
are "comiled" into this PSM1 file. If you're looking at this file prior to build, now you know how
all the functions will be loaded later. If you're looking at this file after build, now you know
why this file has so many lines :)
#>

$script:Messages = @{}
Import-LocalizedData -BindingVariable 'script:Messages' -FileName 'messages'
Export-ModuleMember -Cmdlet * -Alias * -Function *


# SIG # Begin signature block
# MIId+QYJKoZIhvcNAQcCoIId6jCCHeYCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU2yqx5l3Zgd8Kzp16gdmMYqcq
# 6PWgghglMIIE/jCCA+agAwIBAgIQDUJK4L46iP9gQCHOFADw3TANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgVGltZXN0YW1waW5nIENBMB4XDTIxMDEwMTAwMDAwMFoXDTMxMDEw
# NjAwMDAwMFowSDELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMu
# MSAwHgYDVQQDExdEaWdpQ2VydCBUaW1lc3RhbXAgMjAyMTCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBAMLmYYRnxYr1DQikRcpja1HXOhFCvQp1dU2UtAxQ
# tSYQ/h3Ib5FrDJbnGlxI70Tlv5thzRWRYlq4/2cLnGP9NmqB+in43Stwhd4CGPN4
# bbx9+cdtCT2+anaH6Yq9+IRdHnbJ5MZ2djpT0dHTWjaPxqPhLxs6t2HWc+xObTOK
# fF1FLUuxUOZBOjdWhtyTI433UCXoZObd048vV7WHIOsOjizVI9r0TXhG4wODMSlK
# XAwxikqMiMX3MFr5FK8VX2xDSQn9JiNT9o1j6BqrW7EdMMKbaYK02/xWVLwfoYer
# vnpbCiAvSwnJlaeNsvrWY4tOpXIc7p96AXP4Gdb+DUmEvQECAwEAAaOCAbgwggG0
# MA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsG
# AQUFBwMIMEEGA1UdIAQ6MDgwNgYJYIZIAYb9bAcBMCkwJwYIKwYBBQUHAgEWG2h0
# dHA6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAfBgNVHSMEGDAWgBT0tuEgHf4prtLk
# YaWyoiWyyBc1bjAdBgNVHQ4EFgQUNkSGjqS6sGa+vCgtHUQ23eNqerwwcQYDVR0f
# BGowaDAyoDCgLoYsaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJl
# ZC10cy5jcmwwMqAwoC6GLGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFz
# c3VyZWQtdHMuY3JsMIGFBggrBgEFBQcBAQR5MHcwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBPBggrBgEFBQcwAoZDaHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0U0hBMkFzc3VyZWRJRFRpbWVzdGFtcGluZ0NB
# LmNydDANBgkqhkiG9w0BAQsFAAOCAQEASBzctemaI7znGucgDo5nRv1CclF0CiNH
# o6uS0iXEcFm+FKDlJ4GlTRQVGQd58NEEw4bZO73+RAJmTe1ppA/2uHDPYuj1UUp4
# eTZ6J7fz51Kfk6ftQ55757TdQSKJ+4eiRgNO/PT+t2R3Y18jUmmDgvoaU+2QzI2h
# F3MN9PNlOXBL85zWenvaDLw9MtAby/Vh/HUIAHa8gQ74wOFcz8QRcucbZEnYIpp1
# FUL1LTI4gdr0YKK6tFL7XOBhJCVPst/JKahzQ1HavWPWH1ub9y4bTxMd90oNcX6X
# t/Q/hOvB46NJofrOp79Wz7pZdmGJX36ntI5nePk2mOHLKNpbh6aKLzCCBTEwggQZ
# oAMCAQICEAqhJdbWMht+QeQF2jaXwhUwDQYJKoZIhvcNAQELBQAwZTELMAkGA1UE
# BhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
# ZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4X
# DTE2MDEwNzEyMDAwMFoXDTMxMDEwNzEyMDAwMFowcjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEx
# MC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIFRpbWVzdGFtcGluZyBD
# QTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAL3QMu5LzY9/3am6gpnF
# OVQoV7YjSsQOB0UzURB90Pl9TWh+57ag9I2ziOSXv2MhkJi/E7xX08PhfgjWahQA
# OPcuHjvuzKb2Mln+X2U/4Jvr40ZHBhpVfgsnfsCi9aDg3iI/Dv9+lfvzo7oiPhis
# EeTwmQNtO4V8CdPuXciaC1TjqAlxa+DPIhAPdc9xck4Krd9AOly3UeGheRTGTSQj
# MF287DxgaqwvB8z98OpH2YhQXv1mblZhJymJhFHmgudGUP2UKiyn5HU+upgPhH+f
# MRTWrdXyZMt7HgXQhBlyF/EXBu89zdZN7wZC/aJTKk+FHcQdPK/P2qwQ9d2srOlW
# /5MCAwEAAaOCAc4wggHKMB0GA1UdDgQWBBT0tuEgHf4prtLkYaWyoiWyyBc1bjAf
# BgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzASBgNVHRMBAf8ECDAGAQH/
# AgEAMA4GA1UdDwEB/wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDCDB5BggrBgEF
# BQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBD
# BggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0
# QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2Ny
# bDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDig
# NoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9v
# dENBLmNybDBQBgNVHSAESTBHMDgGCmCGSAGG/WwAAgQwKjAoBggrBgEFBQcCARYc
# aHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzALBglghkgBhv1sBwEwDQYJKoZI
# hvcNAQELBQADggEBAHGVEulRh1Zpze/d2nyqY3qzeM8GN0CE70uEv8rPAwL9xafD
# DiBCLK938ysfDCFaKrcFNB1qrpn4J6JmvwmqYN92pDqTD/iy0dh8GWLoXoIlHsS6
# HHssIeLWWywUNUMEaLLbdQLgcseY1jxk5R9IEBhfiThhTWJGJIdjjJFSLK8pieV4
# H9YLFKWA1xJHcLN11ZOFk362kmf7U2GJqPVrlsD0WGkNfMgBsbkodbeZY4UijGHK
# eZR+WfyMD+NvtQEmtmyl7odRIeRYYJu6DC0rbaLEfrvEJStHAgh8Sa4TtuF8QkIo
# xhhWz0E0tmZdtnR79VYzIi8iNrJLokqV2PWmjlIwggbmMIIEzqADAgECAhB3vQ4D
# obcI+FSrBnIQ2QRHMA0GCSqGSIb3DQEBCwUAMFMxCzAJBgNVBAYTAkJFMRkwFwYD
# VQQKExBHbG9iYWxTaWduIG52LXNhMSkwJwYDVQQDEyBHbG9iYWxTaWduIENvZGUg
# U2lnbmluZyBSb290IFI0NTAeFw0yMDA3MjgwMDAwMDBaFw0zMDA3MjgwMDAwMDBa
# MFkxCzAJBgNVBAYTAkJFMRkwFwYDVQQKExBHbG9iYWxTaWduIG52LXNhMS8wLQYD
# VQQDEyZHbG9iYWxTaWduIEdDQyBSNDUgQ29kZVNpZ25pbmcgQ0EgMjAyMDCCAiIw
# DQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBANZCTfnjT8Yj9GwdgaYw90g9z9Dl
# jeUgIpYHRDVdBs8PHXBg5iZU+lMjYAKoXwIC947Jbj2peAW9jvVPGSSZfM8RFpsf
# e2vSo3toZXer2LEsP9NyBjJcW6xQZywlTVYGNvzBYkx9fYYWlZpdVLpQ0LB/okQZ
# 6dZubD4Twp8R1F80W1FoMWMK+FvQ3rpZXzGviWg4QD4I6FNnTmO2IY7v3Y2FQVWe
# HLw33JWgxHGnHxulSW4KIFl+iaNYFZcAJWnf3sJqUGVOU/troZ8YHooOX1ReveBb
# z/IMBNLeCKEQJvey83ouwo6WwT/Opdr0WSiMN2WhMZYLjqR2dxVJhGaCJedDCndS
# sZlRQv+hst2c0twY2cGGqUAdQZdihryo/6LHYxcG/WZ6NpQBIIl4H5D0e6lSTmpP
# VAYqgK+ex1BC+mUK4wH0sW6sDqjjgRmoOMieAyiGpHSnR5V+cloqexVqHMRp5rC+
# QBmZy9J9VU4inBDgoVvDsy56i8Te8UsfjCh5MEV/bBO2PSz/LUqKKuwoDy3K1JyY
# ikptWjYsL9+6y+JBSgh3GIitNWGUEvOkcuvuNp6nUSeRPPeiGsz8h+WX4VGHaeki
# zIPAtw9FbAfhQ0/UjErOz2OxtaQQevkNDCiwazT+IWgnb+z4+iaEW3VCzYkmeVmd
# a6tjcWKQJQ0IIPH/AgMBAAGjggGuMIIBqjAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0l
# BAwwCgYIKwYBBQUHAwMwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQU2rON
# wCSQo2t30wygWd0hZ2R2C3gwHwYDVR0jBBgwFoAUHwC/RoAK/Hg5t6W0Q9lWULvO
# ljswgZMGCCsGAQUFBwEBBIGGMIGDMDkGCCsGAQUFBzABhi1odHRwOi8vb2NzcC5n
# bG9iYWxzaWduLmNvbS9jb2Rlc2lnbmluZ3Jvb3RyNDUwRgYIKwYBBQUHMAKGOmh0
# dHA6Ly9zZWN1cmUuZ2xvYmFsc2lnbi5jb20vY2FjZXJ0L2NvZGVzaWduaW5ncm9v
# dHI0NS5jcnQwQQYDVR0fBDowODA2oDSgMoYwaHR0cDovL2NybC5nbG9iYWxzaWdu
# LmNvbS9jb2Rlc2lnbmluZ3Jvb3RyNDUuY3JsMFYGA1UdIARPME0wQQYJKwYBBAGg
# MgEyMDQwMgYIKwYBBQUHAgEWJmh0dHBzOi8vd3d3Lmdsb2JhbHNpZ24uY29tL3Jl
# cG9zaXRvcnkvMAgGBmeBDAEEATANBgkqhkiG9w0BAQsFAAOCAgEACIhyJsav+qxf
# BsCqjJDa0LLAopf/bhMyFlT9PvQwEZ+PmPmbUt3yohbu2XiVppp8YbgEtfjry/Rh
# ETP2ZSW3EUKL2Glux/+VtIFDqX6uv4LWTcwRo4NxahBeGQWn52x/VvSoXMNOCa1Z
# a7j5fqUuuPzeDsKg+7AE1BMbxyepuaotMTvPRkyd60zsvC6c8YejfzhpX0FAZ/ZT
# fepB7449+6nUEThG3zzr9s0ivRPN8OHm5TOgvjzkeNUbzCDyMHOwIhz2hNabXAAC
# 4ShSS/8SS0Dq7rAaBgaehObn8NuERvtz2StCtslXNMcWwKbrIbmqDvf+28rrvBfL
# uGfr4z5P26mUhmRVyQkKwNkEcUoRS1pkw7x4eK1MRyZlB5nVzTZgoTNTs/Z7KtWJ
# QDxxpav4mVn945uSS90FvQsMeAYrz1PYvRKaWyeGhT+RvuB4gHNU36cdZytqtq5N
# iYAkCFJwUPMB/0SuL5rg4UkI4eFb1zjRngqKnZQnm8qjudviNmrjb7lYYuA2eDYB
# +sGniXomU6Ncu9Ky64rLYwgv/h7zViniNZvY/+mlvW1LWSyJLC9Su7UpkNpDR7xy
# 3bzZv4DB3LCrtEsdWDY3ZOub4YUXmimi/eYI0pL/oPh84emn0TCOXyZQK8ei4pd3
# iu/YTT4m65lAYPM8Zwy2CHIpNVOBNNwwggcAMIIE6KADAgECAgxVg/P+wUlFDmSx
# M2wwDQYJKoZIhvcNAQELBQAwWTELMAkGA1UEBhMCQkUxGTAXBgNVBAoTEEdsb2Jh
# bFNpZ24gbnYtc2ExLzAtBgNVBAMTJkdsb2JhbFNpZ24gR0NDIFI0NSBDb2RlU2ln
# bmluZyBDQSAyMDIwMB4XDTIxMTEyOTE1NDEwMVoXDTIyMTEzMDE1NDEwMVowbjEL
# MAkGA1UEBhMCREsxETAPBgNVBAcMCEJyw7huZGJ5MR4wHAYDVQQKExVNaWxlc3Rv
# bmUgU3lzdGVtcyBBL1MxDDAKBgNVBAsMA1ImRDEeMBwGA1UEAxMVTWlsZXN0b25l
# IFN5c3RlbXMgQS9TMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA0Eti
# hqPkmu1KV6LRSN9xz96UtyFGEZhOdBFPstJsSKMXQRCYcH3wcVxz6/pERdOTllrU
# uojlhmJKXUJ04ak7aHxY6/5WzUQZmm5b5uoZ27p9qSOVVgkYnpgwPAs41b5bC8qB
# V+NSELQniXQEOuacIRHN+oynZitag8Fy/7qtDzYaqmH74PBABr7vBMOUovEuxAxa
# r6v0dRIc2MYbWqOfF6jTJTX9fG0hW4nGQsC04EbENAdMHCZubiTswL20FgbjNq9O
# LpCp+Eu7sYRkmnHz2Kxd++9QlS4DtmI6Hbw6jy7a2WQP0vAsrqdVd1nh2DCbf058
# rHUTXNT/csXv82uSpwdLpSZVZigaKFsTmBmW94sVFK+TQYTe/4lpo+F/sMyPrw0i
# Fv/jJMlp5e3WKa8leiAIFiEfu7vnmI1FFslOlVfHYHc1fe2ERCe1DDW/hq3KFz8D
# 1q2CGMJpY6zY2iZ/mq2bJnMRASM9qOtRdTYLxPzm697bUdgK7p8SLtm1TzbzS1Js
# XpgRslxWXUWUAkSUeEeMXZHaF3wXZIn507FD/oupj0Goc4riHPxjhm9avY5tcMGY
# 8pflyYG1OOjNUlcHhW/cFX3/Tzr4UB2/sWkJ1Jopm22ZopCoDf905LrmgZOlVXc8
# cgApylcnpUKN9bl9XfqXxhYiaw0nz4hfImpjkTsCAwEAAaOCAbEwggGtMA4GA1Ud
# DwEB/wQEAwIHgDCBmwYIKwYBBQUHAQEEgY4wgYswSgYIKwYBBQUHMAKGPmh0dHA6
# Ly9zZWN1cmUuZ2xvYmFsc2lnbi5jb20vY2FjZXJ0L2dzZ2NjcjQ1Y29kZXNpZ25j
# YTIwMjAuY3J0MD0GCCsGAQUFBzABhjFodHRwOi8vb2NzcC5nbG9iYWxzaWduLmNv
# bS9nc2djY3I0NWNvZGVzaWduY2EyMDIwMFYGA1UdIARPME0wQQYJKwYBBAGgMgEy
# MDQwMgYIKwYBBQUHAgEWJmh0dHBzOi8vd3d3Lmdsb2JhbHNpZ24uY29tL3JlcG9z
# aXRvcnkvMAgGBmeBDAEEATAJBgNVHRMEAjAAMEUGA1UdHwQ+MDwwOqA4oDaGNGh0
# dHA6Ly9jcmwuZ2xvYmFsc2lnbi5jb20vZ3NnY2NyNDVjb2Rlc2lnbmNhMjAyMC5j
# cmwwEwYDVR0lBAwwCgYIKwYBBQUHAwMwHwYDVR0jBBgwFoAU2rONwCSQo2t30wyg
# Wd0hZ2R2C3gwHQYDVR0OBBYEFIdg9gMxxfcpe7gViT9QiXomUUGjMA0GCSqGSIb3
# DQEBCwUAA4ICAQAY8MM+XlsN3AEj6VASoSGTEVEIdMSZD0p2UPa8zkTWm7ZX0JLN
# O81eTCDRXzdS9jAz7U+ldJZwKYeogKshoBkPN0jBz+BCQNhOBYNfpGCFdYwH/cIF
# RPbcPB2sIQpe9Lr4+ZrU7MT5kx2ltznSaLHDf4wwvow0FsdrWfqJhNMsg4eYAuXU
# xnq0dmG2eJzB8XzoiFNhfv215Z45zYlG2vlczZnV2H8VfgvGA1Y6zLE7hLn5xaWg
# s+lwp2e87KIshP8qd+0DK3H+g0pFdcSQKjNnEgoASBAsdxSI0OLbZOLElOqAuaD2
# RAeEleEm0Ww1jZrncpDr8bq4LqT1q+fBMdPOhU28L+V20lHJrLwpjSWtatnKzEn2
# C5pBMgO2tcjPf7HH/Q30HIgCN6e9CxTxT/oO3eyxl7E9KtBL3edjE2/qFhpLVbdo
# P0KUb3go5jsnF9nmkualdsFICtzI7q4QARQi9+toixHk53eCUnxYEo9E01L45y6E
# BrvUWjvdj/A4RQZNjaiokzMYRluj6detWnN1cP0S18z3gwDlKA61bEAdpqvo3ifE
# oDtsCVQQty0duVp8nPzh8jT12V9+nQ9WQfch8x8TEm1qyNpHsIYDKGwsSLju95Zq
# ApYO7GFvzxvRjZMnUw85ePn/n8BY6HhALbPUmkn4PaXEee36XtllOAVvozGCBT4w
# ggU6AgEBMGkwWTELMAkGA1UEBhMCQkUxGTAXBgNVBAoTEEdsb2JhbFNpZ24gbnYt
# c2ExLzAtBgNVBAMTJkdsb2JhbFNpZ24gR0NDIFI0NSBDb2RlU2lnbmluZyBDQSAy
# MDIwAgxVg/P+wUlFDmSxM2wwCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAI
# oAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIB
# CzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFFrwPCyhHB82Km8w98vh
# JPFx1n4oMA0GCSqGSIb3DQEBAQUABIICAJdpYo/tyotZM227V2kYW9kBXXdSePgl
# ilLDe/tcMq7Ul8l32AnFvoIgclKQm8vNsPDELiFzRlkl3FlNd4BuyTqBfKY/820s
# g+4rbE/TOCd3SRoAuGQljpdoYObouNWA176wfG0dkBJR9/44ugaMD9vlvcIqt+Gl
# 92MNjkAP5gy6KupDixQom7bAHNfRt2QGdsbOds1FZkj5dUKHJ9peRgAWz5XFlx6l
# oU4kP7/cBbRozMZd7Ab4VGhHXUNawyTKt6A7sD7Imb1q93rHXd+USgxNlPpG2p2N
# p0So9p0E9zj7YALu2WbQLPK3byTcaTG1xHVYeF1jWkfxwDX51utJGezTz6qEhvM3
# YLkSJIIx88IV9FXMl4Bu1qYdxhWo5nwQLmFZSY8lXyhkmgRCYkCZAKdRyVpDS+dy
# AF5IuAy7jBrRrPCKRrRbythBKTLFDIobJU0vtzevPlsm1iIcjtEvR7OqglWfoNmf
# ViLwaiX5lMVoSZkpqW2ZmfDh2bNrrbnAqOjnihHk0G26GED5ZYSsCGnK/Ebc2zrt
# BBidPBSINUM6FqCvLcOtzODhWOh8E4VUug/Gz+OVluzLifcvb6J2cePF5G7ydESX
# 5wsmMjFubWIE8wX97GsYa17h66UdftIlNymt/HOw6t+vp4WyRJxXExVWA0/VH0do
# szPmw/pExCaVoYICMDCCAiwGCSqGSIb3DQEJBjGCAh0wggIZAgEBMIGGMHIxCzAJ
# BgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5k
# aWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBU
# aW1lc3RhbXBpbmcgQ0ECEA1CSuC+Ooj/YEAhzhQA8N0wDQYJYIZIAWUDBAIBBQCg
# aTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yMjAy
# MDkwNjMxNDJaMC8GCSqGSIb3DQEJBDEiBCDrIfkISbnIPOONFqtWkxy0e6qwioyq
# iSbJpIRHKk7WiTANBgkqhkiG9w0BAQEFAASCAQCa0kMZd1T5k0C49qHf9XvoyVWk
# MXXcW/q89NE2xrSNVd0+LN4CD1l4PeSm9ZRpBGCI6ygEGJTKSBdcDI14kbLIN/2p
# G/hKbxCmY+U5B2i17sA9fjUhLojK0GamxLWbr81RAB0D3lzWUGTz5Mtsqp79g09X
# gMIJae5Jyk74S2JQRQcVR/1Wjj08w7y5tJ8rI2novV5Iwhy4l/CHqYTdyWsm6Awy
# MLg4N6QQF98RZRbTnGe440Pxam5tNhpq8cOl6fhwgH+HphlxSzmpeTLpvEUm7Qvx
# ez70uCanU8CURMU/gZoxE7YCkA1NPvHU+rlYmWNnBRJEPU9OFun1irsiKCHO
# SIG # End signature block

function Get-StatusFromValue {
    Param($SV)
    switch ($SV) {
        0 { "Disconnected" }
        1 { "Connecting" }
        2 { "Connected" }
        3 { "Disconnecting" }
        4 { "Hardware not present" }
        5 { "Hardware disabled" }
        6 { "Hardware malfunction" }
        7 { "Media disconnected" }
        8 { "Authenticating" }
        9 { "Authentication succeeded" }
        10 { "Authentication failed" }
        11 { "Invalid Address" }
        12 { "Credentials Required" }
        Default { "Not connected" }
    }
}  
function Get-VmsStorageInfo {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.RecordingServer[]]
        $RecordingServer
    )

    process {
        if ($null -eq $RecordingServer -or $RecordingServer.Count -eq 0) {
            $RecordingServer = Get-RecordingServer
        }
        foreach ($recorder in $RecordingServer) {
            foreach ($storage in $recorder | Get-VmsStorage) {
                $info = $storage.ReadStorageInformation()
                $usedSpace = $info.GetProperty('UsedSpace') -as [double]
                $lockedUsedSpace = 0
                if ($info.GetPropertyKeys() -contains 'LockedUsedSpace') {
                    $lockedUsedSpace = $info.GetProperty('LockedUsedSpace') -as [double]
                }

                [pscustomobject]@{
                    RecordingServer = $recorder.Name
                    Name            = $storage.Name
                    Path            = $storage.DiskPath
                    RetainMinutes   = $storage.RetainMinutes
                    MaxSize         = $storage.MaxSize
                    UsedSpace       = $usedSpace * 1MB
                    LockedUsedSpace = $lockedUsedSpace * 1MB
                    Signing         = $storage.Signing
                }
                Remove-Variable info, usedSpace, lockedUsedSpace

                foreach ($archive in $storage | Get-VmsArchiveStorage) {
                    $info = $archive.ReadArchiveStorageInformation()
                    $usedSpace = $info.GetProperty('UsedSpace') -as [double]
                    $lockedUsedSpace = 0
                    if ($info.GetPropertyKeys() -contains 'LockedUsedSpace') {
                        $lockedUsedSpace = $info.GetProperty('LockedUsedSpace') -as [double]
                    }

                    [pscustomobject]@{
                        RecordingServer = $recorder.Name
                        Name            = $archive.Name
                        Path            = $archive.DiskPath
                        RetainMinutes   = $archive.RetainMinutes
                        MaxSize         = $archive.MaxSize
                        UsedSpace       = $usedSpace * 1MB
                        LockedUsedSpace = $lockedUsedSpace * 1MB
                        Signing         = $archive.Signing
                    }
                }
            }
        }
    }
}
function Convert {
    Param($RetainMinutes)
    
    $ts = New-TimeSpan -Minutes $RetainMinutes
    if ($ts.TotalDays -lt 1) {
        $ts.TotalHours.ToString() + " Hours" 
    }
    else { $ts.TotalDays.ToString() + " Days" }
}  

function Get-RecorderReport {
    [CmdletBinding()]
    param (
        # Specifies one or more Recording Servers from which to generate a camera report. By default all Recording Servers will be used.
        [Parameter(ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.RecordingServer[]]
        $RecordingServer
    )

    begin {
        $runspacepool = [runspacefactory]::CreateRunspacePool(4, 16)
        $runspacepool.Open()
        $threads = New-Object System.Collections.Generic.List[pscustomobject]

        $process = {
            param(
                [VideoOS.Platform.ConfigurationItems.RecordingServer]$Recorder
            )
            try {
                $hardware = $recorder | Get-Hardware
                $cameras = $hardware | Get-Camera
                $enabledHardware = $hardware | Where-Object Enabled
                $enabledCameras = $cameras | Where-Object { $_.Enabled -and $_.ParentItemPath -in $enabledHardware.Path }

                $obj = [PSCustomObject]@{
                    RecordingServer           = $recorder.Name
                    TotalHardware             = $hardware.Count
                    EnabledHardware           = $enabledHardware.Count
                    TotalCameras              = $cameras.Count
                    EnabledCameras            = $enabledCameras.Count
                    CamerasStarted            = 0
                    UsedSpaceInBytes          = 0
                    RecordedBPS               = 0
                    TotalBPS                  = 0
                    OverflowCount             = 0
                    CamerasWithErrors         = 0
                    CamerasNotConnected       = 0
                    DatabaseRepairsInProgress = 0
                    DatabaseWriteErrors       = 0
                    CamerasNotLicensed        = 0
                }

                try {
                    $svc = $recorder | Get-RecorderStatusService2
                    $stats = $svc.GetVideoDeviceStatistics((Get-Token), [guid[]]$enabledCameras.Id)
                    $status = $svc.GetCurrentDeviceStatus((Get-Token), [guid[]]$enabledCameras.Id)
                    $liveStreams = $stats.VideoStreamStatisticsArray | Where-Object RecordingStream -eq $false
                    $recordedStreams = $stats.VideoStreamStatisticsArray | Where-Object RecordingStream
                    $recordedBPS = $recordedStreams | Measure-Object -Property BPS -Sum | Select-Object -ExpandProperty Sum

                    $obj.CamerasStarted = ($status.CameraDeviceStatusArray | Where-Object Started).Count
                    $obj.UsedSpaceInBytes = $stats | Measure-Object -Property UsedSpaceInBytes -Sum | Select-Object -ExpandProperty Sum
                    $obj.RecordedBPS = $recordedBPS
                    $obj.TotalBPS = $recordedBPS + ( $liveStreams | Measure-Object -Property BPS -Sum | Select-Object -ExpandProperty Sum )
                    $obj.OverflowCount = ($status.CameraDeviceStatusArray | Where-Object ErrorOverflow).Count
                    $obj.CamerasWithErrors = ($status.CameraDeviceStatusArray | Where-Object Error).Count
                    $obj.CamerasNotConnected = ($status.CameraDeviceStatusArray | Where-Object ErrorNoConnection).Count
                    $obj.DatabaseRepairsInProgress = ($status.CameraDeviceStatusArray | Where-Object DbRepairInProgress).Count
                    $obj.DatabaseWriteErrors = ($status.CameraDeviceStatusArray | Where-Object ErrorWritingGop).Count
                    $obj.CamerasNotLicensed = ($status.CameraDeviceStatusArray | Where-Object CamerasNotLicensed).Count
                }
                catch {
                    Write-Error -Exception $_.Exception -Message "Error collecting statistics from $($recorder.Name) ($($recorder.Hostname))"
                }

                Write-Output $obj
            }
            catch {
                Write-Error -Exception $_.Exception -Message "Unexpected error: $($_.Message). $($recorder.Name) ($($recorder.Hostname)) will not be included in the report."
            }
            finally {
                $svc.Dispose()
            }
        }
    }

    process {
        $progressParams = @{
            Activity         = $MyInvocation.MyCommand.Name
            CurrentOperation = ''
            PercentComplete  = 0
            Completed        = $false
        }

        if ($null -eq $RecordingServer) {
            $RecordingServer = Get-RecordingServer
        }

        try {
            foreach ($recorder in $RecordingServer) {
                $ps = [powershell]::Create()
                $ps.RunspacePool = $runspacepool
                $asyncResult = $ps.AddScript($process).AddParameters(@{
                        Recorder = $recorder
                    }).BeginInvoke()
                $threads.Add([pscustomobject]@{
                        PowerShell = $ps
                        Result     = $asyncResult
                    })
            }

            if ($threads.Count -eq 0) {
                return
            }

            $progressParams.CurrentOperation = 'Processing requests for recorder information'
            $completedThreads = New-Object System.Collections.Generic.List[pscustomobject]
            $totalJobs = $threads.Count
            while ($threads.Count -gt 0) {
                $progressParams.PercentComplete = ($totalJobs - $threads.Count) / $totalJobs * 100
                $progressParams.Status = "Processed $($totalJobs - $threads.Count) out of $totalJobs requests"
                Write-Progress @progressParams
                foreach ($thread in $threads) {
                    if ($thread.Result.IsCompleted) {
                        $thread.PowerShell.EndInvoke($thread.Result)
                        $thread.PowerShell.Dispose()
                        $completedThreads.Add($thread)
                    }
                }
                $completedThreads | Foreach-Object { [void]$threads.Remove($_) }
                $completedThreads.Clear()
                if ($threads.Count -eq 0) {
                    break;
                }
                Start-Sleep -Seconds 1
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

Write-Progress -Activity "Connecting to Servers" -Status "1% Complete:" -PercentComplete 1
$User = ".\Administrator" 
$PWord = ConvertTo-SecureString -String "Milestone1$" -AsPlainText -Force 
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, $PWord
$Sessions = New-CimSession -ComputerName '10.1.0.24' -Credential $Credential 

$inc = 3; $i = 0


######################################################################################################################################################################################
# Aplication - Milestone Installed Services - Win32_Service
######################################################################################################################################################################################

$i += $inc; Write-Progress -Activity "Collecting Data - Aplication - Milestone Installed Services" -Status "$i% Complete:" -PercentComplete $i

$Win32_Service = Get-CimInstance -Query "Select * from Win32_Service Where Name like 'Milestone%' or Name like 'VideoOS%'" -CimSession $Sessions `
| Select-Object PSComputerName, Name, Description, State, StartMode, StartName, PathName `
| Sort-Object -Property PSComputerName, Name 

######################################################################################################################################################################################
# Aplication - Milestone Installed products - Win32_Product
######################################################################################################################################################################################
$i += $inc; Write-Progress -Activity "Collecting Data - Aplication - Milestone Installed products" -Status "$i% Complete:" -PercentComplete $i


$Win32_Product = Get-CIMInstance -ClassName Win32_Product -Filter "Vendor like '%Milestone Systems%'" -CimSession $Sessions `
| Select-Object PSComputerName, Name, Vendor, Version, InstallDate `
| Sort-Object -Property PSComputerName, Name 

$Win32_InstalledWin32Program = Get-CIMInstance -ClassName  Win32_InstalledWin32Program -Filter "Vendor like '%Milestone%' and Name like '%Hotfix%'" `
| Select-Object PSComputerName, Name, Vendor, Version, InstallDate `
| Sort-Object -Property PSComputerName, Name 


$ManagementServer = $Win32_Product | Group-Object -Property Name `
| Where-Object Name -like '*Management Server*' `
| Select-Object -ExpandProperty Group `
| Sort-Object -Property PSComputerName, Name 

$RecordingServers = $Win32_Product | Group-Object -Property Name `
| Where-Object Name -like '*Recording*' `
| Select-Object -ExpandProperty Group `
| Sort-Object -Property PSComputerName, Name 

$DriverPacks = $Win32_Product | Group-Object -Property Name `
| Where-Object Name -like '*Device Pack*' `
| Select-Object -ExpandProperty Group `
| Sort-Object -Property PSComputerName, Name 

$EventServerServers = $Win32_Product | Group-Object -Property Name `
| Where-Object Name -like '*Event Server*' `
| Select-Object -ExpandProperty Group `
| Sort-Object -Property PSComputerName, Name 

$MobileServers = $Win32_Product | Group-Object -Property Name `
| Where-Object Name -like '*XProtect Mobile*' `
| Select-Object -ExpandProperty Group `

$DataCollectors = $Win32_Product | Group-Object -Property Name `
| Where-Object Name -like '*Data Collector*' `
| Select-Object -ExpandProperty Group `
| Sort-Object -Property PSComputerName, Name 

$otherServers = $Win32_Product | Group-Object -Property Name | Where-Object { 
    ($_.Name -notlike "*Management Server*" `
        -and $_.Name -notlike "*Recording*" `
        -and $_.Name -notlike "*Device Pack*" `
        -and $_.Name -notlike "*Data Collector*" `
        -and $_.Name -notlike "*Event Server*" `
        -and $_.Name -notlike "*XProtect Mobile*" `
    ) } | Select-Object -ExpandProperty Group `
| Sort-Object -Property PSComputerName, Name 




######################################################################################################################################################################################
# Operating System - Updates - Win32_QuickFixEngineering
######################################################################################################################################################################################
$i += $inc; Write-Progress -Activity "Collecting Data - Operating System - Updates" -Status "$i% Complete:" -PercentComplete $i

$Win32_QuickFixEngineering = Get-CIMInstance -ClassName Win32_QuickFixEngineering -CimSession $Sessions `
| Select-Object PSComputerName, InstalledOn, HotFixID `
| Sort-Object -Property PSComputerName, HotFixID 

######################################################################################################################################################################################
# Operating System - Start Commands  - Win32_StartupCommand
######################################################################################################################################################################################
$i += $inc; Write-Progress -Activity "Collecting Data - Operating System - Start Commands" -Status "$i% Complete:" -PercentComplete $i

$Win32_StartupCommand = Get-CIMInstance -ClassName Win32_StartupCommand -CimSession $Sessions `
| Select-Object PSComputerName, Name, Command `
| Sort-Object -Property PSComputerName, Name 


######################################################################################################################################################################################
##################################################################### MILESTONE PS TOOLS #############################################################################################
######################################################################################################################################################################################

Connect-ManagementServer -Server SGIU-VM1 -Credential $Credential -Force -AcceptEula

######################################################################################################################################################################################
# VMS - Licensed Products  - Get-LicensedProducts
######################################################################################################################################################################################
$i += $inc; Write-Progress -Activity "Collecting Data - Operating System - Start Commands" -Status "$i% Complete:" -PercentComplete $i

$licencedProducts = Get-LicensedProducts `
| Select-Object ProductDisplayName, Slc, ExpirationDate, CarePlus, CarePremium, Name, DisplayName

######################################################################################################################################################################################
# VMS - Licensed Products  - Get-LicensedProducts
######################################################################################################################################################################################
$i += $inc; Write-Progress -Activity "Collecting Data - Operating System - Start Commands" -Status "$i% Complete:" -PercentComplete $i

$licenceDetails = Get-LicenseDetails `
| Select-Object LicenseType, Activated, ChangesWithoutActivation, InGrace, GraceExpired , NotLicensed, Note

######################################################################################################################################################################################
# VMS - License Overview  - Get-LicenseOverview
######################################################################################################################################################################################
$i += $inc; Write-Progress -Activity "Collecting Data - Operating System - Start Commands" -Status "$i% Complete:" -PercentComplete $i

$licenceOverview = Get-LicenseOverview `
| Select-Object LicenseType, Activated 

######################################################################################################################################################################################
# VMS - Get Storage Usage 
######################################################################################################################################################################################
$i += $inc; Write-Progress -Activity "VMS - Get Storage Usage" -Status "$i% Complete:" -PercentComplete $i



$storageInforation = Get-RecordingServer `
| Get-VmsStorageInfo `
| Select-Object @{Name = "PSComputerName"; Expression = { $_.RecordingServer } } , Name  `
    , @{Name = "Retain"; Expression = { Convert($_.RetainMinutes) } }`
    , @{Name = "MaxSize_GB"; Expression = { [math]::round($_.MaxSize / 1KB, 2) } } `
    , @{Name = "UsedSpace_GB"; Expression = { [math]::round($_.UsedSpace / 1GB, 2) } } `
    , @{Name = "UsedSpaca_%"; Expression = { (($_.UsedSpace / 1GB) / ($_.MaxSize / 1KB)).tostring("P") } } `
    , LockedUsedSpace, Path

######################################################################################################################################################################################
# VMS - Camera Report 
######################################################################################################################################################################################
$i += $inc; Write-Progress -Activity "VMS - Get VMS Camera Report " -Status "$i% Complete:" -PercentComplete $i

$cameraReport = Get-VmsCameraReport -IncludeSnapshots `
| Select-Object RecorderName, HardwareName, Name , Model, Address `
    , MAC, Firmware, Driver, RecordingEnabled, RecordKeyframesOnly , MotionEnabled `
    , MotionKeyframesOnly, State , Snapshot `
| Sort-Object -Property  RecorderName, HardwareName, Name

$Directory = "c:\Temp\"
foreach ($C in $cameraReport) {
    $FileName = '{0}-{1}.jpg' -f (Join-Path (Resolve-Path $Directory) ($C.Name)), (Get-Date).ToString('yyyyMMdd_HHmmss')
    $EncoderParamSet = New-Object System.Drawing.Imaging.EncoderParameters(1) 
    $EncoderParamSet.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, 70) 
    $JPGCodec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/png' }
    $C.Snapshot.Save($FileName , $JPGCodec, $EncoderParamSet)
    $C.Snapshot = $FileName
}


######################################################################################################################################################################################
# VMS - Recording Server Report 
######################################################################################################################################################################################
$i += $inc; Write-Progress -Activity "VMS - Get Recording Server Report" -Status "$i% Complete:" -PercentComplete $i

$recordinServerReport = Get-RecorderReport | Select-Object RecordingServer, TotalHardware, EnabledHardware `
    , TotalCameras, EnabledCameras, RecordedBPS, TotalBPS, OverflowCount, CamerasWithErrors `
    , DatabaseRepairsInProgress, DatabaseWriteErrors, CamerasNotLicensed


######################################################################################################################################################################################
# VMS - Logs 
######################################################################################################################################################################################
$i += $inc; Write-Progress -Activity "VMS - Get Logs" -Status "$i% Complete:" -PercentComplete $i

$logs = Get-Log -Tail -Minutes 10080 | Select-Object "Log level", "Local time", "Message text", "Category", "Source type", "Source name", "Event type"

$logs_system_group = Get-Log -LogType System -Tail -Minutes 10080 `
| Group-Object "Source name", "Message text" `
| Select-Object Count, Name `
| Sort-Object -Property Count -Descending

$logs_Audit_group = Get-Log -LogType Audit -Tail -Minutes 10080 `
| Group-Object "User", "Message text" `
| Select-Object Count, Name `
| Sort-Object -Property Count -Descending

$logs_Rules_group = Get-Log -LogType Rules  -Tail -Minutes 10080 `
| Group-Object "Source name", "Event type" `
| Select-Object Count, Name `
| Sort-Object -Property Count -Descending

######################################################################################################################################################################################
############################################################## PERFORMANCE COUNTERS ##################################################################################################
######################################################################################################################################################################################



######################################################################################################################################################################################
# Performance - CPU
######################################################################################################################################################################################
$i += $inc; Write-Progress -Activity "Performance - CPU" -Status "$i% Complete:" -PercentComplete $i
$samples = 10

$cpu_samples = New-Object System.Collections.Generic.List[System.Object]
foreach ($s in $Sessions) {
    $cpu_samples += [pscustomobject]@{
        DisplayName    = $s.ComputerName
        PSComputerName = $s.ComputerName
        Samples        = New-Object System.Collections.Generic.List[System.Object]
        Max            = 100
    }
}

for ($i = 0; $i -lt $samples; $i++) {
    $cpu_ = Get-CimInstance Win32_Processor -CimSession $Sessions | Select-Object LoadPercentage, PSComputerName
    $timestamp = (Get-Date).ToString('u')
    foreach ($cpu_sample_ in $cpu_ ) {
            ($cpu_samples | Where-Object PSComputerName -eq $cpu_sample_.PSComputerName).Samples.Add(
            [pscustomobject]@{
                TimeStamp = $timestamp
                Value     = $cpu_sample_.LoadPercentage
            }
        )
    }
    Start-Sleep -Seconds 1
}


######################################################################################################################################################################################
# Performance - Memory
######################################################################################################################################################################################
$i += $inc; Write-Progress -Activity "Performance - Memory " -Status "$i% Complete:" -PercentComplete $i
$samples = 10

$TotalMemory = Get-CIMInstance Win32_OperatingSystem -CimSession $Sessions | Select-Object TotalVisibleMemorySize, PSComputerName 

$mem_samples = New-Object System.Collections.Generic.List[System.Object]
foreach ($s in $Sessions) {
    $mem_samples += [pscustomobject]@{
        DisplayName    = $s.ComputerName
        PSComputerName = $s.ComputerName
        Samples        = New-Object System.Collections.Generic.List[System.Object]
        Max            = ($TotalMemory | Where-Object PSComputerName -eq $s.ComputerName).TotalVisibleMemorySize
    }
}

for ($i = 0; $i -lt $samples; $i++) {
    $mem_ = Get-CIMInstance Win32_OperatingSystem -CimSession $Sessions | Select-Object @{Name = "UsedMemory"; Expression = { $_.TotalVisibleMemorySize - $_.FreePhysicalMemory } } , PSComputerName 
    
    
    $timestamp = (Get-Date).ToString('u')
    foreach ($mem_sample_ in $mem_ ) {
            ($mem_samples | Where-Object PSComputerName -eq $mem_sample_.PSComputerName).Samples.Add(
            [pscustomobject]@{
                TimeStamp = $timestamp
                Value     = $mem_sample_.UsedMemory
            }
        )
    }
    Start-Sleep -Seconds 1
}

######################################################################################################################################################################################
# Performance - Network
######################################################################################################################################################################################
$i += $inc; Write-Progress -Activity "Performance - Network " -Status "$i% Complete:" -PercentComplete $i
$samples = 10

$network_interfaces = Get-CIMInstance Win32_PerfFormattedData_Tcpip_NetworkInterface -CimSession $Sessions | Select-Object CurrentBandwidth, Name , PSComputerName
    
$network_samples = New-Object System.Collections.Generic.List[System.Object]

foreach ($PSCcomputer_interface in $network_interfaces | Select-Object PSComputerName -Unique ) {
    foreach ($networ_interface in $network_interfaces | Select-Object Name, CurrentBandwidth -Unique) {

        $network_samples += [pscustomobject]@{
            PSComputerName = $PSCcomputer_interface.PSComputerName
            Name           = $networ_interface.Name
            DisplayName    = "$($PSCcomputer_interface.PSComputerName) - $($networ_interface.Name)"
            #Max            = [math]::Round($networ_interface.CurrentBandwidth * 0.001 , 2)
            Samples        = New-Object System.Collections.Generic.List[System.Object]
        }
    }
}

for ($i = 0; $i -lt $samples; $i++) {
    
    Write-Progress -Activity "Performance - Network (${i+1})" -Status "$i% Complete:" -PercentComplete $i
 
    $network_raw_samples = Get-CIMInstance -class Win32_PerfFormattedData_Tcpip_NetworkInterface -CimSession $Sessions | Select-Object BytesReceivedPersec, BytesSentPersec , BytesTotalPersec , Name , PSComputerName
    $timestamp = (Get-Date).ToString('u')

    foreach ($network_sample in $network_raw_samples ) {
    
        $BytesReceivedPersec = [math]::Round($network_sample.BytesReceivedPersec * 8000 * 0.000001 , 2)
        $BytesSentPersec = [math]::Round($network_sample.BytesSentPersec * 8000 * 0.000001 , 2)
        $BytesTotalPersec = [math]::Round($network_sample.BytesTotalPersec * 8000 * 0.000001 , 2)

        $actual_sample = ($network_samples | Where-Object { ($_.PSComputerName -eq $network_sample.PSComputerName) -and ($_.Name -eq $network_sample.Name) })
        $actual_sample.Samples.Add(
            [pscustomobject]@{
                TimeStamp = $timestamp
                Value     = $BytesTotalPersec
                Value2    = $BytesReceivedPersec
                Value3    = $BytesSentPersec
            })
    }
    Start-Sleep -Seconds 5
}

######################################################################################################################################################################################
# Build output and save 
######################################################################################################################################################################################


$products = [pscustomobject]@{
    Win32_Service      = $Win32_Service
    ManagementServer   = $ManagementServer
    RecordingServers   = $RecordingServers 
    EventServerServers = $EventServerServers
    MobileServers      = $MobileServers
    DriverPacks        = $DriverPacks
    DataCollectors     = $DataCollectors
    OtherServers       = $otherServers
    Hotfix             = $Win32_InstalledWin32Program
}

$Report = [pscustomobject]@{
  # cameraReport = $cameraReport
    
}

$RecordingServerReport = [pscustomobject]@{
    RecordingServerReport = $recordinServerReport
}

$licenses = [pscustomobject]@{
    LicencedProducts = $licencedProducts 
    LicenceDetails   = $licenceDetails 
    licenceOverview  = $licenceOverview
}

$performance = [pscustomobject]@{
    cpu_samples     = $cpu_samples 
    mem_samples     = $mem_samples 
    network_samples = $network_samples
}

$logs = [pscustomobject]@{
    Logs_system_group = $logs_system_group
    Logs_Audit_group  = $logs_Audit_group
    Logs_Rules_group  = $logs_Rules_group
    Logs              = $logs

}

$serversInformation = [pscustomobject]@{
    Win32_ComputerSystem                           = $Win32_ComputerSystem
    Win32_OperatingSystem                          = $Win32_OperatingSystem
    Win32_Processor                                = $Win32_Processor
    Win32_PhysicalMemory                           = $Win32_PhysicalMemory
    Win32_Volume                                   = $Win32_Volume
    StorageInforation                              = $storageInforation
    Win32_VideoController                          = $Win32_VideoController
    Win32_NetworkAdapter                           = $Win32_NetworkAdapter
    Win32_NetworkAdapterConfiguration              = $Win32_NetworkAdapterConfiguration
    Win32_PerfFormattedData_Tcpip_NetworkInterface = $Win32_PerfFormattedData_Tcpip_NetworkInterface
    Win32_QuickFixEngineering                      = $Win32_QuickFixEngineering
    Win32_StartupCommand                           = $Win32_StartupCommand
}

$obj = [pscustomobject]@{
    Products              = $products 
    Licenses              = $licenses 
    ServersInformation    = $serversInformation
    Report                = $Report
    RecordingServerReport = $RecordingServerReport
    Performance           = $performance
    Logs                  = $logs
}

$obj | ConvertTo-Json -Depth 100 | Set-Content C:\Users\sgiu\source\repos\Electron-MHC\json\obj.json

#region Import Milestone PSTools
Import-Module .\pstools\MipSdkRedist\21.2.0\MipSdkRedist.psm1
Import-Module .\pstools\MilestonePSTools\21.2.6\MilestonePSTools.psm1
#endregion

#region Initialize 

## Get Server Name 
$ServerName = $env:COMPUTERNAME
## Get Current Time 
$TimeStamp = Get-Date

## Perfomance Counter Sample Amount 
$samples = 1

$CompletePercentage = 0;
$inc = 100 / $FunctionsToCall;

$RecordingServerLogPath = "{$env:ProgramData}\Milestone\Mileston Recording Server\Logs"


#endregion

#region Determine Server type
$MilestoneServices = Get-CimInstance -Query "Select * from Win32_Service Where Name like 'Milestone%' or Name like 'VideoOS%'" `
| Select-Object PSComputerName, Name, Description, State, StartMode, StartName, PathName `
| Sort-Object -Property PSComputerName, Name 

$IsManagementServer = ($MilestoneServices -match 'Milestone XProtect Management Server').Count > 0

$IsRecordingServer = ($MilestoneServices -match 'Milestone XProtect Recording Server').Count > 0

$HasMilestoneService = ($MilestoneServices -match 'Milestone').Count > 0

# DEBUG
$IsManagementServer = $false
$IsRecordingServer = $true
$HasMilestoneService = $true

$FunctionsToCall = 1
if ($IsManagementServer) { $FunctionsToCall += 3 }
if ($IsRecordingServer) { $FunctionsToCall += 2 }
if ($HasMilestoneService) { $FunctionsToCall += 5 }

#endregion

#region HasMilestoneService
if ($HasMilestoneService) {

    #region XProtect Cumulative Updates
    $CompletePercentage += $inc; Write-Progress -Activity "XProtect Cumulative Updates" -Status "$CompletePercentage% Complete:" -PercentComplete $CompletePercentage
    
    $XProtectCumulativeUpdates = [pscustomobject]@{
        CumulativeUpdates = $CumulativeUpdates
    }
    #endregion

    #region System RAM utilization
    $SystemRAMutilization = [pscustomobject]@{
        Samples = New-Object System.Collections.Generic.List[System.Object]
        Max     = ($TotalMemory | Where-Object PSComputerName -eq $s.ComputerName).TotalVisibleMemorySize
    }

    $TotalMemory = Get-CIMInstance Win32_OperatingSystem | Select-Object TotalVisibleMemorySize

    for ($i = 0; $i -lt $samples; $i++) {
        $CompletePercentage += $inc / $samples ; Write-Progress -Activity " System RAM utilization" -Status "$CompletePercentage% Complete:" -PercentComplete $CompletePercentage

        $mem_sample = Get-CIMInstance Win32_OperatingSystem | Select-Object @{Name = "UsedMemory"; Expression = { $_.TotalVisibleMemorySize - $_.FreePhysicalMemory } }
        
        $timestamp = (Get-Date).ToString('u')

        $SystemRAMutilization.Samples.Add(
            [pscustomobject]@{
                TimeStamp = $timestamp
                Value     = $mem_sample.UsedMemory
            }
        )
        
        Start-Sleep -Seconds 1
    }
    #endregion

    #region System CPU utilization

    $SystemCPUutilization = [pscustomobject]@{
        Samples = New-Object System.Collections.Generic.List[System.Object]
        Max     = 100
    }
    
    for ($i = 0; $i -lt $samples; $i++) {
        $CompletePercentage += $inc / $samples; Write-Progress -Activity " System CPU utilization" -Status "$CompletePercentage% Complete:" -PercentComplete $CompletePercentage

        $cpu_sample = Get-CimInstance Win32_Processor | Select-Object LoadPercentage
        
        $timestamp = (Get-Date).ToString('u')
        
        $SystemCPUutilization.Samples.Add(
            [pscustomobject]@{
                TimeStamp = $timestamp
                Value     = $cpu_sample.LoadPercentage
            }
        )
        #Start-Sleep -Seconds 1 # No Need to sleep, CIM call is quite slow 
    }
    #endregion

    #region Hardware acceleration capability
    $CompletePercentage += $inc; Write-Progress -Activity "Hardware acceleration capability" -Status "$CompletePercentage% Complete:" -PercentComplete $CompletePercentage
    
    $HardwareaAcelerationCapability = Get-CIMInstance -ClassName Win32_VideoController | Select-Object VideoProcessor, Name, AdapterRAM, DriverDate, DriverVersion
    #endregion

    #region Antivirus presence
    $CompletePercentage += $inc; Write-Progress -Activity "Antivirus presence" -Status "$CompletePercentage% Complete:" -PercentComplete $CompletePercentage

    $AntivirusPresence = [pscustomobject]@{
        Antivirus = $Antivirus
    }
    #endregion
}
#endregion

#region IsManagementServer
if ($IsManagementServer) {

    #region Connect ManagementServer
    $CompletePercentage += $inc; Write-Progress -Activity "Connect ManagementServer" -Status "$CompletePercentage% Complete:" -PercentComplete $CompletePercentage
    Connect-ManagementServer -AcceptEula
    #endregion

    #region Milestone XProtect Version
    $CompletePercentage += $inc; Write-Progress -Activity "Milestone XProtect Version" -Status "$CompletePercentage% Complete:" -PercentComplete $CompletePercentage

    $MilestoneXProtectVersion = [pscustomobject]@{
        Version = $Version
        Product = $Product
    }
    #endregion

    #region Milestone Care Status
    $CompletePercentage += $inc; Write-Progress -Activity "Milestone Care Status" -Status "$CompletePercentage% Complete:" -PercentComplete $CompletePercentage

    $MilestoneCareStatus = [pscustomobject]@{
        CarePlus       = $CarePlus
        ExpirationDate = $ExpirationDate
    }
    #endregion

    #region Failover Configuration
    $FailoverConfiguration = [pscustomobject]@{
    #endregion

    }

}
#endregion

#region IsRecordingServer
if ($IsRecordingServer) {

    #region Media Deletion Due to Low Disk Space
    $CompletePercentage += $inc; Write-Progress -Activity "Media Deletion Due to Low Disk Space" -Status "$CompletePercentage% Complete:" -PercentComplete $CompletePercentage

    #JOIN DATABASE LOGS AND FIND THE ERROR

    $MediaDeletionDuetoLowDiskSpace = [pscustomobject]@{
        DeletionDuetoLowDiskSpaceErrorList = $DeletionDuetoLowDiskSpaceErrorList
    }
    #endregion
    
    #region Media Deletion Due to Overflow
    $CompletePercentage += $inc; Write-Progress -Activity "Media Deletion Due to Overflow" -Status "$CompletePercentage% Complete:" -PercentComplete $CompletePercentage

    # READ DEVICEHANDLING LOG AND FINF OVERFLOW ERROR

    $MediaDeletionDuetoOverflow = [pscustomobject]@{
        MediaDeletionDuetoOverflowList = $MediaDeletionDuetoOverflowList
    }
    #endregion
}
#endregion 

#region build Output
$Output = [PSCustomObject]@{
    ServerName                     = $ServerName 
    TimeStamp                      = $TimeStamp
    MilestoneXProtectVersion       = $MilestoneXProtectVersion
    MilestoneCareStatus            = $MilestoneCareStatus
    XProtectCumulativeUpdates      = $XProtectCumulativeUpdates
    MediaDeletionDuetoLowDiskSpace = $MediaDeletionDuetoLowDiskSpace
    MediaDeletionDuetoOverflow     = $MediaDeletionDuetoOverflow
    SystemRAMutilization           = $SystemRAMutilization
    SystemCPUutilization           = $SystemCPUutilization
    FailoverConfiguration          = $FailoverConfiguration
    HardwareaAcelerationCapability = $HardwareaAcelerationCapability
    AntivirusPresence              = $AntivirusPresence
}

$CompletePercentage += $inc; Write-Progress -Activity "Write Output" -Status "100% Complete:" -PercentComplete 100
$Output | ConvertTo-Json -Depth 100 | Set-Content .\output.json
    
#endregion
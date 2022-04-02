##############################################################################
## Import Milestone PSTools
##############################################################################

Import-Module .\pstools\MipSdkRedist\21.2.0\MipSdkRedist.psm1
Import-Module .\pstools\MilestonePSTools\21.2.6\MilestonePSTools.psm1

##############################################################################
## Initialize 
##############################################################################

## Get Server Name 
$ServerName = $env:COMPUTERNAME
## Get Current Time 
$TimeStamp = Get-Date

##############################################################################
## Determine Server type
##############################################################################

$MilestoneServices = Get-CimInstance -Query "Select * from Win32_Service Where Name like 'Milestone%' or Name like 'VideoOS%'" `
| Select-Object PSComputerName, Name, Description, State, StartMode, StartName, PathName `
| Sort-Object -Property PSComputerName, Name 


$IsManagementServer = ($MilestoneServices -match 'Milestone XProtect Management Server').Count > 0

$IsRecordingServer = ($MilestoneServices -match 'Milestone XProtect Recording Server').Count > 0

$HasMilestoneService = ($MilestoneServices -match 'Milestone').Count > 0

##############################################################################


if ($HasMilestoneService) {

    ## XProtect Cumulative Updates
    
    $XProtectCumulativeUpdates = [pscustomobject]@{
        CumulativeUpdates = $CumulativeUpdates
    }

    ##############################################################################
    ## Performance Counters
    ##############################################################################

    $samples = 10

    ## System RAM utilization

    $SystemRAMutilization = [pscustomobject]@{
        Samples = New-Object System.Collections.Generic.List[System.Object]
        Max     = ($TotalMemory | Where-Object PSComputerName -eq $s.ComputerName).TotalVisibleMemorySize
    }

    $TotalMemory = Get-CIMInstance Win32_OperatingSystem | Select-Object TotalVisibleMemorySize

    for ($i = 0; $i -lt $samples; $i++) {
        $mem_sample = Get-CIMInstance Win32_OperatingSystem | Select-Object @{Name = "UsedMemory"; Expression = { $_.TotalVisibleMemorySize - $_.FreePhysicalMemory } }
        
        $timestamp = (Get-Date).ToString('u')

        $ $SystemRAMutilization.Samples.Add(
            [pscustomobject]@{
                TimeStamp = $timestamp
                Value     = $mem_sample.UsedMemory
            }
        )
        
        Start-Sleep -Seconds 1
    }

    ## System CPU utilization

    
    $SystemCPUutilization = [pscustomobject]@{
        Samples = New-Object System.Collections.Generic.List[System.Object]
        Max     = 100
    }
    
    for ($i = 0; $i -lt $samples; $i++) {
        $cpu_sample = Get-CimInstance Win32_Processor | Select-Object LoadPercentage
        
        $timestamp = (Get-Date).ToString('u')
        
        $SystemCPUutilization.Samples.Add(
            [pscustomobject]@{
                TimeStamp = $timestamp
                Value     = $cpu_sample.LoadPercentage
            }
        )
        Start-Sleep -Seconds 1
    }
    
    ##############################################################################

    ## Hardware acceleration capability

    $HardwareaAcelerationCapability = [pscustomobject]@{
        GPUManufacturer = $GPUManufacturer
        GPUModel        = $GPUModel
        GPUMemory       = $GPUMemory
    }

    ## Antivirus presence
    $AntivirusPresence = [pscustomobject]@{
        Antivirus = $Antivirus
    }
}

if ($IsManagementServer) {

    Connect-ManagementServer -AcceptEula

    ## MilestoneXProtectVersion
    $MilestoneXProtectVersion = [pscustomobject]@{
        Version = $Version
        Product = $Product
    }

    ## Milestone Care Status
    $MilestoneCareStatus = [pscustomobject]@{
        CarePlus       = $CarePlus
        ExpirationDate = $ExpirationDate
    }

    ## Failover Configuration
    $FailoverConfiguration = [pscustomobject]@{

    }

}

if ($IsRecordingServer) {

    $RecordingServerLogPath = "{$env:ProgramData}\Milestone\Mileston Recording Server\Logs"

    ## Media Deletion Due to Low Disk Space
    #JOIN DATABASE LOGS AND FIND THE ERROR

    $MediaDeletionDuetoLowDiskSpace = [pscustomobject]@{
        DeletionDuetoLowDiskSpaceErrorList = $DeletionDuetoLowDiskSpaceErrorList
    }

    
    ## Media Deletion Due to Overflow
    # READ DEVICEHANDLING LOG AND FINF OVERFLOW ERROR

    $MediaDeletionDuetoOverflow = [pscustomobject]@{
        MediaDeletionDuetoOverflowList = $MediaDeletionDuetoOverflowList
    }


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

    $Output | ConvertTo-Json -Depth 100 | Set-Content C:\Users\sgiu\source\repos\Electron-MHC\json\obj.json
    
}
## Import Milestone PSTools
Import-Module .\pstools\MipSdkRedist\21.2.0\MipSdkRedist.psm1
Import-Module .\pstools\MilestonePSTools\21.2.6\MilestonePSTools.psm1

## Determine Server type

$MilestoneServices = Get-CimInstance -Query "Select * from Win32_Service Where Name like 'Milestone%' or Name like 'VideoOS%'" `
| Select-Object PSComputerName, Name, Description, State, StartMode, StartName, PathName `
| Sort-Object -Property PSComputerName, Name 


$IsManagementServer = ($MilestoneServices -match 'Milestone XProtect Management Server').Count > 0

$IsRecordingServer = ($MilestoneServices -match 'Milestone XProtect Recording Server').Count > 0

$HasAMilestoneService = ($MilestoneServices -match 'Milestone').Count > 0

if ($HasAMilestoneService) {

    ## XProtect Cumulative Updates

    ## System RAM utilization

    ## System CPU utilization

    ## Antivirus presence

}

if ($IsManagementServer) {

    Connect-ManagementServer 

    ## Milestone XProtect Version

    ## Milestone Care Status

    ## Failover Configuration

}

if ($IsRecordingServer) {

    ## Media Deletion Due to Low Disk Space

    ## Media Deletion Due to Overflow

    ## Hardware acceleration capability

}
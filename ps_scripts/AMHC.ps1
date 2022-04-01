## Import Milestone PSTools
Import-Module ..\pstools\MipSdkRedist\21.2.0\MipSdkRedist.psm1
Import-Module ..\pstools\MilestonePSTools\21.2.6\MilestonePSTools.psm1

## Determine Server type

Get-CimInstance -Query "Select * from Win32_Service Where Name like 'Milestone%' or Name like 'VideoOS%'" `
| Select-Object PSComputerName, Name, Description, State, StartMode, StartName, PathName `
| Sort-Object -Property PSComputerName, Name 





## build login 
#$User = "MEX-LAB\SGIU" 
#$PWord = ConvertTo-SecureString -String "Milestone1$" -AsPlainText -Force 
#$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, $PWord
#$Session = New-CimSession 



#Connect-ManagementServer 

## Milestone XProtect Version


## Milestone Care Status

## XProtect Cumulative Updates

## Media Deletion Due to Low Disk Space

## Media Deletion Due to Overflow

## System RAM utilization

## System CPU utilization

## Failover Configuration

## Hardware acceleration capability

## Antivirus presence
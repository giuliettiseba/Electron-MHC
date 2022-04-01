Import-Module "$PSScriptRoot\Bin\MipSdkRedist.dll"
$MipSdkPath = (Get-Item "$PSScriptRoot\Bin").FullName
if ($ENV:Path -notlike "*$MipSdkPath*") {
    $ENV:Path = "$($ENV:Path);$MipSdkPath"
}
Export-ModuleMember -Variable MipSdkPath
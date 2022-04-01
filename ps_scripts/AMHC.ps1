## build login 

$User = "MEX-LAB\SGIU" 
$PWord = ConvertTo-SecureString -String "Milestone1$" -AsPlainText -Force 
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, $PWord
$Session = New-CimSession -ComputerName 

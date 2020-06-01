$services = Get-Service

$services | Add-Member -MemberType NoteProperty -Name ApplianceUpdated -Value "No"

$services | Get-Member
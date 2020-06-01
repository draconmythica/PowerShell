Get-Service -DisplayName "Dell*" | Stop-Service
Get-Service -DisplayName "Dell*" | Set-Service -StartupType Disabled
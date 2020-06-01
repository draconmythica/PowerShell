Get-Service -DisplayName "Dell*" | Stop-Service
Start-Process -Filepath "C:\Program Files\Dell\Enterprise Edition\Server Configuration Tool\Credant.Configuration.exe" -Wait | Out-Null
Get-Service -DisplayName "Dell*" | Start-Service
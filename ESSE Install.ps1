IF ($ENV:PROCESSOR_ARCHITECTURE -eq 'AMD64'){

    Start-BitsTransfer -Source \\SoftwareShare\DDP\Client\x64\*.* -Destination C:\Temp\ -TransferType Download

    Start-Process "C:\Temp\DDPE_64bit_setup.exe" -ArgumentList '/s /v"SERVERHOSTNAME=server.domain.com POLICYPROXYHOSTNAME=server.domain.com DEVICESERVERURL=https://server.domain.com:8443/xapi/ REBOOT=ReallySuppress /l*v ShieldInstall.log /qn"' -wait

    Start-Process "C:\Temp\EMAgent_64bit_setup.exe" -ArgumentList '/s /v"CM_EDITION=1 SERVERHOST=server.domain.com SERVERPORT=8888 SECURITYSERVERHOST=server.domain.com SECURITYSERVERPORT=8443 ARPSYSTEMCOMPONENT=1 /norestart /qn"' -wait

    Start-Process "C:\Temp\ATP_CSF_Plugins_x64.msi" -ArgumentList '/qn REBOOT=ReallySuppress APPFOLDER="C:\Program Files\Dell\Dell Data Protection\Advanced Threat Protection\Plugins" ARPSYSTEMCOMPONENT=1 /l*v "C:\ProgramData\Dell\Dell Data Protection\Installer Logs\AdvancedThreatProtectionPlugins.msi.log' -wait

    Start-Process "C:\Temp\ATP_AgentSetup.exe" -ArgumentList '/s /norestart REBOOT=ReallySuppress APPFOLDER="C:\Program Files\Dell\Dell Data Protection\Advanced Threat Protection" ARPSYSTEMCOMPONENT=1 /l "C:\ProgramData\Dell\Dell Data Protection\Installer Logs\AdvancedThreatProtection.log' -wait
} ELSE {

    Start-BitsTransfer -Source \\SoftwareShare\DDP\Client\x86\*.* -Destination C:\Temp\ -TransferType Download

    Start-Process "C:\Temp\DDPE_32bit_setup.exe" -ArgumentList '/s /v"SERVERHOSTNAME=server.domain.com POLICYPROXYHOSTNAME=server.domain.com DEVICESERVERURL=https://server.domain.com:8443/xapi/ REBOOT=ReallySuppress /l*v ShieldInstall.log /qn"' -wait
    
    Start-Process "C:\Temp\EMAgent_32bit_setup.exe" -ArgumentList '/s /v"CM_EDITION=1 SERVERHOST=server.domain.com SERVERPORT=8888 SECURITYSERVERHOST=server.domain.com SECURITYSERVERPORT=8443 ARPSYSTEMCOMPONENT=1 /norestart /qn"' -wait

    Start-Process "C:\Temp\ATP_CSF_Plugins_x86.msi" -ArgumentList '/qn REBOOT=ReallySuppress APPFOLDER="C:\Program Files\Dell\Dell Data Protection\Advanced Threat Protection\Plugins" ARPSYSTEMCOMPONENT=1 /l*v "C:\ProgramData\Dell\Dell Data Protection\Installer Logs\AdvancedThreatProtectionPlugins.msi.log' -wait

    Start-Process "C:\Temp\ATP_AgentSetup.exe" -ArgumentList '/s /norestart REBOOT=ReallySuppress APPFOLDER="C:\Program Files\Dell\Dell Data Protection\Advanced Threat Protection" ARPSYSTEMCOMPONENT=1 /l "C:\ProgramData\Dell\Dell Data Protection\Installer Logs\AdvancedThreatProtectionlog' -wait

}
$Encryption=Get-WmiObject -Class Win32_Product | sort-object Name | select Name, Version | where { $_.Name -eq "Dell Encryption 64-bit" -or $_.Name -eq "Dell Encryption 32-bit"}
$EMAgent=Get-WmiObject -Class Win32_Product | sort-object Name | select Name, Version | where { $_.Name -eq "Dell Encryption Management Agent'"}
$ATP=Get-WmiObject -Class Win32_Product | sort-object Name | select Name, Version | where { $_.Name -eq "Cylance PROTECT"ù}
$ATPPlugins=Get-WmiObject -Class Win32_Product | sort-object Name | select Name, Version | where { $_.Name -eq "Cylance PROTECT Dell Plugins"ù}
If ($Encryption) {echo DDP Encryption Installed Successfully}else{echo DDP Encryption Failed To Install}
If ($EMAgent) {echo Client Security Framework Installed Successfully}else{echo Client Security Framework Failed To Install}
If ($ATP) {echo ATP Agent Installed Successfully}else{echo ATP Agent Failed To Install}
If ($ATPPlugins) {echo ATP Plugins Installed Successfully}else{echo ATP Plugins Failed To Install}

del "c:\users\public\desktop\Dell Data Security Console.lnk"


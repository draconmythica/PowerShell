#############Customizable Variables################
$LogDir="C:\ProgramData\Dell\Dell Data Protection\Installer Logs"
$InstallerSource="\\Server\temp\DDS_Client_8.17\"
$LocalCache="C:\Temp\"
$Serverhost="ddp.domain.com"
$UpgradeOlderThan = "8.17.1"
###################################################

#Create relevant folders if needed
If (-Not (Test-Path $LogDir)){md $LogDir}
If (-Not (Test-Path $LocalCache)){md $LocalCache}
$Log=$LogDir+"\ESSE_Script.log"

$Encryption=Get-WmiObject -Class Win32_Product | sort-object Name | select Name, Version | where { $_.Name -eq “Dell Encryption 64-bit” -or $_.Name -eq "Dell Encryption 32-bit"}
If ($Encryption -AND $Encryption.Version -gt $UpgradeOlderThan) {exit}else{echo "DDP Encryption Not Installed or Old Version" |Out-file $Log -Append}

#Check system type, copy down the relevant installers, and run the installation
IF ($ENV:PROCESSOR_ARCHITECTURE -eq 'AMD64'){ $Arch,$Bit="x64","64" } ELSE { $Arch,$Bit="x86","32"}

Start-BitsTransfer -Source ($InstallerSource + $Arch + "\*.*") -Destination "$LocalCache" -TransferType Download
Start-Process ($LocalCache + "DDPE_" + $Bit + "bit_setup.exe") -ArgumentList ('/s /v"SERVERHOSTNAME=' + $Serverhost + ' POLICYPROXYHOSTNAME=' + $Serverhost + ' DEVICESERVERURL=https://' + $Serverhost + ':8443/xapi/ REBOOT=ReallySuppress /l*v ShieldInstall.log /qn"') -wait
Start-Process ($LocalCache + "EMAgent_" + $Bit + "bit_setup.exe") -ArgumentList ('/s /v"CM_EDITION=1 SERVERHOST=' + $Serverhost + ' SERVERPORT=8888 SECURITYSERVERHOST=' + $Serverhost + ' SECURITYSERVERPORT=8443 ARPSYSTEMCOMPONENT=1 /norestart /qn"') -wait
Start-Process ($LocalCache + "ATP_CSF_Plugins_" + $Arch + ".msi") -ArgumentList ('/qn REBOOT=ReallySuppress APPFOLDER="C:\Program Files\Dell\Dell Data Protection\Advanced Threat Protection\Plugins" ARPSYSTEMCOMPONENT=1 /l*v "C:\ProgramData\Dell\Dell Data Protection\Installer Logs\AdvancedThreatProtectionPlugins.msi.log') -wait
Start-Process ($LocalCache + "ATP_AgentSetup.exe") -ArgumentList ('/s /norestart REBOOT=ReallySuppress APPFOLDER="C:\Program Files\Dell\Dell Data Protection\Advanced Threat Protection" ARPSYSTEMCOMPONENT=1 /l "C:\ProgramData\Dell\Dell Data Protection\Installer Logs\AdvancedThreatProtection.log') -wait

#Check each component and write installation status to previously specified log file
$Encryption=Get-WmiObject -Class Win32_Product | sort-object Name | select Name, Version | where { $_.Name -eq “Dell Encryption 64-bit” -or $_.Name -eq "Dell Encryption 32-bit"}
$EMAgent=Get-WmiObject -Class Win32_Product | sort-object Name | select Name, Version | where { $_.Name -eq “Dell Encryption Management Agent - x64” -or $_.Name -eq “Dell Encryption Management Agent - x86”}
$ATP=Get-WmiObject -Class Win32_Product | sort-object Name | select Name, Version | where { $_.Name -eq “Cylance PROTECT”}
$ATPPlugins=Get-WmiObject -Class Win32_Product | sort-object Name | select Name, Version | where { $_.Name -eq “Cylance PROTECT - Dell Plugins”}

#Check registry if you prefer that instead of WMI
#$ATP = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall | Get-ItemProperty | Where-Object {$_.DisplayName -eq "Cylance PROTECT" } | Select-Object -Property DisplayVersion
#$ATPPlugins = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall | Get-ItemProperty | Where-Object {$_.DisplayName -eq "Cylance PROTECT - Dell Plugins" } | Select-Object -Property DisplayVersion
#$EMAgent = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall | Get-ItemProperty | Where-Object {$_.DisplayName -eq "Dell Encryption Management Agent" } | Select-Object -Property DisplayVersion
#$Shield = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall | Get-ItemProperty | Where-Object {$_.DisplayName -eq "Dell Encryption 64-bit” -or $_.Name -eq "Dell Encryption 32-bit" } | Select-Object -Property DisplayVersion

If ($Encryption) {echo "DDP Encryption Installed Successfully" |Out-file $Log -Append}else{echo "DDP Encryption Failed To Install" |Out-file $Log -Append}
If ($EMAgent) {echo "Client Security Framework Installed Successfully" |Out-file $Log -Append}else{echo "Client Security Framework Failed To Install" |Out-file $Log -Append}
If ($ATP) {echo "ATP Agent Installed Successfully" |Out-file $Log -Append}else{echo "ATP Agent Failed To Install" |Out-file $Log -Append}
If ($ATPPlugins) {echo "ATP Plugins Installed Successfully" |Out-file $Log -Append}else{echo "ATP Plugins Failed To Install" |Out-file $Log -Append}

#Get rid of the desktop icon
del "c:\users\public\desktop\Dell Data Security Console.lnk"
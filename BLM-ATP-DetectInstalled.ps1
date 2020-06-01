#############Customizable Variables################
$LogDir="C:\ProgramData\Dell\Dell Data Protection\Installer Logs"
$InstallerSource="\\Server\Folder"
$LocalCache="C:\Temp\"
$Serverhost="ddp.domain.com"
$UpgradeOlderThan = "10.2"
###################################################

#Create relevant folders if needed
If (-Not (Test-Path $LogDir)){md $LogDir}
If (-Not (Test-Path $LocalCache)){md $LocalCache}
$Log=$LogDir+"\ESSE_Script.log"

#Get current installed version of Dell Encryption Management Agent (if any) and exit if we are already at or above the desired version
$EMAgent=Get-WmiObject -Class Win32_Product | sort-object Name | select Name, Version | where { $_.Name -eq "Dell Encryption Management Agent"}
If ($EMAgent -AND $EMAgent.Version -gt $UpgradeOlderThan) {exit}else{echo "Encryption Manager Not Installed or Old Version" |Out-file $Log -Append}

#Check system type and set bit level variables accordingly
IF ($ENV:PROCESSOR_ARCHITECTURE -eq 'AMD64'){ $Arch,$Bit="x64","64" } ELSE { $Arch,$Bit="x86","32"}

#Copy down the relevant installers for the specified source
Start-BitsTransfer -Source ($InstallerSource + $Arch + "\*.*") -Destination "$LocalCache" -TransferType Download

#Run the installation
Start-Process ($LocalCache + "EMAgent_" + $Bit + "bit_setup.exe") -ArgumentList ('/s /v"FEATURE=BLM CM_EDITION=1 SERVERHOST=' + $Serverhost + ' SERVERPORT=8888 SECURITYSERVERHOST=' + $Serverhost + ' SECURITYSERVERPORT=8443 ARPSYSTEMCOMPONENT=1 /norestart /qn"') -wait
Start-Process ($LocalCache + "ATP_CSF_Plugins_" + $Arch + ".msi") -ArgumentList ('/qn') -wait
Start-Process ($LocalCache + "ATP_AgentSetup.exe") -ArgumentList ('/s /norestart REBOOT=ReallySuppress APPFOLDER="C:\Program Files\Dell\Dell Data Protection\Advanced Threat Prevention" ARPSYSTEMCOMPONENT=1') -wait

#Query WMI for the installed versions of each component to verify it installed successfully, log the results
$EMAgent=Get-WmiObject -Class Win32_Product | sort-object Name | select Name, Version | where { $_.Name -eq "Dell Encryption Management Agent"}
$ATP=Get-WmiObject -Class Win32_Product | sort-object Name | select Name, Version | where { $_.Name -eq "Cylance PROTECT"}
$ATPPlugins=Get-WmiObject -Class Win32_Product | sort-object Name | select Name, Version | where { $_.Name -eq "Cylance PROTECT - Dell Plugins"}

If ($EMAgent) {echo "Client Security Framework Installed Successfully" |Out-file $Log -Append}else{echo "Client Security Framework Failed To Install" |Out-file $Log -Append}
If ($ATP) {echo "ATP Agent Installed Successfully" |Out-file $Log -Append}else{echo "ATP Agent Failed To Install" |Out-file $Log -Append}
If ($ATPPlugins) {echo "ATP Plugins Installed Successfully" |Out-file $Log -Append}else{echo "ATP Plugins Failed To Install" |Out-file $Log -Append}

#Get rid of the desktop icon
del "c:\users\public\desktop\Dell Data Security Console.lnk"
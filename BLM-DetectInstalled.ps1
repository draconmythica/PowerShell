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
$Log=$LogDir+"\BLM_Install_Script.log"

#Get current installed version of DDP Encryption (if any) and exit if we are already at or above the desired version
$Encryption=Get-WmiObject -Class Win32_Product | sort-object Name | select Name, Version | where { $_.Name -eq “Dell Encryption Management Agent”}
If ($Encryption -AND $Encryption.Version -gt $UpgradeOlderThan) {exit}else{echo "DDP Encryption Not Installed or Old Version" |Out-file $Log -Append}

#Run the installation
Start-Process ($LocalCache + "EMAgent_64bit_setup.exe") -ArgumentList ('/s /v"FEATURE=BLM CM_EDITION=1 SERVERHOST=' + $Serverhost + ' SERVERPORT=8888 SECURITYSERVERHOST=' + $Serverhost + ' SECURITYSERVERPORT=8443 ARPSYSTEMCOMPONENT=1 /norestart /l*v C:\ProgramData\Dell\Dell Data Protection\Installer Logs\EMAgent.log /qn"') -wait


#Query WMI for the installed versions of each component to verify it installed successfully, log the results
$EMAgent=Get-WmiObject -Class Win32_Product | sort-object Name | select Name, Version | where { $_.Name -eq “Dell Encryption Management Agent”}
If ($EMAgent) {echo "Client Security Framework Installed Successfully" |Out-file $Log -Append}else{echo "Client Security Framework Failed To Install" |Out-file $Log -Append}

#Get rid of the desktop icon
del "c:\users\public\desktop\Dell Data Security Console.lnk"
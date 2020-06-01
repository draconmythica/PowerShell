<#
Simple script to collect information about the TPM using the Get-Tpm module 
The output is to write the relevant bits out to the registry 
where they can be easily queried by SCCM/KACE/Whatever for inventory purposes
#>

#Get the information about the Tpm
$TPM=Get-Tpm

#Assign some variables we'll use to populate the registry values
$TPMAvailable=0
if($TPM.TpmPresent){$TPMAvailable=1}

$TPMActivated = 0
if($TPM.TpmReady){$TPMActivated=1}

#Define where in the registry we want the information
$regPath = "HKLM:\SOFTWARE\Tpm"

#Create that path if needed
If(!(test-path $regPath)) {New-Item -Path $regPath -Force}

#Add the values
New-ItemProperty -Path $regPath -Name "TpmAvailable" -Value $TPMAvailable -PropertyType DWORD -Force
New-ItemProperty -Path $regPath -Name "TpmActivated" -Value $TPMActivated -PropertyType DWORD -Force
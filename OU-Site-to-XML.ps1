#Import the excel module needed to read xlsx files
Import-Module PSExcel

#Define our source data
$sourceFileAmer = ".\OU Site List - Amer.xlsx"
$sourceFileApac = ".\OU Site List - APAC.xlsx"
$sourceFileEmea = ".\OU Site List - EMEA.xlsx"
$sites = New-Object System.Collections.ArrayList

#read in all our sites from the source excel files
foreach ($site in (Import-XLSX -Path $sourceFileAmer -RowStart 1)){$sites.Add($site)}
foreach ($site in (Import-XLSX -Path $sourceFileApac -RowStart 1)){$sites.Add($site)}
foreach ($site in (Import-XLSX -Path $sourceFileEmea -RowStart 1)){$sites.Add($site)}

$BeginXML = Get-Content -Path '.\XML_Sec_1.txt'
$MiddleXML = Get-Content -Path '.\XML_Sec_3.txt'
$EndXML = Get-Content -Path '.\XML_Sec_5.txt'
$output = ".\DellFactoryComplete.xml"

#Run through all the entries in the source data and format them into the choices in the xml file
$Choices=""
foreach ($site in $sites) {
    $Choices+='             <Choice Option="'+$site.SiteCode+'" Value="'+$site.SiteCode+'"/>'
    $Choices+="`n"
}

#Run through all the entries in the source data and format them into the actions in the xml file
$Actions=""
foreach ($site in $sites) {
    $Actions+='            <Action Type="TSVar" Name="OSDTimeZone" Condition=''"%SITECODE%" = "'+$site.SiteCode+'"''>'+$site.OSDTimeZone+'</Action>'+"`n"
    $Actions+='            <Action Type="TSVar" Name="UILanguage" Condition=''"%SITECODE%" = "'+$site.SiteCode+'"''>'+$site.UILanguage+'</Action>'+"`n"
    $Actions+='            <Action Type="TSVar" Name="KeyboardLocale" Condition=''"%SITECODE%" = "'+$site.SiteCode+'"''>'+$site.KeyboardLocale+'</Action>'+"`n"
    $Actions+='            <Action Type="TSVar" Name="OSDDomainName" Condition=''"%SITECODE%" = "'+$site.SiteCode+'"''>'+$site.OSDDomainName+'</Action>'+"`n"
    $Actions+='            <Action Type="TSVar" Name="OSDDomainOUName" Condition=''"%SITECODE%" = "'+$site.SiteCode+'"''>'+$site.OSDDomainOUName+'</Action>'+"`n"
    $Actions+="`n"
}

#Write all the content out to the file
$BeginXML | Out-File $output -Encoding utf8
$Choices | Out-File $output -Append -Encoding utf8
$MiddleXML | Out-File $output -Append -Encoding utf8
$Actions | Out-File $output -Append -Encoding utf8
$EndXML | Out-File $output -Append -Encoding utf8
## We need the correct token (should probably prompt the user for this)
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$cylanceToken = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your Cylance Token", "Select Token")

## Create a directory to work in
$path = "C:\Temp"
If(!(test-path $path)) { New-Item -ItemType Directory -Force -Path $path }

## Fetch the raw data reports
Invoke-WebRequest -Uri https://my-vs3.cylance.com/Reports/ThreatDataReportV1/threats/$cylanceToken -OutFile C:\Temp\threats.csv -UseBasicParsing -UserAgent "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36"
Invoke-WebRequest -Uri https://my-vs3.cylance.com/Reports/ThreatDataReportV1/devices/$cylanceToken -OutFile C:\Temp\devices.csv -UseBasicParsing -UserAgent "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36"
Invoke-WebRequest -Uri https://my-vs3.cylance.com/Reports/ThreatDataReportV1/cleared/$cylanceToken -OutFile C:\Temp\cleared.csv -UseBasicParsing -UserAgent "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36"
Invoke-WebRequest -Uri https://my-vs3.cylance.com/Reports/ThreatDataReportV1/events/$cylanceToken -OutFile C:\Temp\events.csv -UseBasicParsing -UserAgent "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36"
Invoke-WebRequest -Uri https://my-vs3.cylance.com/Reports/ThreatDataReportV1/indicators/$cylanceToken -OutFile C:\Temp\indicators.csv -UseBasicParsing -UserAgent "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36"
Invoke-WebRequest -Uri https://my-vs3.cylance.com/Reports/ThreatDataReportV1/memoryprotection/$cylanceToken -OutFile C:\Temp\memory.csv -UseBasicParsing -UserAgent "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36"

## Load the data into objects for easy access
$rawThreats = Import-Csv -Path C:\Temp\threats.csv
$rawCleared = Import-Csv -Path C:\Temp\cleared.csv
$devices = Import-Csv -Path C:\Temp\devices.csv
$memory = Import-Csv -Path C:\Temp\memory.csv

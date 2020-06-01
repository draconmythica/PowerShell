Import-Module PSExcel
## Define the three functions (Populate-Threats, Populate-OfflineDevices, and Populate-DuplicateDevices) 
## that are going to actually populate our PUP Template with data (we'll call these later)
function Populate-Threats ($tabName, $classification)
{
       
    Write-Host "Now populating $tabName"
	
    ## Loop through each active threat and see if it belongs on this tab if so add it to the dataset and then write the whole thing out
	$dataSet = $rawthreats | Foreach-Object -Process{
        
		if ($_.Classification -like  $classification -AND $_.'Last Found' -gt $startDate){
			
			
            ## Setup a few status values we'll define later
			$globalStatus = "Unknown"
            $thisCount = 1            
            $threatCount = $groupedThreats | Where-Object -Property Name -eq -Value $_.SHA256
            if ($threatCount.Count -eq 0) {$thisCount = 1} else {$thisCount = $threatCount.Count}

            ## Define global status based off whether the threat is globally quarantined, globally safelisted, or neither
			if ($_.'Global Quarantined' -eq "Yes") { $globalStatus = "Quarantine" } 
			elseif ($_.'Safelisted' -eq "Yes") { $globalStatus = "Safelist" }
			else { $globalStatus = "None" }
            
            ## Populate an object with the actual row data
            New-Object -TypeName PSObject -Property @{
                Actions = ""
                Notes = ""
                ProdName = $_.'Product Name'
                Device = $_.'DeviceName'
                FileName = $_.'File Name'
                FilePath = $_.'File Path'
                FileOwner = $_.'File Owner'
                GlobalList = $globalStatus
                FileStatus = $_.'File Status'
                Count = $thisCount
                Signed = $_.'Signed'
                Score = $_.'Cylance Score'
                AutoRun = $_.'Auto Run'
                Running = $_.'Running'
                Created = $_.'Create Time'
                FirstFound = $_.'First Found'
                LastFound = $_.'Last Found'
                SHA256 = $_.'SHA256'
            } | Select Actions, Notes, ProdName, Device, FileName, FilePath, FileOwner, GlobalList, FileStatus, Count, Signed, Score, AutoRun, Running, Created, FirstFound, LastFound, SHA256
        
		}
	}
    Write-Host Total number of $classification $dataSet.Length
    if( $dataSet ) { Export-XLSX -InputObject $dataSet -Path $reportPath -WorksheetName $tabName -Append }
    
}

function Populate-OfflineDevices
{
    
    Write-Host "Now Populating Offline Devices"
    $tabName = "Offline Devices"

    ## Loop through each device and see if it belongs on this tab, if so fill out the columns
	$dataSet = $devices | Foreach-Object -Process{
        if ($_.'Is Online' -eq "FALSE"){
            New-Object -TypeName PSObject -Property @{
                Name = $_.'Device Name'
                Serial = $_.'Serial Number'
                OS = $_.'OS Version'
                Agent = $_.'Agent Version'
                Policy = $_.'Policy'
                Zones = $_.'Zones'
                Mac = $_.'Mac Addresses'
                IP = $_.'IP Addresses'
                LastUser = $_.'Last Reported User'
                Created = $_.'Created'
                Offline = $_.'Offline Date'
            } | Select Name, Serial, OS, Agent, Policy, Zones, Mac, IP, LastUser, Created, Offline       
            
        }
    }
    Write-Host Total Offline Devices $dataSet.Length
    if( $dataSet ) { Export-XLSX -InputObject $dataSet -Path $reportPath -WorksheetName $tabName -Append }
}

function Populate-DuplicateDevices
{
    
    Write-Host "Now Populating Duplicate Devices"
    $tabName = "Duplicate Devices"

    $dups = $devices | Group-Object -Property 'Device Name' | Where-Object { $_.count -ge 2 }
    ## Loop through each duplicate and fill out the columns
	$dataSet = $dups | Foreach-Object -Process{
            New-Object -TypeName PSObject -Property @{
                Name = $_.Group[0].'Device Name'
                Serial = $_.Group[0].'Serial Number'
                OS = $_.Group[0].'OS Version'
                Agent = $_.Group[0].'Agent Version'
                Policy = $_.Group[0].'Policy'
                Zones = $_.Group[0].'Zones'
                Mac = $_.Group[0].'Mac Addresses'
                IP = $_.Group[0].'IP Addresses'
                LastUser = $_.Group[0].'Last Reported User'
                Created = $_.Group[0].'Created'
                Offline = $_.Group[0].'Offline Date'
            } | Select Name, Serial, OS, Agent, Policy, Zones, Mac, IP, LastUser, Created, Offline       
    }
    Write-Host Total Duplicate Devices $dataSet.Length
    if( $dataSet ) { Export-XLSX -InputObject $dataSet -Path $reportPath -WorksheetName $tabName -Append }
}

## Now we need a function that'll give the user an easy way to direct us to their PUP_template
function Get-FileName($filePurpose)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.filter = "Excel Spreadsheet (*.xlsx)| *.xlsx"
    $OpenFileDialog.Title = $filePurpose
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

function Get-StartDate()
{
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object Windows.Forms.Form

    $form.Text = 'Select start Date'
    $form.Size = New-Object Drawing.Size @(243,230)
    $form.StartPosition = 'CenterScreen'

    $calendar = New-Object System.Windows.Forms.MonthCalendar
    $calendar.ShowTodayCircle = $false
    $calendar.MaxSelectionCount = 1
    $form.Controls.Add($calendar)

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(38,165)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(113,165)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = 'Cancel'
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    $form.Topmost = $true

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $date = $calendar.SelectionStart
        Write-Host "Date selected: $($date.ToShortDateString())"
        return $date.ToShortDateString()
    }
}

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
#Invoke-WebRequest -Uri https://my-vs3.cylance.com/Reports/ThreatDataReportV1/events/$cylanceToken -OutFile C:\Temp\events.csv -UseBasicParsing -UserAgent "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36"
#Invoke-WebRequest -Uri https://my-vs3.cylance.com/Reports/ThreatDataReportV1/indicators/$cylanceToken -OutFile C:\Temp\indicators.csv -UseBasicParsing -UserAgent "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36"
Invoke-WebRequest -Uri https://my-vs3.cylance.com/Reports/ThreatDataReportV1/memoryprotection/$cylanceToken -OutFile C:\Temp\memory.csv -UseBasicParsing -UserAgent "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36"

## Load the data into objects for easy access
$rawThreats = Import-Csv -Path C:\Temp\threats.csv
$rawCleared = Import-Csv -Path C:\Temp\cleared.csv
$devices = Import-Csv -Path C:\Temp\devices.csv
$memory = Import-Csv -Path C:\Temp\memory.csv

$groupedThreats = $rawThreats | Group-Object -Property 'SHA256' | Where-Object { $_.count -ge 2 }

## Get path to the report template
$excel_file_path = Get-FileName("Select Template File")

## Ask the user where they'd like to save the final report
$SaveChooser = New-Object -TypeName System.Windows.Forms.SaveFileDialog
$SaveChooser.filter = "Excel Spreadsheet (*.xlsx)| *.xlsx"
$SaveChooser.Title = "Choose where to save output report"
$SaveChooser.ShowDialog()
$reportPath = $SaveChooser.Filename

$now = Get-Date -Format "dd-MMM-yyyy-hh-mm"
$startDate = Get-StartDate
Copy-item -path $excel_file_path -Destination $reportPath

## Execute the function we defined earlier to do the actual work
Populate-Threats -tabName "Adware" -classification "PUP - Adware"
Populate-Threats -tabName "Toolbar" -classification "PUP - Toolbar"
Populate-Threats -tabName "Keygen" -classification "PUP - Keygen"
Populate-Threats -tabName "Corrupt" -classification "PUP - Corrupt"
Populate-Threats -tabName "Hacking Tool" -classification "PUP - Hacking Tool"
Populate-Threats -tabName "Other" -classification "PUP - Other"
Populate-Threats -tabName "Portable Application" -classification "PUP - Portable Application"
Populate-Threats -tabName "Remote Access Tool" -classification "PUP - Remote Access Tool"
Populate-Threats -tabName "Scripting Tool" -classification "PUP - Scripting Tool"
Populate-Threats -tabName "Unclassified" -classification ""
Populate-Threats -tabName "Malware" -classification "Malware*"
Populate-Threats -tabName "Trusted Local" -classification "Trusted - Local"
Populate-Threats -tabName "All" -classification "*"
Populate-DuplicateDevices
Populate-OfflineDevices

Return "Complete"
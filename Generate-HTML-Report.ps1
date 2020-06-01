## We need the correct token (should probably prompt the user for this)

[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$cylanceToken = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your Cylance Token", "Select Token")

## Create a directory to work in
$reportPath = "C:\Temp"
If(!(test-path $reportPath)) { New-Item -ItemType Directory -Force -Path $reportPath }

## Fetch the raw data reports
Invoke-WebRequest -Uri https://my-vs3.cylance.com/Reports/ThreatDataReportV1/threats/$cylanceToken -OutFile $reportPath\threats.csv -UseBasicParsing -UserAgent "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36"
#Invoke-WebRequest -Uri https://my-vs3.cylance.com/Reports/ThreatDataReportV1/devices/$cylanceToken -OutFile $reportPath\devices.csv -UseBasicParsing -UserAgent "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36"
#Invoke-WebRequest -Uri https://my-vs3.cylance.com/Reports/ThreatDataReportV1/cleared/$cylanceToken -OutFile $reportPath\cleared.csv -UseBasicParsing -UserAgent "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36"
#Invoke-WebRequest -Uri https://my-vs3.cylance.com/Reports/ThreatDataReportV1/events/$cylanceToken -OutFile $reportPath\events.csv -UseBasicParsing -UserAgent "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36"
#Invoke-WebRequest -Uri https://my-vs3.cylance.com/Reports/ThreatDataReportV1/indicators/$cylanceToken -OutFile $reportPath\indicators.csv -UseBasicParsing -UserAgent "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36"
#Invoke-WebRequest -Uri https://my-vs3.cylance.com/Reports/ThreatDataReportV1/memoryprotection/$cylanceToken -OutFile $reportPath\memory.csv -UseBasicParsing -UserAgent "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36"

## Load the data into objects for easy access
$rawThreats = Import-Csv -Path $reportPath\threats.csv
#$rawSafelist = Import-Csv -Path $reportPath\cleared.csv
#$rawAgents = Import-Csv -Path $reportPath\devices.csv
#$rawExploitAttempts = Import-Csv -Path $reportPath\memory.csv

##Sometimes we only want threats seen in the last day
#$yesterday = (Get-Date).AddDays(-1)
#$rawThreats = $rawThreats | Where {(Get-Date $_.'Last Found') -ge $yesterday}

## Mark threats with a blank classification as Unclassified because blanks bite our ass later
## Need to clean up the rest of the classifications because I keep finding new ones and I'm a terrible person who is using unknown input to dynamically generate variables in the output because it's waaaaay faster than keeping the script updated with a list of all possible classifications
$rawThreats | Foreach-Object -Process{ 
    If($_.Classification -eq ""){$_.Classification = "Unclassified" }
    elseif($_.Classification -eq "Trusted - Local"){$_.Classification = "Trusted" }
    elseif($_.Classification -eq "PUP - Generic"){$_.Classification = "Generic PUP" }
    elseif($_.Classification -eq "Malware - Generic"){$_.Classification = "Generic Malware" }
    elseif($_.Classification -eq "PUP - Remote Access Tool"){$_.Classification = "Remote Access" }
    elseif($_.Classification -eq "Dual Use - Remote Access"){$_.Classification = "Remote Access" }
    
    if($_.Classification -like "PUP*") {$_.Classification = $_.Classification.Replace("PUP - ","")}
    elseif($_.Classification -like "Malware*") {$_.Classification = $_.Classification.Replace("Malware - ","")}
    elseif($_.Classification -like "Dual Use*") {$_.Classification = $_.Classification.Replace("Dual Use - ","")}
    
    $_.Classification = $_.Classification.Replace("-","")
    $_.Classification = $_.Classification.Replace(" ","") 
}

## Create variables to hold the threats formatted in a couple useful ways
$threatsByClass = $rawthreats | group-object -property Classification -AsHashTable
$threatsByHash = $rawThreats | group-object -property SHA256
$classifications = $threatsByClass.Keys | sort
ECHO $classifications
$totals = @()
$threatOutput = @{}

## Parse the threats and produce the totals for the chart as well as the dataset for each tab of the report
foreach ($key in $classifications) {
    $threats = $threatsByClass.$key
    #Group threats with this particular classification by their status so we can see how many of each type we have
    $unhandled = $threats | group-object -property 'File Status' | Where {$_.Name -like "unsafe" -OR $_.Name -like "abnormal"} | Select Count
    $quarantined = $threats | group-object -property 'File Status' | Where {$_.Name -like "quarantined"} | Select Count
    $waived = $threats | group-object -property 'File Status' | Where {$_.Name -like "waived"} | Select Count
    #Build the array of data for the chart, $totals holds the full array but we add to it line by line as we go through each classification
    $line = "" | select classification,unhandled,waived,quarantined
    $line.classification=$key
    $line.unhandled=$unhandled.Count
    $line.waived=$waived.Count
    $line.quarantined=$quarantined.Count
    $totals+= $line
    ECHO "Processing $key"

    #We need to build the data set that will be used to populate each threat tab. We have some fields slightly different from the raw data that need to be calculated first
    $dataSet = $threats | Foreach-Object -Process{
        $globalStatus = "Unknown" 
        if ($_.'Global Quarantined' -eq "Yes") { $globalStatus = "Quarantine" } 
	    elseif ($_.'Safelisted' -eq "Yes") { $globalStatus = "Safelist" }
		else { $globalStatus = "None" }
        
        $threatCount = $threatsByHash | Where-Object -Property Name -eq -Value $_.SHA256
        $priority = "Unknown"
        $pScore = 0
        if ($_.'Cylance Score' > 90) {$pScore++}
        if ($_.'Running' -eq "TRUE") {$pScore++}
        if ($_.'Ever Run' -eq "TRUE") {$pScore++}
        if ($_.'Auto Run' -eq "TRUE") {$pScore++}
        if ($_.'Detected By' -eq "Execution Control") {$pScore=$pScore+5}
        if ($pScore -gt 3) {$priority = "High"} 
        elseif ($pScore -gt 1) {$priority = "Medium"} 
        else {$priority = "Low"}

        
        New-Object -TypeName PSObject -Property @{
                FriendlyApplicationName = $_.'Product Name'
                DeviceName = $_.'DeviceName'
                FileName = $_.'File Name'
                FilePath = $_.'File Path'
                FileOwner = $_.'File Owner'
                GlobalList = $globalStatus
                FileStatus = $_.'File Status'
                Count = $threatCount.'Count'
                Priority = $priority
                Signed = $_.'Signed'
                Score = $_.'Cylance Score'
                AutoRun = $_.'Auto Run'
                Running = $_.'Running'
                DateCreated = $_.'Create Time'
                FirstFound = $_.'First Found'
                LastFound = $_.'Last Found'
                SHA256 = $_.'SHA256'
            } | Select FriendlyApplicationName, DeviceName, FileName, FileOwner, GlobalList, FileStatus, Count, Priority, Signed, Score, AutoRun, Running, DateCreated, FirstFound, LastFound, FilePath, SHA256
    }
    #Threat output is a big hash table, the classifications are the keys and each key holds the dataset containing all properly formatted threats from that classification
    $threatOutput[$key]=$dataset
}

## Ask the user where they'd like to save the final report
#$SaveChooser = New-Object -TypeName System.Windows.Forms.SaveFileDialog
#$SaveChooser.filter = "Web Page (*.html)| *.html"
#$SaveChooser.Title = "Choose where to save output report"
#$SaveChooser.ShowDialog()
#$savePath = $SaveChooser.Filename
$savePath = "$reportPath\ThreatReport.html"

#Define the text that will be written out as the html header, this includes the data for the threat chart, the css styling, and the javascript for the fancy tables
$HtmlHead = @'
<title>Threat Report</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />     
<link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.10.16/css/jquery.dataTables.min.css"/>
<link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/buttons/1.5.1/css/buttons.dataTables.min.css"/>
<link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/fixedheader/3.1.4/css/fixedHeader.dataTables.min.css"/>
<link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/select/1.2.6/css/select.dataTables.min.css"/>

<script src="https://code.jquery.com/jquery-3.3.1.js" integrity="sha256-2Kok7MbOyxpgUVvAk/HJ2jigOSYS2auK4Pfzbm7uH60=" crossorigin="anonymous"></script>
<script type="text/javascript" src="https://cdn.datatables.net/1.10.16/js/jquery.dataTables.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/buttons/1.5.1/js/dataTables.buttons.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/buttons/1.5.1/js/buttons.html5.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/buttons/1.5.1/js/buttons.print.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/buttons/1.5.1/js/buttons.colVis.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/fixedheader/3.1.4/js/dataTables.fixedHeader.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/select/1.2.6/js/dataTables.select.min.js"></script>    

<style>
/* Style the tab */
.tab {
    border: 1px solid #ccc;
    background-color: #f1f1f1;
}

/* Style the buttons that are used to open the tab content */
.tab button {
    background-color: inherit;
    float: left;
    border: none;
    outline: none;
    cursor: pointer;
    padding: 14px 16px;
    transition: 0.3s;
}

/* Change background color of buttons on hover */
.tab button:hover {
    background-color: #ddd;
}

/* Create an active/current tablink class */
.tab button.active {
    background-color: #ccc;
}

/* Style the tab content */
.tabcontent {
    display: none;
    padding: 6px 12px;
    border: 1px solid #ccc;
    border-top: none;
    border-right: none;
}
</style>
 

<script type="text/javascript"> 
	$(document).ready( function () {
'@
#This is initializing the datatables
foreach ($key in $classifications){
        $Htmlhead+=@"

        `$('#$key').DataTable({
            paging: false,
            select: true,
            fixedHeader: true,
            dom: 'Bfrtip',
            buttons: [
                'colvis', 'copy', 'csv', 'print'
            ]
        });

"@
}
#Not we're setting up the chart that appears at the top of the page
$Htmlhead+=@'
	} );
</script>
<script type="text/javascript" src="https://www.google.com/jsapi"></script>
    <script type="text/javascript">
      google.load("visualization", "1", {packages:["bar"]});
      google.setOnLoadCallback(drawChart);

      function drawChart() {
        var data = google.visualization.arrayToDataTable([
         ['Classification', 'Unhandled', 'Waived', 'Quarantined']
'@
# We populated an object called $totals earlier, now use it for the array of data for the chart
$totals | Foreach-Object -Process{
    $c=$_.classification
    #The chart doesn't scale the heading well, if the classification is too long it runs over top of the next label so we need to clean up the classifications a bit
    if($c.length -gt 14) {$c = $c.Substring(0,13)}    
    $u=$_.unhandled
    $w=$_.waived
    $q=$_.quarantined
    $HtmlHead+=@"
,
    ['$c',$u,$w,$q]
"@
}

#Append the rest of the header including chart options
$createdDate=Get-Date
$HtmlHead+=@"
]);

    var options = {
        chart: {title : 'Threats Chart $createdDate'},
        colors: ['#db4437', '#f4b400', '#4285f4']
    };

    var chart = new google.charts.Bar(document.getElementById('columnchart_material'));
    chart.draw(data, google.charts.Bar.convertOptions(options));
  }
    </script>


"@

#Define the text that will be written out as the html body
$HtmlBody = '<div id="columnchart_material" style="height: 500px;"></div>
            </br></br></br>
            <div class="tab">
            '

#Creates tabs in the body for each class of threat
$first=1
foreach ($key in $classifications) {
    #The first tab should be selected/shown by default, but we don't know which tab that'll be since we're generating them dynamically so we do a little extra jumping through hoops here
        if($first){
        $HtmlBody+=@"
            <button class="tablinks" onclick="openType(event,'$key DIV')" id="defaultOpen">$key</button>

"@
        $first=0
        }else{
        $HtmlBody+=@"
            <button class="tablinks" onclick="openType(event,'$key DIV')">$key</button>

"@
        }
        }

#Write the data to each tab
foreach ($key in $classifications) {
        $divTags=@"
        </div>
        <div id="$key DIV" class="tabcontent">

"@
        $HtmlBody+= $threatOutput[$key] | ConvertTo-Html -Fragment -PreContent $divTags
        $tableTag=('<table class="display compact cell-border" id="'+$key+'"><thead>')
        $HtmlBody=$HtmlBody.Replace("<table> <colgroup><col/><col/><col/><col/><col/><col/><col/><col/><col/><col/><col/><col/><col/><col/><col/><col/><col/></colgroup>",$tableTag)
        $HtmlBody=$HtmlBody.Replace("<th>SHA256</th></tr> <tr><td>","<th>SHA256</th></tr></thead> <tr><td>")
        }    

#Define the closing data that gets written out after the last tab is populated
$Closingtags = @'
<script>
function openType(evt, typeName) {
    // Declare all variables
    var i, tabcontent, tablinks;

    // Get all elements with class="tabcontent" and hide them
    tabcontent = document.getElementsByClassName("tabcontent");
    for (i = 0; i < tabcontent.length; i++) {
        tabcontent[i].style.display = "none";
    }

    // Get all elements with class="tablinks" and remove the class "active"
    tablinks = document.getElementsByClassName("tablinks");
    for (i = 0; i < tablinks.length; i++) {
        tablinks[i].className = tablinks[i].className.replace(" active", "");
    }

    // Show the current tab, and add an "active" class to the button that opened the tab
    document.getElementById(typeName).style.display = "block";
    evt.currentTarget.className += " active";
}
document.getElementById("defaultOpen").click();
</script>
'@


#Write the whole thing out to the chosen file
ConvertTo-Html -Head $HtmlHead -Body $HtmlBody -PostContent $Closingtags | Out-File -FilePath $savePath
Return "Complete"

#Read in the raw data from the csv
$rawUsers = Import-Csv -Path .\Downloads\export.csv

#Define where we want the output to go
$reportPath = "C:\Temp"
If(!(test-path $reportPath)) { New-Item -ItemType Directory -Force -Path $reportPath }
$savePath = "$reportPath\Output.txt"

#Group up the users by their managers and create a new array of managers and a string to hold our final output
$usersByManager = $rawUsers | group-object -property "Manager Email" -AsHashTable
$managers = $usersByManager.Keys | sort
$userOutput = ""


#Loop through all the managers and list the users
foreach ($key in $managers){
    $usersUnderThisManager = $usersByManager.$key
    $userOutput += "$key; "
    foreach ($user in $usersUnderThisManager){
            $userOutput += $user.'Display Name'
            $userOutput += ","
    }
    $userOutput = $userOutput.Substring(0,$userOutput.Length-1)
    $userOutput += "
    "
}

$userOutput | Out-File $savePath
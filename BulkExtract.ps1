Add-Type -AssemblyName System.Windows.Forms
$inputbrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$null = $inputbrowser.ShowDialog()
$input = $inputbrowser.SelectedPath

$outputbrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$null = $outputbrowser.ShowDialog()
$output = $outputbrowser.SelectedPath

$files = Get-ChildItem -Path $input -File -Recurse -Filter '*.rar'

foreach ($file in $files){
    $f = $file.FullName
    Start-Process -FilePath "7z.exe" -ArgumentList ('e "'+$f+'" -o"'+$output+'" -aou') -Wait
    }
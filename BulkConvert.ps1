Add-Type -AssemblyName System.Windows.Forms
$inputbrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$null = $inputbrowser.ShowDialog()
$input = $inputbrowser.SelectedPath

$outputbrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$null = $outputbrowser.ShowDialog()
$output = $outputbrowser.SelectedPath

$files = Get-ChildItem -Path $input -File -Recurse -Filter '*.vmdk'

foreach ($file in $files){
    $f = $file.FullName
	ConvertTo-MvmcVirtualHardDisk -SourceLiteralPath $f -VhdType DynamicHardDisk -VhdFormat vhdx -destination $output
    }
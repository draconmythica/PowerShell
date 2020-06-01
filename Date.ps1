function DatePicker($title) {
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
     
    $global:date = $null
    $form = New-Object Windows.Forms.Form
    $form.Size = New-Object Drawing.Size(233,190)
    $form.StartPosition = "CenterScreen"
    $form.KeyPreview = $true
    $form.FormBorderStyle = "FixedSingle"
    $form.Text = $title
    $calendar = New-Object System.Windows.Forms.MonthCalendar
    $calendar.ShowTodayCircle = $false
    $calendar.MaxSelectionCount = 1
    $form.Controls.Add($calendar)
    $form.TopMost = $true
     
    $form.add_KeyDown({
        if($_.KeyCode -eq "Escape") {
            $global:date = $false
            $form.Close()
        }
    })
     
    $calendar.add_DateSelected({
        $global:date = $calendar.SelectionStart
        $form.Close()
    })
     
    [void]$form.add_Shown($form.Activate())
    [void]$form.ShowDialog()
    return $global:date
}
 
Write-Host (DatePicker "Start Date")

$startDate = $(DatePicker "Start Date").ToShortDateString()
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Create Simplified OFX'
$form.Size = New-Object System.Drawing.Size(980, 700)
$form.StartPosition = 'CenterScreen'

$labelEntry = New-Object System.Windows.Forms.Label
$labelEntry.Text = 'Entries (required, one per line: MM/DD/YY,Amount,Years)'
$labelEntry.Location = New-Object System.Drawing.Point(15, 15)
$labelEntry.AutoSize = $true
$form.Controls.Add($labelEntry)

$textEntry = New-Object System.Windows.Forms.TextBox
$textEntry.Multiline = $true
$textEntry.ScrollBars = 'Vertical'
$textEntry.Size = New-Object System.Drawing.Size(930, 130)
$textEntry.Location = New-Object System.Drawing.Point(15, 40)
$form.Controls.Add($textEntry)

$entryHint = New-Object System.Windows.Forms.Label
$entryHint.ForeColor = [System.Drawing.Color]::DarkRed
$entryHint.Location = New-Object System.Drawing.Point(15, 175)
$entryHint.Size = New-Object System.Drawing.Size(930, 18)
$entryHint.Text = ''
$form.Controls.Add($entryHint)

$labelOutput = New-Object System.Windows.Forms.Label
$labelOutput.Text = 'Output file path'
$labelOutput.Location = New-Object System.Drawing.Point(15, 205)
$labelOutput.AutoSize = $true
$form.Controls.Add($labelOutput)

$textOutput = New-Object System.Windows.Forms.TextBox
$textOutput.Location = New-Object System.Drawing.Point(15, 230)
$textOutput.Size = New-Object System.Drawing.Size(820, 22)
$textOutput.Text = 'C:\\output\\import.ofx'
$form.Controls.Add($textOutput)

$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowse.Text = 'Browse...'
$buttonBrowse.Location = New-Object System.Drawing.Point(845, 227)
$buttonBrowse.Size = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($buttonBrowse)

$checkForce = New-Object System.Windows.Forms.CheckBox
$checkForce.Text = 'Force overwrite if output exists'
$checkForce.Location = New-Object System.Drawing.Point(15, 265)
$checkForce.AutoSize = $true
$form.Controls.Add($checkForce)

$checkWhatIf = New-Object System.Windows.Forms.CheckBox
$checkWhatIf.Text = 'WhatIf (preview only)'
$checkWhatIf.Location = New-Object System.Drawing.Point(260, 265)
$checkWhatIf.AutoSize = $true
$form.Controls.Add($checkWhatIf)

$checkVerbose = New-Object System.Windows.Forms.CheckBox
$checkVerbose.Text = 'Verbose logging'
$checkVerbose.Location = New-Object System.Drawing.Point(430, 265)
$checkVerbose.AutoSize = $true
$form.Controls.Add($checkVerbose)

$caution = New-Object System.Windows.Forms.Label
$caution.Text = 'CAUTION: This utility may overwrite existing files when -Force is selected.'
$caution.ForeColor = [System.Drawing.Color]::DarkRed
$caution.Location = New-Object System.Drawing.Point(15, 295)
$caution.Size = New-Object System.Drawing.Size(930, 20)
$form.Controls.Add($caution)

$runButton = New-Object System.Windows.Forms.Button
$runButton.Text = 'Run CreateSimplifiedOFX'
$runButton.Location = New-Object System.Drawing.Point(15, 320)
$runButton.Size = New-Object System.Drawing.Size(230, 34)
$form.Controls.Add($runButton)

$summaryLabel = New-Object System.Windows.Forms.Label
$summaryLabel.Text = 'Summary: Not run yet.'
$summaryLabel.Location = New-Object System.Drawing.Point(260, 328)
$summaryLabel.Size = New-Object System.Drawing.Size(685, 20)
$form.Controls.Add($summaryLabel)

$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Text = 'Execution output'
$outputLabel.Location = New-Object System.Drawing.Point(15, 365)
$outputLabel.AutoSize = $true
$form.Controls.Add($outputLabel)

$outputBox = New-Object System.Windows.Forms.RichTextBox
$outputBox.Location = New-Object System.Drawing.Point(15, 390)
$outputBox.Size = New-Object System.Drawing.Size(930, 180)
$outputBox.ReadOnly = $true
$outputBox.BackColor = [System.Drawing.Color]::White
$outputBox.WordWrap = $false
$form.Controls.Add($outputBox)

$generatedLabel = New-Object System.Windows.Forms.Label
$generatedLabel.Text = 'Generated outputs'
$generatedLabel.Location = New-Object System.Drawing.Point(15, 580)
$generatedLabel.AutoSize = $true
$form.Controls.Add($generatedLabel)

$generatedPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$generatedPanel.Location = New-Object System.Drawing.Point(15, 605)
$generatedPanel.Size = New-Object System.Drawing.Size(930, 50)
$generatedPanel.AutoScroll = $true
$form.Controls.Add($generatedPanel)

$openDialog = New-Object System.Windows.Forms.SaveFileDialog
$openDialog.Filter = 'OFX files (*.ofx)|*.ofx|All files (*.*)|*.*'
$openDialog.DefaultExt = 'ofx'
$openDialog.InitialDirectory = 'C:\\output'
$openDialog.OverwritePrompt = $true

$scriptPath = Join-Path -Path $PSScriptRoot -ChildPath 'CreateSimplifiedOFX.ps1'

function Add-OutputLine {
    param([string]$line)

    $color = [System.Drawing.Color]::Black
    if ($line -like 'SUCCESS:*') { $color = [System.Drawing.Color]::ForestGreen }
    elseif ($line -like 'WARNING:*') { $color = [System.Drawing.Color]::DarkGoldenrod }
    elseif ($line -like 'ERROR:*') { $color = [System.Drawing.Color]::Firebrick }

    $outputBox.SelectionStart = $outputBox.TextLength
    $outputBox.SelectionLength = 0
    $outputBox.SelectionColor = $color
    $outputBox.AppendText($line + [Environment]::NewLine)
    $outputBox.SelectionColor = $outputBox.ForeColor
    $outputBox.ScrollToCaret()
}

function Add-GeneratedLink {
    param([string]$path)

    if (-not (Test-Path -LiteralPath $path)) {
        return
    }

    foreach ($ctrl in $generatedPanel.Controls) {
        if ($ctrl.Tag -eq $path) {
            return
        }
    }

    $link = New-Object System.Windows.Forms.LinkLabel
    $link.Text = $path
    $link.Tag = $path
    $link.AutoSize = $true
    $link.Margin = New-Object System.Windows.Forms.Padding(3, 3, 3, 3)
    $link.add_LinkClicked({
        param($sender, $args)
        Start-Process -FilePath $sender.Tag
    })
    $generatedPanel.Controls.Add($link)
}

$buttonBrowse.Add_Click({
    $openDialog.FileName = [System.IO.Path]::GetFileName($textOutput.Text)
    if ($openDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textOutput.Text = $openDialog.FileName
    }
})

$runButton.Add_Click({
    $entryHint.Text = ''
    $summaryLabel.Text = 'Summary: Running...'

    if ([string]::IsNullOrWhiteSpace($textEntry.Text)) {
        $entryHint.Text = 'Entry input is required. Provide at least one line in the format MM/DD/YY,Amount,Years.'
        return
    }

    if ([string]::IsNullOrWhiteSpace($textOutput.Text)) {
        [System.Windows.Forms.MessageBox]::Show('Output path is required.', 'Validation', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    try {
        $fullOutputPath = [System.IO.Path]::GetFullPath($textOutput.Text)
        $windowsDirectory = [System.IO.Path]::GetFullPath($env:WINDIR)
        if ($fullOutputPath.StartsWith($windowsDirectory, [System.StringComparison]::OrdinalIgnoreCase)) {
            [System.Windows.Forms.MessageBox]::Show("Output cannot be inside the Windows directory: $windowsDirectory", 'Validation', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        $textOutput.Text = $fullOutputPath
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Validation', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) {
        [System.Windows.Forms.MessageBox]::Show("Script not found: $scriptPath", 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }

    $runButton.Enabled = $false
    $generatedPanel.Controls.Clear()
    $outputBox.Clear()

    $entryLines = @($textEntry.Lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    $cmdParts = New-Object System.Collections.Generic.List[string]
    [void]$cmdParts.Add('&')
    [void]$cmdParts.Add("`"$scriptPath`"")
    foreach ($line in $entryLines) {
        $escapedLine = $line.Replace('"', '`"')
        [void]$cmdParts.Add('-Entry')
        [void]$cmdParts.Add("`"$escapedLine`"")
    }

    $escapedOutput = $textOutput.Text.Replace('"', '`"')
    [void]$cmdParts.Add('-OutputPath')
    [void]$cmdParts.Add("`"$escapedOutput`"")

    if ($checkForce.Checked) { [void]$cmdParts.Add('-Force') }
    if ($checkWhatIf.Checked) { [void]$cmdParts.Add('-WhatIf') }
    if ($checkVerbose.Checked) { [void]$cmdParts.Add('-Verbose') }

    [void]$cmdParts.Add('2>&1')

    $command = $cmdParts -join ' '

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'powershell.exe'
    $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -Command $command"
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi

    $lines = New-Object System.Collections.Generic.List[string]
    $stdoutDone = $false
    $stderrDone = $false

    $process.add_OutputDataReceived({
        param($sender, $args)
        if ($null -eq $args.Data) {
            $script:stdoutDone = $true
        }
        else {
            $form.BeginInvoke([Action]{ Add-OutputLine -line $args.Data; [void]$lines.Add($args.Data) }) | Out-Null
        }
    })

    $process.add_ErrorDataReceived({
        param($sender, $args)
        if ($null -eq $args.Data) {
            $script:stderrDone = $true
        }
        else {
            $line = "ERROR: $($args.Data)"
            $form.BeginInvoke([Action]{ Add-OutputLine -line $line; [void]$lines.Add($line) }) | Out-Null
        }
    })

    $null = $process.Start()
    $process.BeginOutputReadLine()
    $process.BeginErrorReadLine()

    while (-not $process.HasExited -or -not $script:stdoutDone -or -not $script:stderrDone) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 80
    }

    $exitCode = $process.ExitCode

    foreach ($line in $lines) {
        if ($line -like 'OUTPUT:*') {
            $path = $line.Substring(7).Trim()
            Add-GeneratedLink -path $path
        }
    }

    $processed = 0
    $succeeded = 0
    $failed = 0
    $skipped = 0
    foreach ($line in $lines) {
        if ($line -match '^\s*Processed:\s*(\d+)') { $processed = [int]$matches[1] }
        elseif ($line -match '^\s*Succeeded:\s*(\d+)') { $succeeded = [int]$matches[1] }
        elseif ($line -match '^\s*Failed:\s*(\d+)') { $failed = [int]$matches[1] }
        elseif ($line -match '^\s*Skipped:\s*(\d+)') { $skipped = [int]$matches[1] }
    }

    if ($exitCode -eq 0) {
        $summaryLabel.ForeColor = [System.Drawing.Color]::ForestGreen
    }
    else {
        $summaryLabel.ForeColor = [System.Drawing.Color]::Firebrick
    }

    $summaryLabel.Text = "Summary: Processed=$processed; Succeeded=$succeeded; Failed=$failed; Skipped=$skipped; ExitCode=$exitCode"
    $runButton.Enabled = $true
})

[void]$form.ShowDialog()

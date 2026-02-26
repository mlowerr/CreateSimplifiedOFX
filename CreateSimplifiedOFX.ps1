[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [Alias('Entries')]
    [string[]]$Entry,

    [Parameter()]
    [Alias('OutFile')]
    [string]$OutputPath = 'C:\output\import.ofx',

    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$invariantCulture = [System.Globalization.CultureInfo]::InvariantCulture

function Write-Status {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('SUCCESS', 'WARNING', 'ERROR', 'OUTPUT')][string]$Prefix,
        [Parameter(Mandatory = $true)][string]$Message
    )

    Write-Output ("{0}: {1}" -f $Prefix, $Message)
}

function Convert-ToNormalizedEntryList {
    param([Parameter(Mandatory = $true)][string[]]$RawEntries)

    $normalized = New-Object System.Collections.Generic.List[string]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::Ordinal)

    foreach ($item in $RawEntries) {
        if ($null -eq $item) { continue }

        foreach ($line in ($item -split "`r?`n")) {
            $trimmed = $line.Trim()
            if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
            if ($seen.Add($trimmed)) {
                [void]$normalized.Add($trimmed)
            }
        }
    }

    return $normalized
}

function Convert-ToNormalizedDate {
    param([Parameter(Mandatory = $true)][string]$InputMmDdYy)

    $parts = $InputMmDdYy.Trim() -split '/'
    if ($parts.Count -ne 3) {
        throw "Invalid date '$InputMmDdYy'. Expected MM/DD/YY."
    }

    $mm = $parts[0].Trim().PadLeft(2, '0')
    $dd = $parts[1].Trim().PadLeft(2, '0')
    $yy = $parts[2].Trim().PadLeft(2, '0')

    if ($mm -notmatch '^\d{2}$' -or $dd -notmatch '^\d{2}$' -or $yy -notmatch '^\d{2}$') {
        throw "Invalid date '$InputMmDdYy'. Expected numeric MM/DD/YY."
    }

    $year = [int]("20$yy")
    try {
        return [datetime]::new($year, [int]$mm, [int]$dd, 0, 0, 0, [DateTimeKind]::Unspecified)
    }
    catch {
        throw "Invalid calendar date after normalization: $mm/$dd/$year"
    }
}

function Convert-ToEntry {
    param([Parameter(Mandatory = $true)][string]$Value)

    $parts = $Value -split ','
    if ($parts.Count -ne 3) {
        throw "Invalid -Entry '$Value'. Expected 'MM/DD/YY,Amount,Years'."
    }

    $startDate = Convert-ToNormalizedDate -InputMmDdYy $parts[0]

    $amountText = $parts[1].Trim()
    $amount = 0m
    if (-not [decimal]::TryParse($amountText, [System.Globalization.NumberStyles]::Number, $invariantCulture, [ref]$amount)) {
        throw "Invalid amount '$amountText' in -Entry '$Value'. Expected decimal format like 21.78."
    }

    $yearsText = $parts[2].Trim()
    $years = 0
    if (-not [int]::TryParse($yearsText, [ref]$years)) {
        throw "Invalid years '$yearsText' in -Entry '$Value'. Expected an integer."
    }
    if ($years -lt 0) {
        throw "Years must be >= 0 in -Entry '$Value'."
    }

    [pscustomobject]@{
        StartDate = $startDate
        Amount    = $amount
        Years     = $years
    }
}

function To-OfxDate {
    param([datetime]$DateValue)
    $DateValue.ToString('yyyyMMdd')
}

$summary = [ordered]@{
    Processed = 0
    Succeeded = 0
    Failed    = 0
    Skipped   = 0
}

try {
    $entryList = Convert-ToNormalizedEntryList -RawEntries $Entry
    if ($entryList.Count -eq 0) {
        throw 'No valid entry rows were provided. Add at least one row in format MM/DD/YY,Amount,Years.'
    }

    $fullOutputPath = [System.IO.Path]::GetFullPath($OutputPath)
    $outputDirectory = [System.IO.Path]::GetDirectoryName($fullOutputPath)
    if ([string]::IsNullOrWhiteSpace($outputDirectory)) {
        throw 'Output directory was not resolved from -OutputPath.'
    }

    $windowsDirectory = $null
    if (-not [string]::IsNullOrWhiteSpace($env:WINDIR)) {
        $windowsDirectory = [System.IO.Path]::GetFullPath($env:WINDIR)
    }

    if (-not [string]::IsNullOrWhiteSpace($windowsDirectory) -and $fullOutputPath.StartsWith($windowsDirectory, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Writing to the Windows directory is not allowed. Choose an output path outside: $windowsDirectory"
    }

    if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
        New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
        Write-Status -Prefix 'WARNING' -Message "Created output directory: $outputDirectory"
    }

    if ((Test-Path -LiteralPath $fullOutputPath -PathType Leaf) -and -not $Force.IsPresent) {
        throw "Output file already exists: $fullOutputPath. Use -Force to overwrite."
    }

    $transactions = New-Object System.Collections.Generic.List[object]
    foreach ($row in $entryList) {
        $summary.Processed++
        try {
            $parsed = Convert-ToEntry -Value $row
            for ($i = 0; $i -le $parsed.Years; $i++) {
                $dateValue = $parsed.StartDate.AddYears($i)
                $transactions.Add([pscustomobject]@{
                    Date   = $dateValue
                    Amount = $parsed.Amount
                    Name   = 'Annual Deposit'
                    Memo   = ('Deposit {0} on {1}' -f $parsed.Amount.ToString('N2', $invariantCulture), $dateValue.ToString('MM/dd/yyyy'))
                })
            }
            $summary.Succeeded++
            Write-Status -Prefix 'SUCCESS' -Message ("Parsed input row: $row")
        }
        catch {
            $summary.Failed++
            Write-Status -Prefix 'ERROR' -Message $_.Exception.Message
        }
    }

    if ($transactions.Count -eq 0) {
        throw 'No transactions were generated from inputs.'
    }

    $transactions = $transactions | Sort-Object -Property Date, Amount

    $dtStart = To-OfxDate -DateValue $transactions[0].Date
    $dtEnd = To-OfxDate -DateValue $transactions[$transactions.Count - 1].Date
    $dtAsof = $dtEnd

    $builder = New-Object System.Text.StringBuilder
    function Add-Line { param([string]$Value) [void]$builder.AppendLine($Value) }

    Add-Line 'OFXHEADER:100'
    Add-Line 'DATA:OFXSGML'
    Add-Line 'VERSION:102'
    Add-Line 'SECURITY:NONE'
    Add-Line 'ENCODING:UTF-8'
    Add-Line 'CHARSET:65001'
    Add-Line 'COMPRESSION:NONE'
    Add-Line 'OLDFILEUID:NONE'
    Add-Line 'NEWFILEUID:NONE'
    Add-Line ''
    Add-Line '<OFX>'
    Add-Line '  <SIGNONMSGSRSV1>'
    Add-Line '    <SONRS>'
    Add-Line '      <STATUS>'
    Add-Line '        <CODE>0'
    Add-Line '        <SEVERITY>INFO'
    Add-Line '      </STATUS>'
    Add-Line "      <DTSERVER>$dtAsof"
    Add-Line '      <LANGUAGE>ENG'
    Add-Line '    </SONRS>'
    Add-Line '  </SIGNONMSGSRSV1>'
    Add-Line '  <BANKMSGSRSV1>'
    Add-Line '    <STMTTRNRS>'
    Add-Line '      <TRNUID>1'
    Add-Line '      <STATUS>'
    Add-Line '        <CODE>0'
    Add-Line '        <SEVERITY>INFO'
    Add-Line '      </STATUS>'
    Add-Line '      <STMTRS>'
    Add-Line '        <CURDEF>USD'
    Add-Line '        <BANKACCTFROM>'
    Add-Line '          <BANKID>000000000'
    Add-Line '          <ACCTID>0'
    Add-Line '          <ACCTTYPE>IMPORT'
    Add-Line '        </BANKACCTFROM>'
    Add-Line '        <BANKTRANLIST>'
    Add-Line "          <DTSTART>$dtStart"
    Add-Line "          <DTEND>$dtEnd"

    $seq = 0
    foreach ($transaction in $transactions) {
        $seq++
        $posted = To-OfxDate -DateValue $transaction.Date
        $fitid = ('IMP{0}{1:0000}' -f $posted, $seq)

        Add-Line '          <STMTTRN>'
        Add-Line '            <TRNTYPE>CREDIT'
        Add-Line "            <DTPOSTED>$posted"
        Add-Line ('            <TRNAMT>{0}' -f $transaction.Amount.ToString('F2', $invariantCulture))
        Add-Line "            <FITID>$fitid"
        Add-Line "            <NAME>$($transaction.Name)"
        Add-Line "            <MEMO>$($transaction.Memo)"
        Add-Line '          </STMTTRN>'
    }

    Add-Line '        </BANKTRANLIST>'
    Add-Line '        <LEDGERBAL>'
    Add-Line '          <BALAMT>0.00'
    Add-Line "          <DTASOF>$dtAsof"
    Add-Line '        </LEDGERBAL>'
    Add-Line '      </STMTRS>'
    Add-Line '    </STMTTRNRS>'
    Add-Line '  </BANKMSGSRSV1>'
    Add-Line '</OFX>'

    $ofxText = $builder.ToString()

    if ($PSCmdlet.ShouldProcess($fullOutputPath, 'Write OFX output')) {
        Write-Status -Prefix 'WARNING' -Message 'Preparing to write output file. This may overwrite an existing file when -Force is used.'
        $utf8Bom = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllText($fullOutputPath, $ofxText, $utf8Bom)
        Write-Status -Prefix 'SUCCESS' -Message "Created OFX file with $($transactions.Count) transactions."
        Write-Status -Prefix 'OUTPUT' -Message $fullOutputPath
        Write-Status -Prefix 'WARNING' -Message 'Write operation completed.'
    }
    else {
        $summary.Skipped++
        Write-Status -Prefix 'WARNING' -Message "Skipped writing file because -WhatIf was used for $fullOutputPath"
    }
}
catch {
    Write-Status -Prefix 'ERROR' -Message $_.Exception.Message
    throw
}
finally {
    Write-Output 'SUMMARY:'
    Write-Output ("  Processed:  {0}" -f $summary.Processed)
    Write-Output ("  Succeeded: {0}" -f $summary.Succeeded)
    Write-Output ("  Failed:    {0}" -f $summary.Failed)
    Write-Output ("  Skipped:   {0}" -f $summary.Skipped)
}

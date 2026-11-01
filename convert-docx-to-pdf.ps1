# Batch DOCX -> PDF converter using Microsoft Word COM automation.
#
# This uses Word's own PDF export engine (the same one behind File > Export > Create PDF),
# so the output is a lossless, pixel-faithful rendering of the DOCX -- fonts, spacing,
# images, tables and page geometry all match exactly.
#
# A single Word instance is reused for the whole batch, since starting Word is the slow part.
#
# Usage: powershell -File convert-docx-to-pdf.ps1 -InputDir <dir> -OutputDir <dir>
# Every *.docx in InputDir is converted to a same-named *.pdf in OutputDir.

param(
    [Parameter(Mandatory = $true)][string]$InputDir,
    [Parameter(Mandatory = $true)][string]$OutputDir
)

$ErrorActionPreference = 'Stop'

# Word enum values
$wdFormatPDF = 17
$wdDoNotSaveChanges = 0
$wdAlertsNone = 0
$msoAutomationSecurityForceDisable = 3

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
}

$files = @(Get-ChildItem -Path $InputDir -Filter *.docx -File | Sort-Object Name)
if ($files.Count -eq 0) {
    Write-Output 'NO_FILES'
    exit 0
}

Write-Output "STARTING_WORD"
$word = $null
$converted = 0
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = $wdAlertsNone
    # Never run macros/content that could show a prompt and block the automation
    $word.AutomationSecurity = $msoAutomationSecurityForceDisable
    Write-Output "WORD_READY"

    foreach ($file in $files) {
        $inPath = $file.FullName
        # Must be a plain .NET string: with Join-Path's result, SaveAs2 never returns.
        [string]$outPath = [System.IO.Path]::Combine([string]$OutputDir, $file.BaseName + '.pdf')
        Write-Output "OPEN $($file.Name)"

        $doc = $null
        try {
            # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
            $doc = $word.Documents.Open($inPath, $false, $true, $false)
            Write-Output "OPENED $($file.Name)"

            # SaveAs2 with wdFormatPDF uses Word's native PDF export engine
            $doc.SaveAs2($outPath, $wdFormatPDF)
            Write-Output "SAVED $($file.Name)"
            $converted++
        }
        finally {
            if ($null -ne $doc) {
                $doc.Close($wdDoNotSaveChanges) | Out-Null
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
            }
        }
    }
}
finally {
    if ($null -ne $word) {
        # No argument: Quit(0) throws under Windows PowerShell 5.1 and leaves Word running.
        # Every document is already closed by this point.
        $word.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Output "CONVERTED=$converted"

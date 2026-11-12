# Build-time tool (needs Windows + Microsoft Word): reads the one text box in a certificate template
# that holds the changing fields ({NAME}, {Designation}, ...) and writes its geometry and fully
# resolved formatting to JSON. The browser PDF export lays that box out itself; everything else on
# the certificate comes from a Word-rendered background PDF (see tools/build-pdf-assets.js).
param([Parameter(Mandatory=$true)][string]$Docx, [Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($Docx, $false, $true, $false)
    $ps = $doc.PageSetup
    $target = $null
    foreach ($s in $doc.Shapes) {
        if ($s.TextFrame.HasText -ne 0) {
            if ($s.TextFrame.TextRange.Text -match '\{NAM[Ee]\}') { $target = $s; break }
        }
    }
    if ($null -eq $target) { throw 'No text box with a NAME field found.' }
    $paras = @()
    foreach ($p in $target.TextFrame.TextRange.Paragraphs) {
        $pf = $p.Format
        $runs = @()
        $cur = $null
        foreach ($c in $p.Range.Characters) {
            $t = $c.Text
            if ($t -eq "`r" -or $t -eq [string][char]7) { continue }
            $f = $c.Font
            $key = "$($f.Name)|$($f.Size)|$($f.Bold)|$($f.Italic)|$($f.Color)|$($f.Spacing)|$($f.Superscript)|$($f.Subscript)|$($f.AllCaps)|$($f.SmallCaps)|$($f.Underline)|$($f.TextColor.RGB)"
            if ($null -ne $cur -and $cur.key -eq $key) { $cur.text += $t; continue }
            if ($null -ne $cur) { $runs += $cur }
            $cur = [ordered]@{ key = $key; text = $t; font = $f.Name; size = [double]$f.Size; bold = ($f.Bold -ne 0); italic = ($f.Italic -ne 0);
                color = [int]$f.TextColor.RGB; spacing = [double]$f.Spacing; superscript = ($f.Superscript -ne 0); subscript = ($f.Subscript -ne 0);
                allCaps = ($f.AllCaps -ne 0); smallCaps = ($f.SmallCaps -ne 0); underline = ([int]$f.Underline) }
        }
        if ($null -ne $cur) { $runs += $cur }
        foreach ($r in $runs) { $r.Remove('key') }
        $paras += [ordered]@{
            align = [int]$pf.Alignment; spaceBefore = [double]$pf.SpaceBefore; spaceAfter = [double]$pf.SpaceAfter
            lineRule = [int]$pf.LineSpacingRule; lineSpacing = [double]$pf.LineSpacing
            leftIndent = [double]$pf.LeftIndent; rightIndent = [double]$pf.RightIndent; firstLineIndent = [double]$pf.FirstLineIndent
            runs = $runs
        }
    }
    $anchorPara = $target.Anchor.Paragraphs(1).Range
    $spec = [ordered]@{
        page = [ordered]@{ width = [double]$ps.PageWidth; height = [double]$ps.PageHeight; marginLeft = [double]$ps.LeftMargin; marginTop = [double]$ps.TopMargin }
        box = [ordered]@{
            left = [double]$target.Left; top = [double]$target.Top; width = [double]$target.Width; height = [double]$target.Height
            relH = [int]$target.RelativeHorizontalPosition; relV = [int]$target.RelativeVerticalPosition
            insetLeft = [double]$target.TextFrame.MarginLeft; insetTop = [double]$target.TextFrame.MarginTop
            insetRight = [double]$target.TextFrame.MarginRight; insetBottom = [double]$target.TextFrame.MarginBottom
            anchorParaTopOnPage = [double]$anchorPara.Information(6); anchorParaLeftOnPage = [double]$anchorPara.Information(5)
            anchorPage = [int]$anchorPara.Information(3)
        }
        paragraphs = $paras
    }
    $doc.Close(0)
    $spec | ConvertTo-Json -Depth 8 | Set-Content -Path $Out -Encoding UTF8
    Write-Output "OK $Out"
}
finally { $word.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) }

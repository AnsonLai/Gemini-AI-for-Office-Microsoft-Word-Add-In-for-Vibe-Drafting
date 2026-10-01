param(
    [Parameter(Mandatory = $true)][string]$SourcePath,
    [Parameter(Mandatory = $true)][string]$OutputDir
)

$ErrorActionPreference = 'Stop'
$word = $null
$document = $null
$runPath = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Force -Path $runPath | Out-Null

function Save-WordOpenXml([string]$name, $doc) {
    $path = Join-Path $runPath "$name.xml"
    $xml = [string]$doc.Content.WordOpenXML
    [System.IO.File]::WriteAllText($path, $xml, [System.Text.UTF8Encoding]::new($false))
    [pscustomobject]@{ name = $name; file = [System.IO.Path]::GetFileName($path); characters = $xml.Length }
}

function Replace-ParagraphText($doc, [int]$index, [string]$text) {
    $range = $doc.Paragraphs.Item($index).Range
    $range.Text = $text + [char]13
}

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($SourcePath, $false, $false, $false)
    $captures = [System.Collections.Generic.List[object]]::new()
    $captures.Add((Save-WordOpenXml '01-source-read-a' $document))
    $captures.Add((Save-WordOpenXml '02-source-read-b' $document))

    $document.TrackRevisions = $true
    $captures.Add((Save-WordOpenXml '03-tracking-on-no-edit' $document))
    $document.TrackRevisions = $false
    $captures.Add((Save-WordOpenXml '04-tracking-off-no-edit' $document))

    $document.TrackRevisions = $true
    Replace-ParagraphText $document 4 'Target paragraph P4 edited under tracking.'
    $captures.Add((Save-WordOpenXml '05-tracked-edit' $document))
    $document.RejectAllRevisions()
    $captures.Add((Save-WordOpenXml '06-after-reject-all' $document))
    $document.TrackRevisions = $false
    $captures.Add((Save-WordOpenXml '07-after-reject-tracking-off' $document))

    $document.Close(0)
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null
    $document = $null

    [ordered]@{
        wordVersion = [string]$word.Version
        wordBuild = [string]$word.Build
        captures = @($captures.ToArray())
        captureLimit = 'Seven captures only: source repeated reads, tracking toggles, tracked edit, Reject All, and tracking-off read.'
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $runPath 'word-capture-report.json') -Encoding UTF8
}
catch {
    [System.IO.File]::WriteAllText((Join-Path $runPath 'word-capture-error.txt'), ($_ | Out-String))
    throw
}
finally {
    if ($document) {
        try { $document.Close(0) | Out-Null } catch {}
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null
    }
    if ($word) {
        try { $word.Quit() } catch {}
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

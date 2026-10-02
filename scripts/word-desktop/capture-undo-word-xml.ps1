param(
    [Parameter(Mandatory = $true)][string]$SourcePath,
    [Parameter(Mandatory = $true)][string]$OutputDir
)
# Bounded Undo probe: no save, no restart. Edits are made in memory then reverted with Document.Undo().
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
    $undoResults = [ordered]@{}
    $captures.Add((Save-WordOpenXml '01-source-read-a' $document))
    $captures.Add((Save-WordOpenXml '02-source-read-b' $document))

    # Scenario A: tracked edit, then Undo
    $document.TrackRevisions = $true
    Replace-ParagraphText $document 4 'Target paragraph P4 edited under tracking.'
    $captures.Add((Save-WordOpenXml '05-tracked-edit' $document))
    $undoResults['trackedUndoReturned'] = [bool]$document.Undo(1)
    $undoResults['trackedRevisionCountAfterUndo'] = [int]$document.Revisions.Count
    $captures.Add((Save-WordOpenXml '11-after-undo' $document))
    $document.TrackRevisions = $false
    $captures.Add((Save-WordOpenXml '12-after-undo-tracking-off' $document))

    # Scenario B: untracked edit via InsertXML (the add-in's write path), then Undo
    $document.TrackRevisions = $false
    $flat = '<?xml version="1.0" standalone="yes"?><pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Target paragraph P4 rewritten via InsertXML.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>'
    $r = $document.Paragraphs.Item(4).Range
    $r.MoveEnd(1, -1) | Out-Null
    $r.InsertXML($flat)
    $captures.Add((Save-WordOpenXml '13-untracked-insertxml-edit' $document))
    $undoResults['untrackedUndoReturned'] = [bool]$document.Undo(1)
    $captures.Add((Save-WordOpenXml '14-after-untracked-undo' $document))
    $captures.Add((Save-WordOpenXml '15-after-untracked-undo-read-b' $document))

    $document.Close(0)
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null
    $document = $null

    [ordered]@{
        wordVersion = [string]$word.Version
        wordBuild = [string]$word.Build
        captures = @($captures.ToArray())
        undoResults = $undoResults
        captureLimit = 'Undo probe only: tracked Range.Text edit + Undo, untracked InsertXML edit + Undo. No save, no restart.'
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

param(
    [int]$TimeoutSeconds = 180,
    [switch]$Worker,
    [string]$TempDirectory
)

# Word-authored sources for the convert_to_list / change_list_format / continue_list host cases.
# Every source is built through Word COM, saved, reopened and observed in Word itself. The
# observation record (not the engine under test) is what the manifest expectations start from.
$ErrorActionPreference = 'Stop'
$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$fixtureDir = Join-Path $repoRoot 'tests\fixtures\agentic-lists'
$generationDir = Join-Path $repoRoot '.cache\reliability\agentic-list-action-fixture-generation'
New-Item -ItemType Directory -Force -Path $fixtureDir, $generationDir | Out-Null
$progressPath = Join-Path $generationDir 'progress.json'
$stdoutPath = Join-Path $generationDir 'stdout.log'
$stderrPath = Join-Path $generationDir 'stderr.log'
$observationPath = Join-Path $fixtureDir 'list-actions-source-observations.json'
$wns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

if (-not $Worker) {
    Remove-Item -LiteralPath $progressPath, $stdoutPath, $stderrPath -Force -ErrorAction SilentlyContinue
    $started = [DateTime]::UtcNow
    $tempDirectory = Join-Path $env:TEMP ("agentic-list-actions-{0}" -f [Guid]::NewGuid().ToString('N'))
    $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $PSCommandPath + '"'),
        '-Worker', '-TimeoutSeconds', [string]$TimeoutSeconds, '-TempDirectory', ('"' + $tempDirectory + '"'))
    $workerProcess = Start-Process -FilePath 'powershell.exe' -ArgumentList $arguments -WindowStyle Hidden -PassThru -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
    $workerProcess.Handle | Out-Null
    if (-not $workerProcess.WaitForExit($TimeoutSeconds * 1000)) {
        $progress = if (Test-Path -LiteralPath $progressPath) { Get-Content -LiteralPath $progressPath -Raw | ConvertFrom-Json } else { $null }
        if ($progress -and $progress.ownedWordPid -and ([DateTime]$progress.updatedAt).ToUniversalTime() -ge $started) {
            $ownedWord = Get-Process -Id ([int]$progress.ownedWordPid) -ErrorAction SilentlyContinue
            if ($ownedWord -and $ownedWord.StartTime.ToUniversalTime().ToString('o') -eq [string]$progress.ownedWordStartedAt) { Stop-Process -Id $ownedWord.Id -Force }
        }
        Stop-Process -Id $workerProcess.Id -Force -ErrorAction SilentlyContinue
        if (Test-Path -LiteralPath $stdoutPath) { Get-Content -LiteralPath $stdoutPath }
        throw "Word-authored list-action fixture generation timed out after $TimeoutSeconds seconds at stage '$($progress.stage)'."
    }
    if (Test-Path -LiteralPath $stdoutPath) { Get-Content -LiteralPath $stdoutPath }
    if (Test-Path -LiteralPath $stderrPath) { Get-Content -LiteralPath $stderrPath | Write-Host }
    if ($workerProcess.ExitCode -ne 0) { throw "Word-authored list-action fixture generator exited with code $($workerProcess.ExitCode)." }
    exit 0
}

function Save-Progress([string]$stage, $ownedWord = $null) {
    $progress = [ordered]@{
        updatedAt = [DateTime]::UtcNow.ToString('o'); stage = $stage
        ownedWordPid = if ($ownedWord) { [int]$ownedWord.Id } else { $null }
        ownedWordStartedAt = if ($ownedWord) { $ownedWord.StartTime.ToUniversalTime().ToString('o') } else { $null }
    }
    [System.IO.File]::WriteAllText($progressPath, (ConvertTo-Json -InputObject $progress -Depth 4), [System.Text.UTF8Encoding]::new($false))
}
function Normalize-WordText([string]$text) {
    return $text.Replace("`r", '').Replace([string][char]11, '').Replace([string][char]12, '').TrimEnd([char[]]"`n")
}

# Each source: plain text paragraphs, then Word list formatting applied to ranges (1-based paragraph numbers).
# kind: bullet = standard bullet gallery template 1, number = simple "1." numbered gallery template 1.
# Every list range starts a new list (ContinuePreviousList = false).
$specs = @(
    [ordered]@{ file = 'list-headers-source.docx'; case = 'L1'
        paragraphs = @('Lead-in paragraph.', '1. Alpha', 'Alpha body paragraph.', '2. Beta', 'Beta body paragraph.', '3. Gamma', 'Closing paragraph.')
        lists = @() }
    [ordered]@{ file = 'list-letter-headers-source.docx'; case = 'L2'
        paragraphs = @('Preamble paragraph.', 'C. Scope', 'Scope body paragraph.', 'D. Fees', 'Fees body paragraph.')
        lists = @() }
    [ordered]@{ file = 'list-bullets-source.docx'; case = 'L3'
        paragraphs = @('Lead-in paragraph.', 'Passport', 'Driver licence', 'Health card', 'Trailing paragraph.')
        lists = @(@{ first = 2; last = 4; kind = 'bullet' }) }
    [ordered]@{ file = 'list-two-lists-source.docx'; case = 'L4'
        paragraphs = @('Opening paragraph.', 'Alpha item', 'Beta item', 'Interlude paragraph.', 'Gamma item', 'Delta item', 'Closing paragraph.')
        lists = @(@{ first = 2; last = 3; kind = 'number' }, @{ first = 5; last = 6; kind = 'number' }) }
    [ordered]@{ file = 'list-changed-header-source.docx'; case = 'L5'
        paragraphs = @('Lead-in paragraph.', '1. Old title', 'Header body paragraph.')
        lists = @() }
    [ordered]@{ file = 'list-ambiguous-marker-source.docx'; case = 'L6'
        paragraphs = @('Lead-in paragraph.', 'I. Introduction', 'Introduction body paragraph.')
        lists = @() }
)
$expectedStrings = @{
    'list-bullets-source.docx' = @{ Passport = 'bullet'; 'Driver licence' = 'bullet'; 'Health card' = 'bullet' }
    'list-two-lists-source.docx' = @{ 'Alpha item' = '1.'; 'Beta item' = '2.'; 'Gamma item' = '1.'; 'Delta item' = '2.' }
}

function Get-ListObservation($paragraph, [int]$index) {
    $range = $paragraph.Range
    $listFormat = $range.ListFormat
    $listType = [int]$listFormat.ListType
    $isList = $listType -ne 0
    $row = [ordered]@{
        paragraphIndex = $index; text = Normalize-WordText ([string]$range.Text); isList = $isList; listType = $listType
        listString = $null; listLevel = $null; listValue = $null; numId = $null; numFmt = $null
    }
    if ($isList) {
        $row.listString = [string]$listFormat.ListString
        $row.listLevel = [int]$listFormat.ListLevelNumber
        $row.listValue = [int]$listFormat.ListValue
    }
    foreach ($object in @($listFormat, $range)) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($object) | Out-Null }
    return $row
}

function Add-PackageNumbering($docxPath, $rows) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($docxPath)
    try {
        $read = {
            param($name)
            $entry = $archive.GetEntry($name)
            if (-not $entry) { return $null }
            $reader = [System.IO.StreamReader]::new($entry.Open(), [System.Text.Encoding]::UTF8)
            try { return [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
        }
        [xml]$documentXml = & $read 'word/document.xml'
        $numberingXml = & $read 'word/numbering.xml'
        $namespaces = [System.Xml.XmlNamespaceManager]::new($documentXml.NameTable)
        $namespaces.AddNamespace('w', $wns)
        $paragraphs = $documentXml.SelectNodes('/w:document/w:body/w:p', $namespaces)
        $abstractByNum = @{}
        $formatByLevel = @{}
        if ($numberingXml) {
            $numberingNamespaces = [System.Xml.XmlNamespaceManager]::new($numberingXml.NameTable)
            $numberingNamespaces.AddNamespace('w', $wns)
            foreach ($num in $numberingXml.SelectNodes('/w:numbering/w:num', $numberingNamespaces)) {
                $abstractByNum[$num.GetAttribute('numId', $wns)] = $num.SelectSingleNode('./w:abstractNumId', $numberingNamespaces).GetAttribute('val', $wns)
            }
            foreach ($level in $numberingXml.SelectNodes('/w:numbering/w:abstractNum/w:lvl', $numberingNamespaces)) {
                $key = $level.ParentNode.GetAttribute('abstractNumId', $wns) + '|' + $level.GetAttribute('ilvl', $wns)
                $formatByLevel[$key] = $level.SelectSingleNode('./w:numFmt', $numberingNamespaces).GetAttribute('val', $wns)
            }
        }
        if ($paragraphs.Count -ne $rows.Count) { throw "Package paragraph count $($paragraphs.Count) differs from Word's $($rows.Count) in $docxPath." }
        for ($i = 0; $i -lt $paragraphs.Count; $i++) {
            $numNode = $paragraphs.Item($i).SelectSingleNode('./w:pPr/w:numPr/w:numId', $namespaces)
            $levelNode = $paragraphs.Item($i).SelectSingleNode('./w:pPr/w:numPr/w:ilvl', $namespaces)
            if ($numNode) {
                $numId = $numNode.GetAttribute('val', $wns)
                $rows[$i].numId = $numId
                $levelIndex = if ($levelNode) { $levelNode.GetAttribute('val', $wns) } else { '0' }
                $rows[$i].numFmt = $formatByLevel[$abstractByNum[$numId] + '|' + $levelIndex]
            }
        }
    } finally { $archive.Dispose() }
}

$word = $null
$ownedWord = $null
$document = $null
$verifyDocument = $null
try {
    Save-Progress 'startup'
    $localAppData = $env:LOCALAPPDATA
    $tempRoot = Join-Path $localAppData 'Temp'
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    $env:TEMP = $tempRoot; $env:TMP = $tempRoot
    foreach ($folder in @(
        (Join-Path $localAppData 'Microsoft\Windows\INetCache\Content.Word'),
        (Join-Path $localAppData 'Microsoft\Windows\INetCache\Content.MSO'),
        (Join-Path $localAppData 'Microsoft\Windows\Temporary Internet Files\Content.Word')
    )) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    if ([string]::IsNullOrWhiteSpace($TempDirectory)) { throw 'TempDirectory is required for the worker.' }
    New-Item -ItemType Directory -Force -Path $TempDirectory | Out-Null
    $existingWordPids = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $word = New-Object -ComObject Word.Application
    $createdWord = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $_.Id -notin $existingWordPids })
    if ($createdWord.Count -ne 1) { throw "Could not establish one generator-owned Word process (new process count: $($createdWord.Count)); no user Word session was changed." }
    $ownedWord = $createdWord[0]
    Save-Progress 'com-started' $ownedWord
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $bulletTemplate = $word.ListGalleries.Item(1).ListTemplates.Item(1)
    $numberTemplate = $word.ListGalleries.Item(2).ListTemplates.Item(1)

    $sources = @()
    foreach ($spec in $specs) {
        Save-Progress "author $($spec.file)" $ownedWord
        $document = $word.Documents.Add()
        $document.Range(0, 0).Text = $spec.paragraphs -join "`r"
        foreach ($list in $spec.lists) {
            $start = $document.Paragraphs.Item([int]$list.first).Range.Start
            $end = $document.Paragraphs.Item([int]$list.last).Range.End - 1
            $template = if ($list.kind -eq 'bullet') { $bulletTemplate } else { $numberTemplate }
            $document.Range($start, $end).ListFormat.ApplyListTemplateWithLevel($template, $false, 2, 2, 1)
        }
        $tempPath = Join-Path $TempDirectory $spec.file
        [string]$savePath = $tempPath
        [int]$saveFormat = 16
        $document.SaveAs2([ref]$savePath, [ref]$saveFormat)
        $document.Close(0) | Out-Null
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null
        $document = $null
        if (-not (Test-Path -LiteralPath $tempPath)) { throw "Word SaveAs did not create $($spec.file)." }
        $fixturePath = Join-Path $fixtureDir $spec.file
        Copy-Item -LiteralPath $tempPath -Destination $fixturePath -Force

        Save-Progress "reopen $($spec.file)" $ownedWord
        $verifyDocument = $word.Documents.OpenNoRepairDialog($fixturePath)
        $rows = @()
        for ($index = 1; $index -le $verifyDocument.Paragraphs.Count; $index++) { $rows += Get-ListObservation $verifyDocument.Paragraphs.Item($index) $index }
        $verifyDocument.Close(0) | Out-Null
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($verifyDocument) | Out-Null
        $verifyDocument = $null
        # Word may keep an export-only empty tail; the authored paragraphs come first.
        $rows = @($rows | Select-Object -First $spec.paragraphs.Count)
        Add-PackageNumbering $fixturePath $rows

        for ($i = 0; $i -lt $spec.paragraphs.Count; $i++) {
            if ($rows[$i].text -cne $spec.paragraphs[$i]) { throw "$($spec.file): paragraph $($i + 1) text differs after save/reopen: $($rows[$i].text)" }
            $inList = $false
            foreach ($list in $spec.lists) { if ($i + 1 -ge $list.first -and $i + 1 -le $list.last) { $inList = $true } }
            if ($rows[$i].isList -ne $inList) { throw "$($spec.file): list state differs at '$($rows[$i].text)'." }
        }
        $strings = $expectedStrings[$spec.file]
        if ($strings) {
            foreach ($row in $rows) {
                if ($strings.ContainsKey($row.text)) {
                    $expectedLabel = $strings[$row.text]
                    $labelOk = if ($expectedLabel -eq 'bullet') { $row.listType -in @(2, 6) -and $row.numFmt -eq 'bullet' } else { $row.listString -ceq $expectedLabel -and $row.numFmt -eq 'decimal' }
                    if (-not $labelOk) { throw "$($spec.file): Word-authored list differs at '$($row.text)': $($row | ConvertTo-Json -Compress)" }
                }
            }
        }
        if ($spec.case -eq 'L3') {
            $ids = @($rows | Where-Object { $_.isList } | ForEach-Object { $_.numId } | Select-Object -Unique)
            if ($ids.Count -ne 1) { throw 'L3 bullets must share one numId.' }
        }
        if ($spec.case -eq 'L4') {
            $first = @($rows | Where-Object { $_.text -in @('Alpha item', 'Beta item') } | ForEach-Object { $_.numId } | Select-Object -Unique)
            $second = @($rows | Where-Object { $_.text -in @('Gamma item', 'Delta item') } | ForEach-Object { $_.numId } | Select-Object -Unique)
            if ($first.Count -ne 1 -or $second.Count -ne 1 -or $first[0] -eq $second[0]) { throw "L4 needs two separate numbered lists; got [$first] and [$second]." }
        }
        $sources += [ordered]@{ case = $spec.case; file = $spec.file; paragraphs = $rows }
    }

    $report = [ordered]@{
        schemaVersion = 1
        provenance = [ordered]@{
            authoredBy = 'Microsoft Word desktop COM'
            wordVersion = [string]$word.Version; wordBuild = [string]$word.Build
            saveFormat = 'wdFormatDocumentDefault (16)'
            verifiedAfterSaveReopen = $true
            note = 'Typed-header sources are plain paragraphs; list sources use the Word bullet/number gallery templates with ContinuePreviousList=false.'
        }
        sources = @($sources)
    }
    [System.IO.File]::WriteAllText($observationPath, (ConvertTo-Json -InputObject $report -Depth 8), [System.Text.UTF8Encoding]::new($false))
    Write-Output "Created and reopened $($sources.Count) Word-authored list-action sources in $fixtureDir"
    Save-Progress 'complete' $ownedWord
}
finally {
    foreach ($documentToClose in @($verifyDocument, $document)) {
        if ($documentToClose) {
            try { $documentToClose.Close(0) | Out-Null } catch { }
            try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($documentToClose) | Out-Null } catch { }
        }
    }
    if ($word) {
        if ($ownedWord) { try { $word.Quit() | Out-Null } catch { } }
        try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null } catch { }
    }
    if ($ownedWord) {
        $stillRunning = Get-Process -Id $ownedWord.Id -ErrorAction SilentlyContinue
        if ($stillRunning -and $stillRunning.StartTime.ToUniversalTime().ToString('o') -eq $ownedWord.StartTime.ToUniversalTime().ToString('o')) {
            try { Stop-Process -Id $stillRunning.Id -Force } catch { }
        }
    }
    if ($TempDirectory -and (Test-Path -LiteralPath $TempDirectory)) { Remove-Item -LiteralPath $TempDirectory -Recurse -Force -ErrorAction SilentlyContinue }
}

param(
    [int]$TimeoutSeconds = 120,
    [switch]$Worker,
    [string]$TempDocxPath
)

$ErrorActionPreference = 'Stop'
$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$fixtureDir = Join-Path $repoRoot 'tests\fixtures\agentic-lists'
$generationDir = Join-Path $repoRoot '.cache\reliability\agentic-list-fixture-generation'
New-Item -ItemType Directory -Force -Path $fixtureDir | Out-Null
New-Item -ItemType Directory -Force -Path $generationDir | Out-Null
$progressPath = Join-Path $generationDir 'source-generation-progress.json'
$stdoutPath = Join-Path $generationDir 'source-generation-stdout.log'
$stderrPath = Join-Path $generationDir 'source-generation-stderr.log'
$sourcePath = Join-Path $fixtureDir 'nested-lists-source.docx'

if (-not $Worker) {
    if ($TimeoutSeconds -lt 1) { throw 'TimeoutSeconds must be positive.' }
    Remove-Item -LiteralPath $progressPath -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $stdoutPath, $stderrPath -Force -ErrorAction SilentlyContinue
    $started = [DateTime]::UtcNow
    $tempDocxPath = Join-Path $env:TEMP ("agentic-list-source-{0}.docx" -f [Guid]::NewGuid().ToString('N'))
    $arguments = @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $PSCommandPath + '"'),
        '-Worker', '-TimeoutSeconds', [string]$TimeoutSeconds, '-TempDocxPath', ('"' + $tempDocxPath + '"')
    )
    $workerProcess = Start-Process -FilePath 'powershell.exe' -ArgumentList $arguments -WindowStyle Hidden -PassThru -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
    $workerProcess.Handle | Out-Null
    if (-not $workerProcess.WaitForExit($TimeoutSeconds * 1000)) {
        $progress = if (Test-Path -LiteralPath $progressPath) { Get-Content -LiteralPath $progressPath -Raw | ConvertFrom-Json } else { $null }
        if ($progress -and $progress.ownedWordPid -and $progress.ownedWordStartedAt -and ([DateTime]$progress.updatedAt).ToUniversalTime() -ge $started) {
            $ownedWord = Get-Process -Id ([int]$progress.ownedWordPid) -ErrorAction SilentlyContinue
            if ($ownedWord -and $ownedWord.StartTime.ToUniversalTime().ToString('o') -eq [string]$progress.ownedWordStartedAt) {
                Stop-Process -Id $ownedWord.Id -Force
            }
        }
        Stop-Process -Id $workerProcess.Id -Force -ErrorAction SilentlyContinue
        if (Test-Path -LiteralPath $tempDocxPath) { Remove-Item -LiteralPath $tempDocxPath -Force -ErrorAction SilentlyContinue }
        if (Test-Path -LiteralPath $stdoutPath) { Get-Content -LiteralPath $stdoutPath }
        throw "Word-authored fixture generation timed out after $TimeoutSeconds seconds at stage '$($progress.stage)'."
    }
    if (Test-Path -LiteralPath $stdoutPath) { Get-Content -LiteralPath $stdoutPath }
    if (Test-Path -LiteralPath $stderrPath) { Get-Content -LiteralPath $stderrPath | Write-Host }
    if ($workerProcess.ExitCode -ne 0) { throw "Word-authored fixture generator exited with code $($workerProcess.ExitCode)." }
    exit 0
}

function Save-Progress([string]$stage, $ownedWord = $null) {
    $progress = [ordered]@{
        updatedAt = [DateTime]::UtcNow.ToString('o')
        stage = $stage
        ownedWordPid = if ($ownedWord) { [int]$ownedWord.Id } else { $null }
        ownedWordStartedAt = if ($ownedWord) { $ownedWord.StartTime.ToUniversalTime().ToString('o') } else { $null }
    }
    $json = ConvertTo-Json -InputObject $progress -Depth 4
    [System.IO.File]::WriteAllText($progressPath, $json, [System.Text.UTF8Encoding]::new($false))
}

function Normalize-WordText([string]$text) {
    return $text.Replace("`r", '').Replace([string][char]11, '').Replace([string][char]12, '').TrimEnd([char[]]"`n")
}

function Get-ListObservation($paragraph, [int]$index) {
    $range = $null
    $listFormat = $null
    $list = $null
    try {
        $range = $paragraph.Range
        $listFormat = $range.ListFormat
        $listType = [int]$listFormat.ListType
        $isList = $listType -ne 0
        $listId = $null
        $listString = $null
        $listLevel = $null
        $listValue = $null
        if ($isList) {
            try { $list = $listFormat.List; $listId = [int]$list.ID } catch { }
            try { $listString = [string]$listFormat.ListString } catch { }
            try { $listLevel = [int]$listFormat.ListLevelNumber } catch { }
            try { $listValue = [int]$listFormat.ListValue } catch { }
        }
        return [pscustomobject]@{
            paragraphIndex = $index
            text = Normalize-WordText ([string]$range.Text)
            isList = $isList
            listType = $listType
            listId = $listId
            listString = $listString
            listLevel = $listLevel
            listValue = $listValue
            bold = ([int]$range.Font.Bold -eq -1)
        }
    } finally {
        foreach ($object in @($list, $listFormat, $range)) {
            if ($object) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($object) | Out-Null }
        }
    }
}

function Get-PackageNumberingObservation([string]$docxPath) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($docxPath)
    try {
        $entry = $archive.GetEntry('word/document.xml')
        if (-not $entry) { throw 'Word-authored source is missing word/document.xml.' }
        $stream = $entry.Open()
        try {
            $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::UTF8)
            try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        } finally { $stream.Dispose() }
        $namespaces = [System.Xml.XmlNamespaceManager]::new($xml.NameTable)
        $namespaces.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        $rows = @()
        $paragraphs = $xml.SelectNodes('/w:document/w:body/w:p', $namespaces)
        for ($index = 0; $index -lt $paragraphs.Count; $index++) {
            $paragraph = $paragraphs.Item($index)
            $text = (($paragraph.SelectNodes('.//w:t', $namespaces) | ForEach-Object { $_.InnerText }) -join '')
            $numNode = $paragraph.SelectSingleNode('./w:pPr/w:numPr/w:numId', $namespaces)
            $levelNode = $paragraph.SelectSingleNode('./w:pPr/w:numPr/w:ilvl', $namespaces)
            $rows += [pscustomobject]@{
                paragraphIndex = $index + 1
                text = [string]$text
                numId = if ($numNode) { [string]$numNode.GetAttribute('val', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') } else { $null }
                ilvl = if ($levelNode) { [int]$levelNode.GetAttribute('val', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') } else { $null }
            }
        }
        $numberingEntry = $archive.GetEntry('word/numbering.xml')
        if (-not $numberingEntry) { throw 'Word-authored source is missing word/numbering.xml.' }
        $numberingStream = $numberingEntry.Open()
        try {
            $numberingReader = [System.IO.StreamReader]::new($numberingStream, [System.Text.Encoding]::UTF8)
            try { [xml]$numberingXml = $numberingReader.ReadToEnd() } finally { $numberingReader.Dispose() }
        } finally { $numberingStream.Dispose() }
        $numberingNamespaces = [System.Xml.XmlNamespaceManager]::new($numberingXml.NameTable)
        $numberingNamespaces.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        $abstractIds = @{}
        foreach ($num in $numberingXml.SelectNodes('/w:numbering/w:num', $numberingNamespaces)) {
            $numId = $num.GetAttribute('numId', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
            $abstractNode = $num.SelectSingleNode('./w:abstractNumId', $numberingNamespaces)
            if ($abstractNode) { $abstractIds[$numId] = $abstractNode.GetAttribute('val', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') }
        }
        $levelDetails = @{}
        foreach ($level in $numberingXml.SelectNodes('/w:numbering/w:abstractNum/w:lvl', $numberingNamespaces)) {
            $abstractId = $level.ParentNode.GetAttribute('abstractNumId', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
            $levelIndex = $level.GetAttribute('ilvl', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
            $formatNode = $level.SelectSingleNode('./w:numFmt', $numberingNamespaces)
            $textNode = $level.SelectSingleNode('./w:lvlText', $numberingNamespaces)
            $levelDetails["$abstractId|$levelIndex"] = [pscustomobject]@{
                numFmt = if ($formatNode) { $formatNode.GetAttribute('val', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') } else { $null }
                levelText = if ($textNode) { $textNode.GetAttribute('val', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') } else { $null }
            }
        }
        foreach ($row in $rows) {
            if ($row.numId) {
                $abstractId = $abstractIds[[string]$row.numId]
                $levelIndex = if ($null -eq $row.ilvl) { '0' } else { [string]$row.ilvl }
                $detail = $levelDetails["$abstractId|$levelIndex"]
                if ($detail) {
                    $row | Add-Member -NotePropertyName numFmt -NotePropertyValue $detail.numFmt -Force
                    $row | Add-Member -NotePropertyName levelText -NotePropertyValue $detail.levelText -Force
                }
            }
            if ($row.text -in @('Bullet Root A', 'Bullet Insertion Anchor', 'Bullet Untouched Tail', 'Bullet Root B')) {
                $row | Add-Member -NotePropertyName group -NotePropertyValue 'bullet-chain' -Force
            } elseif ($row.text -in @('Number Root A', 'Number Nested Anchor', 'Number Root B', 'Number Continued Item')) {
                $row | Add-Member -NotePropertyName group -NotePropertyValue 'number-continuation' -Force
            } elseif ($row.text -in @('Number Restart Root', 'Number Restart Nested')) {
                $row | Add-Member -NotePropertyName group -NotePropertyValue 'number-restart' -Force
            } else {
                $row | Add-Member -NotePropertyName group -NotePropertyValue $null -Force
            }
        }
        return $rows
    } finally { $archive.Dispose() }
}

function Get-ListGalleryCatalog($word) {
    $catalog = @()
    for ($galleryIndex = 1; $galleryIndex -le 3; $galleryIndex++) {
        $gallery = $word.ListGalleries.Item($galleryIndex)
        try {
            for ($templateIndex = 1; $templateIndex -le $gallery.ListTemplates.Count; $templateIndex++) {
                $template = $gallery.ListTemplates.Item($templateIndex)
                try {
                    $levels = @()
                    for ($levelIndex = 1; $levelIndex -le $template.ListLevels.Count; $levelIndex++) {
                        $level = $template.ListLevels.Item($levelIndex)
                        try {
                            $levels += [pscustomobject]@{
                                level = $levelIndex
                                numberStyle = [int]$level.NumberStyle
                                numberFormat = [string]$level.NumberFormat
                            }
                        } finally { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($level) | Out-Null }
                    }
                    $catalog += [pscustomobject]@{
                        gallery = $galleryIndex
                        template = $templateIndex
                        levelCount = $template.ListLevels.Count
                        levels = $levels
                    }
                } finally { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($template) | Out-Null }
            }
        } finally { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($gallery) | Out-Null }
    }
    return $catalog
}

$word = $null
$document = $null
$verifyDocument = $null
$ownedWord = $null
$numberTemplate = $null
$bulletTemplate = $null
try {
    Save-Progress 'startup'
    $localAppData = $env:LOCALAPPDATA
    if ([string]::IsNullOrWhiteSpace($localAppData)) { throw 'LOCALAPPDATA is unavailable; cannot prepare the Word workfile environment.' }
    $tempRoot = Join-Path $localAppData 'Temp'
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    $env:TEMP = $tempRoot
    $env:TMP = $tempRoot
    foreach ($folder in @(
        (Join-Path $localAppData 'Microsoft\Windows\INetCache\Content.Word'),
        (Join-Path $localAppData 'Microsoft\Windows\INetCache\Content.MSO'),
        (Join-Path $localAppData 'Microsoft\Windows\Temporary Internet Files\Content.Word')
    )) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    if ([string]::IsNullOrWhiteSpace($TempDocxPath)) { throw 'TempDocxPath is required for the worker.' }
    $existingWordPids = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $word = New-Object -ComObject Word.Application
    $createdWord = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $_.Id -notin $existingWordPids })
    if ($createdWord.Count -ne 1) {
        throw "Could not establish one generator-owned Word process (new process count: $($createdWord.Count)); no user Word session was changed."
    }
    $ownedWord = $createdWord[0]
    Save-Progress 'com-started' $ownedWord
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $galleryCatalog = Get-ListGalleryCatalog $word

    $paragraphText = @(
        'Untouched bold sentinel',
        'Plain paragraph before bullet list.',
        'Bullet Root A',
        'Bullet Insertion Anchor',
        'Bullet Untouched Tail',
        'Bullet Root B',
        'Plain transition after bullet list.',
        'Number Root A',
        'Number Nested Anchor',
        'Number Root B',
        'Plain transition before continuation.',
        'Number Continued Item',
        'Plain transition before restart.',
        'Number Restart Root',
        'Number Restart Nested',
        'Plain transition after restart.',
        'Header First Candidate',
        'Plain paragraph between headers.',
        'Header Second Candidate',
        'Plain paragraph after headers.',
        '1. Supported Header First',
        'Plain paragraph between supported headers.',
        '2. Supported Header Second',
        'Plain paragraph after supported headers.'
    )

    Save-Progress 'document-add' $ownedWord
    $document = $word.Documents.Add()
    $document.Range(0, 0).Text = $paragraphText -join "`r"
    $document.Paragraphs.Item(1).Range.Font.Bold = -1

    $bulletStart = $document.Paragraphs.Item(3).Range.Start
    $bulletEnd = $document.Paragraphs.Item(6).Range.End - 1
    $bulletRange = $document.Range($bulletStart, $bulletEnd)
    $bulletTemplate = $word.ListGalleries.Item(3).ListTemplates.Item(3)
    $bulletRange.ListFormat.ApplyListTemplateWithLevel($bulletTemplate, $false, 2, 2, 1)
    $document.Paragraphs.Item(4).Range.ListFormat.ListIndent()

    $numberStart = $document.Paragraphs.Item(8).Range.Start
    $numberEnd = $document.Paragraphs.Item(10).Range.End - 1
    $numberRange = $document.Range($numberStart, $numberEnd)
    $numberTemplate = $word.ListGalleries.Item(3).ListTemplates.Item(2)
    $numberRange.ListFormat.ApplyListTemplateWithLevel($numberTemplate, $false, 2, 2, 1)
    $document.Paragraphs.Item(9).Range.ListFormat.ListIndent()

    # Reuse Word's numbered-list template to create a continuation across a
    # plain paragraph, then explicitly start a separate list for the restart.
    $document.Paragraphs.Item(12).Range.ListFormat.ApplyListTemplateWithLevel($numberTemplate, $true, 2, 2, 1)
    $restartStart = $document.Paragraphs.Item(14).Range.Start
    $restartEnd = $document.Paragraphs.Item(15).Range.End - 1
    $restartRange = $document.Range($restartStart, $restartEnd)
    $restartRange.ListFormat.ApplyListTemplateWithLevel($numberTemplate, $false, 2, 2, 1)
    $document.Paragraphs.Item(15).Range.ListFormat.ListIndent()

    Save-Progress 'save-source' $ownedWord
    [string]$nativeSavePath = $TempDocxPath
    [int]$nativeFormat = 16
    $document.SaveAs2([ref]$nativeSavePath, [ref]$nativeFormat)
    $document.Close(0) | Out-Null
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null
    $document = $null
    $numberTemplate = $null
    if (-not (Test-Path -LiteralPath $TempDocxPath)) { throw 'Word SaveAs returned without creating the source document.' }
    Copy-Item -LiteralPath $TempDocxPath -Destination $sourcePath -Force
    Remove-Item -LiteralPath $TempDocxPath -Force

    Save-Progress 'reopen-source' $ownedWord
    $packageObservations = Get-PackageNumberingObservation $sourcePath
    $verifyDocument = $word.Documents.OpenNoRepairDialog($sourcePath)
    $observations = @()
    for ($index = 1; $index -le $verifyDocument.Paragraphs.Count; $index++) {
        $observations += Get-ListObservation $verifyDocument.Paragraphs.Item($index) $index
    }
    $byText = @{}
    foreach ($row in $observations) { $byText[[string]$row.text] = $row }
    $reportPath = Join-Path $fixtureDir 'source-observations.json'
    $partialReport = [ordered]@{
        schemaVersion = 1
        provenance = [ordered]@{
            authoredBy = 'Microsoft Word desktop COM'
            wordVersion = [string]$word.Version
            wordBuild = [string]$word.Build
            sourceFile = 'nested-lists-source.docx'
            saveFormat = 'wdFormatDocumentDefault (16)'
            verifiedAfterSaveReopen = $true
            structureAssertions = 'pending'
        }
        paragraphCount = $observations.Count
        paragraphs = @($observations)
        packageNumbering = @($packageObservations)
        galleryTemplateCatalog = @($galleryCatalog)
    }
    [System.IO.File]::WriteAllText($reportPath, (ConvertTo-Json -InputObject $partialReport -Depth 8), [System.Text.UTF8Encoding]::new($false))
    foreach ($expected in @(
        @{ text = 'Bullet Root A'; level = 1 },
        @{ text = 'Bullet Insertion Anchor'; level = 2 },
        @{ text = 'Bullet Untouched Tail'; level = 1 },
        @{ text = 'Bullet Root B'; level = 1 },
        @{ text = 'Number Root A'; level = 1 },
        @{ text = 'Number Nested Anchor'; level = 2 },
        @{ text = 'Number Root B'; level = 1 },
        @{ text = 'Number Continued Item'; level = 1 },
        @{ text = 'Number Restart Root'; level = 1 },
        @{ text = 'Number Restart Nested'; level = 2 }
    )) {
        $row = $byText[$expected.text]
        if (-not $row -or -not $row.isList -or $row.listLevel -ne $expected.level) {
            throw "Word-authored list structure differs at '$($expected.text)': $($row | ConvertTo-Json -Compress)"
        }
    }
    if (-not $byText['Untouched bold sentinel'].bold) { throw 'Word-authored bold sentinel lost direct formatting.' }
    foreach ($text in @('Header First Candidate', 'Header Second Candidate', '1. Supported Header First', '2. Supported Header Second')) {
        if ($byText[$text].isList) { throw "Header fixture must be a plain paragraph, not a Word list item: $text" }
    }
    $packageByText = @{}
    foreach ($row in $packageObservations) { $packageByText[[string]$row.text] = $row }
    $continuationNumId = $packageByText['Number Root A'].numId
    if (-not $continuationNumId -or
        $continuationNumId -ne $packageByText['Number Nested Anchor'].numId -or
        $continuationNumId -ne $packageByText['Number Root B'].numId -or
        $continuationNumId -ne $packageByText['Number Continued Item'].numId -or
        $continuationNumId -eq $packageByText['Number Restart Root'].numId -or
        $packageByText['Number Nested Anchor'].ilvl -ne 1 -or
        $packageByText['Number Restart Root'].numId -ne $packageByText['Number Restart Nested'].numId -or
        $packageByText['Number Restart Nested'].ilvl -ne 1) {
        throw "Word OOXML numbering identity differs for continuation/restart: $($packageByText['Number Root A'] | ConvertTo-Json -Compress), $($packageByText['Number Continued Item'] | ConvertTo-Json -Compress), $($packageByText['Number Restart Root'] | ConvertTo-Json -Compress)"
    }
    if ($byText['Number Continued Item'].listValue -le $byText['Number Root B'].listValue -or $byText['Number Restart Root'].listValue -ne 1) {
        throw "Word did not continue/restart the numbered values: $($byText['Number Root B'] | ConvertTo-Json -Compress), $($byText['Number Continued Item'] | ConvertTo-Json -Compress), $($byText['Number Restart Root'] | ConvertTo-Json -Compress)"
    }
    foreach ($text in @('Bullet Root A', 'Bullet Insertion Anchor', 'Bullet Untouched Tail', 'Bullet Root B')) {
        if (-not $packageByText[$text].numId) { throw "Word OOXML bullet paragraph has no direct numId: $text" }
    }
    if ($packageByText['Bullet Root A'].numId -ne $packageByText['Bullet Insertion Anchor'].numId -or
        $packageByText['Bullet Root A'].numId -ne $packageByText['Bullet Untouched Tail'].numId -or
        $packageByText['Bullet Root A'].numId -ne $packageByText['Bullet Root B'].numId -or
        $packageByText['Bullet Insertion Anchor'].ilvl -ne 1 -or
        $packageByText['Bullet Untouched Tail'].ilvl -ne 0) {
        throw "Word OOXML bullet nesting or list identity differs: $($packageByText['Bullet Root A'] | ConvertTo-Json -Compress), $($packageByText['Bullet Insertion Anchor'] | ConvertTo-Json -Compress), $($packageByText['Bullet Untouched Tail'] | ConvertTo-Json -Compress)"
    }
    foreach ($text in @('Bullet Root A', 'Bullet Insertion Anchor', 'Bullet Untouched Tail', 'Bullet Root B')) {
        if ($packageByText[$text].numFmt -ne 'bullet') { throw "Word-authored bullet lost bullet numbering format at '$text': $($packageByText[$text] | ConvertTo-Json -Compress)" }
    }
    if ($packageByText['Number Nested Anchor'].numFmt -ne 'decimal' -or $packageByText['Number Nested Anchor'].levelText -ne '%1.%2.') {
        throw "Word-authored nested numbering format differs: $($packageByText['Number Nested Anchor'] | ConvertTo-Json -Compress)"
    }

    $report = [ordered]@{
        schemaVersion = 1
        provenance = [ordered]@{
            authoredBy = 'Microsoft Word desktop COM'
            wordVersion = [string]$word.Version
            wordBuild = [string]$word.Build
            sourceFile = 'nested-lists-source.docx'
            saveFormat = 'wdFormatDocumentDefault (16)'
            verifiedAfterSaveReopen = $true
            structureAssertions = 'passed'
        }
        paragraphCount = $observations.Count
        paragraphs = @($observations)
        packageNumbering = @($packageObservations)
        galleryTemplateCatalog = @($galleryCatalog)
        logicalGroups = [ordered]@{
            bullets = 'Bullet Root A, Bullet Insertion Anchor, Bullet Untouched Tail, Bullet Root B share one direct numId; anchor is ilvl=1 and tail is ilvl=0'
            numberedContinuation = 'Number Root A, Number Nested Anchor, Number Root B, Number Continued Item share one direct numId; nested anchor is ilvl=1'
            numberedRestart = 'Number Restart Root and Number Restart Nested share a direct numId distinct from the continuation chain; nested item is ilvl=1 and root restarts at 1'
        }
    }
    $json = ConvertTo-Json -InputObject $report -Depth 8
    [System.IO.File]::WriteAllText($reportPath, $json, [System.Text.UTF8Encoding]::new($false))
    Write-Output "Created and reopened Word-authored list source: $sourcePath"
    Write-Output "Source observations: $(Join-Path $fixtureDir 'source-observations.json')"
    Save-Progress 'complete' $ownedWord
}
finally {
    foreach ($documentToClose in @($verifyDocument, $document)) {
        if ($documentToClose) {
            try { $documentToClose.Close(0) | Out-Null } catch { }
            try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($documentToClose) | Out-Null } catch { }
        }
    }
    foreach ($template in @($bulletTemplate, $numberTemplate)) {
        if ($template) { try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($template) | Out-Null } catch { } }
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
}

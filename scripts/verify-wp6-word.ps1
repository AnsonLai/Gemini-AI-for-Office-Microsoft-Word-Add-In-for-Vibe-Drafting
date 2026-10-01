param(
    [string]$ArtifactsDir = '.cache/wp6/word',
    [switch]$SkipRender,
    [switch]$SkipNativeInsert,
    [switch]$Render,
    [string]$CaseName,
    [string]$FixtureManifest, # Optional independently collected Office.js packages.
    [switch]$IncludeKnownDefects, # Compatibility flag; v0.8.2 runs these regressions by default.
    [switch]$VisibleWord,
    [int]$TimeoutSeconds = 120,
    [switch]$Worker
)

# Independent Word oracle: never save edits into an input fixture, never repair
# on open, and never call a successful PDF export a visual review.
$ErrorActionPreference = 'Stop'
$SkipRender = -not $Render
$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$outputDir = [System.IO.Path]::GetFullPath((Join-Path $repoRoot $ArtifactsDir))
$fixtureDir = Join-Path $outputDir 'fixtures'
if ($FixtureManifest) { $fixtureDir = Split-Path -Parent ([System.IO.Path]::GetFullPath((Join-Path $repoRoot $FixtureManifest))) }
$renderDir = Join-Path $outputDir 'rendered'
New-Item -ItemType Directory -Force -Path $fixtureDir, $renderDir | Out-Null
$reportPath = Join-Path $outputDir 'word-report.json'
$progressPath = Join-Path $outputDir 'word-progress.json'
if (-not $Worker) {
    if ($TimeoutSeconds -lt 1) { throw 'TimeoutSeconds must be positive.' }
    $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $PSCommandPath + '"'), '-Worker', '-ArtifactsDir', ('"' + $ArtifactsDir + '"'))
    if ($SkipRender) { $arguments += '-SkipRender' }
    if ($SkipNativeInsert) { $arguments += '-SkipNativeInsert' }
    if ($Render) { $arguments += '-Render' }
    if ($CaseName) { $arguments += @('-CaseName', ('"' + $CaseName + '"')) }
    if ($FixtureManifest) { $arguments += @('-FixtureManifest', ('"' + $FixtureManifest + '"')) }
    if ($IncludeKnownDefects) { $arguments += '-IncludeKnownDefects' }
    if ($VisibleWord) { $arguments += '-VisibleWord' }
    $started = [DateTime]::UtcNow
    $process = Start-Process powershell.exe -ArgumentList $arguments -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $outputDir 'word-stdout.log') -RedirectStandardError (Join-Path $outputDir 'word-stderr.log')
    $process.Handle | Out-Null
    if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
        $progress = if (Test-Path -LiteralPath $progressPath) { Get-Content -LiteralPath $progressPath -Raw | ConvertFrom-Json } else { $null }
        # Only stop a Word process newly created by this worker, with matching
        # creation time. Never stop an existing user Word session.
        if ($progress -and $progress.ownedWordPid -and ([DateTime]$progress.updatedAt).ToUniversalTime() -ge $started) {
            $ownedProcess = Get-Process -Id $progress.ownedWordPid -ErrorAction SilentlyContinue
            if ($ownedProcess -and $ownedProcess.StartTime.ToUniversalTime().ToString('o') -eq $progress.ownedWordStartedAt) { Stop-Process -Id $ownedProcess.Id -Force }
        }
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
        $report = if ($progress -and $progress.report) { $progress.report } else { [pscustomobject]@{ checks = @(); renders = @() } }
        $report.checks = @($report.checks) + [pscustomobject]@{ case = 'word-harness'; view = if ($progress) { $progress.stage } else { 'startup' }; status = 'failed'; error = "Timed out after $TimeoutSeconds seconds" }
        $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $reportPath -Encoding UTF8
        Get-Content -LiteralPath (Join-Path $outputDir 'word-stdout.log')
        Write-Error "Word validation timed out; partial report: $reportPath"
        exit 1
    }
    Get-Content -LiteralPath (Join-Path $outputDir 'word-stdout.log')
    Get-Content -LiteralPath (Join-Path $outputDir 'word-stderr.log') | Write-Host
    $completedReport = if (Test-Path -LiteralPath $reportPath) { Get-Content -LiteralPath $reportPath -Raw | ConvertFrom-Json } else { $null }
    if (-not $completedReport -or ([DateTime]$completedReport.checkedAt).ToUniversalTime() -lt $started -or @($completedReport.checks | Where-Object status -eq 'failed').Count) { exit 1 }
    exit $process.ExitCode
}
if ($FixtureManifest) {
    $manifest = Get-Content -LiteralPath ([System.IO.Path]::GetFullPath((Join-Path $repoRoot $FixtureManifest))) -Raw -Encoding UTF8 | ConvertFrom-Json
} else {
    & node (Join-Path $repoRoot 'tests/ooxml_formatting_visual_tests.mjs') --export-dir $fixtureDir
    if ($LASTEXITCODE -ne 0) { throw 'WP6 fixture generation failed.' }
    $manifest = Get-Content -LiteralPath (Join-Path $fixtureDir 'manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
}
$word = $null
$checks = [System.Collections.Generic.List[object]]::new()
$renders = [System.Collections.Generic.List[object]]::new()
$ownedWord = $null
function Save-Progress([string]$stage) {
    $report = [ordered]@{ checkedAt = [DateTime]::UtcNow.ToString('o'); wordVersion = if ($word) { [string]$word.Version } else { $null }; wordBuild = if ($word) { [string]$word.Build } else { $null }; checks = @($checks.ToArray()); renders = @($renders.ToArray()); renderSkipped = [bool]$SkipRender; nativeInsertSkipped = [bool]$SkipNativeInsert; officeJsTransport = 'not exercised; native Word InsertXML is a separate package parser check' }
    [ordered]@{ updatedAt = [DateTime]::UtcNow.ToString('o'); stage = $stage; ownedWordPid = if ($ownedWord) { $ownedWord.Id } else { $null }; ownedWordStartedAt = if ($ownedWord) { $ownedWord.StartTime.ToUniversalTime().ToString('o') } else { $null }; report = $report } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $progressPath -Encoding UTF8
    $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $reportPath -Encoding UTF8
}

function Resolve-FixturePath([string]$path) {
    if ([IO.Path]::IsPathRooted($path)) { return [IO.Path]::GetFullPath($path) }
    return [IO.Path]::GetFullPath((Join-Path $fixtureDir $path))
}

function Normalize-WordText([string]$text) {
    return $text.Replace("`r", "`n").Replace([string][char]11, "`n").Replace([string][char]12, "`n").TrimEnd([char[]]"`r`n")
}

function Open-WithoutRepair([string]$file) {
    Write-Host "OPEN $([System.IO.Path]::GetFileName($file))"
    Save-Progress "open $([System.IO.Path]::GetFileName($file))"
    # Keep a document window available for Word's revision and export APIs while
    # the application remains hidden. Match the upstream Word oracle.
    return $word.Documents.OpenNoRepairDialog($file)
}

function Check-Document($case, [string]$file, [string]$view, [bool]$render, [bool]$alreadyResolved = $false) {
    $document = $null
    try {
        $document = Open-WithoutRepair $file
        if ($alreadyResolved -and $document.Revisions.Count -ne 0) { throw 'Engine-resolved package still contains body revisions before Word resolution.' }
        if ($alreadyResolved -and $case.expectedHeaderFooterText) {
            foreach ($part in $case.expectedHeaderFooterText) {
                $section = $document.Sections.Item([int]$part.section)
                if ($part.kind -eq 'footer') { $partCollection = $section.Footers } else { $partCollection = $section.Headers }
                $partIndex = switch ($part.type) { 'first' { 2 }; 'even' { 3 }; default { 1 } }
                if ($partCollection.Item($partIndex).Range.Revisions.Count -ne 0) { throw 'Engine-resolved package retained header/footer revisions before Word resolution.' }
            }
        }
        if ($view -eq 'tracked' -and $case.expectedMinimumBodyRevisions -and $document.Revisions.Count -lt [int]$case.expectedMinimumBodyRevisions) { throw 'Word did not recognize the expected body revision markup.' }
        if ($view -eq 'accepted') { $document.AcceptAllRevisions() }
        if ($view -eq 'rejected') { $document.RejectAllRevisions() }
        $actual = Normalize-WordText ([string]$document.Content.Text)
        $expected = if ($view -eq 'rejected' -or $view -eq 'source') { $case.expectedRejectedText } else { $case.expectedAcceptedText }
        if ($view -ne 'tracked' -and $actual -cne $expected) {
            throw "Word $view text differs. Expected=$($expected | ConvertTo-Json -Compress); actual=$($actual | ConvertTo-Json -Compress)"
        }
        $comments = [int]$document.Comments.Count
        if ($view -ne 'source' -and $null -ne $case.expectedComments -and $comments -ne [int]$case.expectedComments) {
            throw "Expected $($case.expectedComments) comments, Word sees $comments."
        }
        if ($view -eq 'accepted' -or $view -eq 'rejected') {
            if ($document.Revisions.Count -ne 0) { throw "Word $view retained revisions." }
        }
        if ($case.expectedNumbering -and $view -ne 'tracked') {
            $numberingState = if ($view -eq 'source') { 'source' } else { $view }
            $expectedNumberingItems = @($case.expectedNumbering.$numberingState)
            $numberingGroups = @{}
            [xml]$numberingDocumentPackage = [string]$document.Content.WordOpenXML
            $numberingNamespaces = [System.Xml.XmlNamespaceManager]::new($numberingDocumentPackage.NameTable)
            $numberingNamespaces.AddNamespace('pkg', 'http://schemas.microsoft.com/office/2006/xmlPackage')
            $numberingNamespaces.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
            foreach ($expectedItem in $expectedNumberingItems) {
                $matchingParagraphs = @()
                for ($numberingIndex = 1; $numberingIndex -le $document.Paragraphs.Count; $numberingIndex++) {
                    $candidateParagraph = $document.Paragraphs.Item($numberingIndex)
                    if ((Normalize-WordText ([string]$candidateParagraph.Range.Text)) -ceq [string]$expectedItem.text) {
                        $matchingParagraphs += [pscustomobject]@{ paragraph = $candidateParagraph; index = $numberingIndex }
                    }
                }
                if ($matchingParagraphs.Count -ne 1) { throw "Word $view numbering target missing or ambiguous: $($expectedItem.text)" }
                $numberingRange = $matchingParagraphs[0].paragraph.Range
                $listFormat = $numberingRange.ListFormat
                $isList = [int]$listFormat.ListType -ne 0
                if ($isList -ne [bool]$expectedItem.isList) { throw "Word $view list/plain state differs: $($expectedItem.text)" }
                foreach ($property in @('listType', 'listLevel', 'listValue', 'listString')) {
                    if ($expectedItem.PSObject.Properties.Name -contains $property) {
                        $actualProperty = switch ($property) {
                            'listType' { [int]$listFormat.ListType }
                            'listLevel' { [int]$listFormat.ListLevelNumber }
                            'listValue' { [int]$listFormat.ListValue }
                            'listString' { [string]$listFormat.ListString }
                        }
                        if ($actualProperty -cne $expectedItem.$property) { throw "Word $view $property differs for $($expectedItem.text): expected=$($expectedItem.$property), actual=$actualProperty" }
                    }
                }
                if ($expectedItem.group) {
                    # Inspect one whole-document package so independently exported
                    # paragraph fragments cannot renumber each list to the same ID.
                    $bodyParagraphIndex = [int]$matchingParagraphs[0].index
                    $numIdNode = $numberingDocumentPackage.SelectSingleNode("//pkg:part[@pkg:name='/word/document.xml']/pkg:xmlData/w:document/w:body/w:p[$bodyParagraphIndex]/w:pPr/w:numPr/w:numId", $numberingNamespaces)
                    if (-not $numIdNode) { throw "Word $view cannot inspect list identity for $($expectedItem.text)" }
                    $actualNumId = $numIdNode.GetAttribute('val', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
                    $groupName = [string]$expectedItem.group
                    if ($numberingGroups.ContainsKey($groupName) -and $numberingGroups[$groupName] -cne $actualNumId) { throw "Word $view continuation list identity differs for group $groupName" }
                    foreach ($otherGroup in @($numberingGroups.Keys)) {
                        if ($otherGroup -cne $groupName -and $numberingGroups[$otherGroup] -ceq $actualNumId) { throw "Word $view restart groups share numbering identity: $otherGroup and $groupName" }
                    }
                    $numberingGroups[$groupName] = $actualNumId
                }
            }
        }
        if ($case.expectedSections) {
            if ($document.Sections.Count -ne @($case.expectedSections).Count) { throw 'Word section count differs.' }
            for ($sectionIndex = 1; $sectionIndex -le $document.Sections.Count; $sectionIndex++) {
                $pageSetup = $document.Sections.Item($sectionIndex).PageSetup
                $expectedSection = $case.expectedSections[$sectionIndex - 1]
                if ($pageSetup.Orientation -ne $expectedSection.orientation -or $pageSetup.TextColumns.Count -ne $expectedSection.columns) { throw "Word $view section layout differs at section $sectionIndex." }
            }
        }
        if ($case.expectedTabStop) {
            $tabs = $document.Paragraphs.Item([int]$case.expectedTabStop.paragraph).TabStops
            # Word also exposes default tab stops. Match the independently
            # specified explicit stop, rather than assuming Count is one.
            $matchingTabs = 0
            for ($tabIndex = 1; $tabIndex -le $tabs.Count; $tabIndex++) {
                $tab = $tabs.Item($tabIndex)
                if ([Math]::Abs($tab.Position - $case.expectedTabStop.positionPoints) -lt 0.01 -and $tab.Alignment -eq $case.expectedTabStop.alignment) { $matchingTabs++ }
            }
            if ($matchingTabs -ne 1) { throw "Word $view explicit tab stop differs." }
        }
        if ($view -ne 'tracked' -and $case.expectedFormatting) {
            $rawBodyText = [string]$document.Content.Text
            foreach ($format in $case.expectedFormatting) {
                $formatText = if ($view -eq 'source' -or $view -eq 'rejected') { $format.sourceText } else { $format.acceptedText }
                $start = $rawBodyText.IndexOf($formatText, [StringComparison]::Ordinal)
                if ($start -lt 0 -or $start -ne $rawBodyText.LastIndexOf($formatText, [StringComparison]::Ordinal)) { throw "Word formatting target is absent or ambiguous: $formatText" }
                $formatRange = $document.Range($document.Content.Start + $start, $document.Content.Start + $start + $formatText.Length)
                $boldFlag = if ($format.bold) { -1 } else { 0 }
                $italicFlag = if ($format.italic) { -1 } else { 0 }
                if ($formatRange.Font.Bold -ne $boldFlag -or $formatRange.Font.Italic -ne $italicFlag -or ($format.font -and $formatRange.Font.Name -cne $format.font) -or ($format.size -and $formatRange.Font.Size -ne $format.size)) { throw "Word $view formatting differs for $formatText." }
            }
        }
        if ($view -ne 'tracked' -and $case.expectedHeaderFooterText) {
            foreach ($part in $case.expectedHeaderFooterText) {
                $section = $document.Sections.Item([int]$part.section)
                # Direct assignment keeps the COM collection intact; returning
                # it through an if-expression enumerates it into a PS array.
                if ($part.kind -eq 'footer') { $collection = $section.Footers } else { $collection = $section.Headers }
                $partIndex = switch ($part.type) { 'first' { 2 }; 'even' { 3 }; default { 1 } }
                $actualPartText = Normalize-WordText ([string]$collection.Item($partIndex).Range.Text)
                if ($alreadyResolved -and $collection.Item($partIndex).Range.Revisions.Count -ne 0) { throw 'Engine-resolved package retained header/footer revisions.' }
                $expectedPartText = if ($view -eq 'rejected' -or $view -eq 'source') { $part.rejected } else { $part.accepted }
                if ($actualPartText -cne $expectedPartText) { throw "Word $view $($part.kind) text differs: $($actualPartText | ConvertTo-Json -Compress)" }
            }
        }
        if ($view -ne 'tracked' -and $case.expectedHyperlinks) {
            $actualLinks = @()
            for ($linkIndex = 1; $linkIndex -le $document.Hyperlinks.Count; $linkIndex++) {
                $link = $document.Hyperlinks.Item($linkIndex)
                $actualLinks += [pscustomobject]@{ target = [string]$link.Address; text = Normalize-WordText ([string]$link.Range.Text) }
            }
            # Empty revision-only hyperlink containers are omitted by Word.
            $actualLinks = @($actualLinks | Where-Object { $_.text.Length -gt 0 })
            if ($actualLinks.Count -ne @($case.expectedHyperlinks).Count) { throw "Word $view hyperlink count differs." }
            foreach ($expectedLink in $case.expectedHyperlinks) {
                $expectedLinkText = if ($view -eq 'source') { $expectedLink.sourceText } elseif ($view -eq 'rejected') { $expectedLink.rejectedText } else { $expectedLink.acceptedText }
                if (@($actualLinks | Where-Object { $_.target -ceq $expectedLink.target -and $_.text -ceq $expectedLinkText }).Count -ne 1) {
                    throw "Word $view hyperlink boundary differs: $($actualLinks | ConvertTo-Json -Compress)"
                }
            }
        }
        if ($view -ne 'source' -and $null -ne $case.expectedThreadIdentities) {
            $identities = @()
            for ($i = 1; $i -le $comments; $i++) {
                $comment = $document.Comments.Item($i)
                $ancestor = $comment.Ancestor
                $identities += [pscustomobject]@{ author = [string]$comment.Author; text = Normalize-WordText ([string]$comment.Range.Text); parent = if ($ancestor) { [pscustomobject]@{ author = [string]$ancestor.Author; text = Normalize-WordText ([string]$ancestor.Range.Text) } } else { $null }; done = [bool]$comment.Done }
            }
            foreach ($expectedIdentity in $case.expectedThreadIdentities) {
                $matches = @($identities | Where-Object { $_.author -ceq $expectedIdentity.author -and $_.text -ceq $expectedIdentity.text })
                if ($matches.Count -ne 1) { throw "Word comment identity missing or ambiguous: $($expectedIdentity | ConvertTo-Json -Compress)" }
                $identity = $matches[0]
                if ($identity.done -ne $expectedIdentity.done -or
                    (($null -eq $identity.parent) -ne ($null -eq $expectedIdentity.parent)) -or
                    ($identity.parent -and ($identity.parent.author -cne $expectedIdentity.parent.author -or $identity.parent.text -cne $expectedIdentity.parent.text))) {
                    throw "Word comment thread state differs: $($identity | ConvertTo-Json -Compress -Depth 4)"
                }
            }
        }
        if ($render -and -not $SkipRender) {
            Write-Host "EXPORT $($case.name) $view"
            $pdf = Join-Path $renderDir "$($case.name)-$view.pdf"
            Save-Progress "ShowRevisions $($case.name) $view"
            $document.ShowRevisions = $view -eq 'tracked'
            Save-Progress "PrintRevisions $($case.name) $view"
            $document.PrintRevisions = $view -eq 'tracked'
            $exportItem = if ($view -eq 'tracked') { 7 } else { 0 }
            Save-Progress "ExportAsFixedFormat $($case.name) $view"
            $document.ExportAsFixedFormat($pdf, 17, $false, 0, 0, 1, 1, $exportItem, $true, $true, 0, $true, $true, $false)
            $bytes = (Get-Item -LiteralPath $pdf).Length
            if ($bytes -lt 1000) { throw "Word produced a suspiciously small PDF: $bytes bytes." }
            $renders.Add([pscustomobject]@{ case = $case.name; view = $view; pdf = $pdf; pages = [int]$document.ComputeStatistics(2); status = 'rendered'; visualReview = 'pending' })
        }
        $checks.Add([pscustomobject]@{ case = $case.name; view = $view; status = 'passed'; comments = $comments; revisions = [int]$document.Revisions.Count })
        Write-Output "PASS $($case.name) $view (comments=$comments)"
    }
    catch {
        $checks.Add([pscustomobject]@{ case = $case.name; view = $view; status = 'failed'; error = $_.Exception.Message })
        Write-Output "FAIL $($case.name) $view : $($_.Exception.Message)"
    }
    finally { if ($document) { $document.Close(0) | Out-Null; [Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null } }
}

try {
    $existingWordPids = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
    $word = New-Object -ComObject Word.Application
    $createdWord = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $_.Id -notin $existingWordPids })
    if ($createdWord.Count -eq 1) { $ownedWord = $createdWord[0] }
    $word.Visible = [bool]$VisibleWord
    $word.DisplayAlerts = 0
    $selectedCases = @($manifest.cases | Where-Object { -not $CaseName -or $_.name -eq $CaseName })
    if (-not $selectedCases.Count) { throw 'No Word fixtures matched the requested lane.' }
    foreach ($case in $selectedCases) {
        Check-Document $case (Resolve-FixturePath $case.source) 'source' $false
        foreach ($view in @('tracked', 'accepted', 'rejected')) {
            Check-Document $case (Resolve-FixturePath $case.tracked) $view $true
        }
        # Also open the engine-resolved package independently of Word's own resolution.
        Check-Document $case (Resolve-FixturePath $case.accepted) 'accepted' $false $true
        Check-Document $case (Resolve-FixturePath $case.rejected) 'rejected' $false $true
        if ($case.insertionXml -and -not $SkipNativeInsert) {
            $document = $null
            $nativePath = $null
            try {
                $document = Open-WithoutRepair (Resolve-FixturePath $case.source)
                $document.TrackRevisions = $false
                if ($case.nativeOperations) {
                    $liveInput = Join-Path $fixtureDir "$($case.name)-live-input.xml"
                    $liveOperations = Join-Path $fixtureDir "$($case.name)-live-operations.json"
                    $liveOutput = Join-Path $fixtureDir "$($case.name)-live-output.xml"
                    Save-Progress "WordOpenXML $($case.name)"
                    [string]$document.Content.WordOpenXML | Set-Content -LiteralPath $liveInput -Encoding UTF8
                    ConvertTo-Json -InputObject @($case.nativeOperations) -Depth 15 | Set-Content -LiteralPath $liveOperations -Encoding UTF8
                    & node (Join-Path $repoRoot 'scripts/apply-live-word-ooxml.mjs') --input $liveInput --operations $liveOperations --output $liveOutput
                    if ($LASTEXITCODE -ne 0) { throw 'Production bridge failed against live Word scope XML.' }
                    $payload = Get-Content -LiteralPath $liveOutput -Raw -Encoding UTF8
                } else {
                    $payload = Get-Content -LiteralPath (Join-Path $fixtureDir $case.insertionXml) -Raw -Encoding UTF8
                }
                Save-Progress "InsertXML $($case.name)"
                # Explicitly marshal the optional Transform parameter.
                $insertionRange = $document.Content
                $insertionRange.InsertXML([string]$payload, [Type]::Missing)
                $nativePath = [string](Join-Path $fixtureDir "$($case.name)-native-insert.docx")
                Save-Progress "SaveAs2 $($case.name)"
                [object]$nativeSavePath = $nativePath
                [object]$nativeFormat = 12
                $document.SaveAs2([ref]$nativeSavePath, [ref]$nativeFormat)
                $checks.Add([pscustomobject]@{ case = $case.name; view = 'native-insert'; status = 'passed' })
            }
            catch {
                $checks.Add([pscustomobject]@{ case = $case.name; view = 'native-insert'; status = 'failed'; error = $_.Exception.Message })
                Write-Output "FAIL $($case.name) native-insert : $($_.Exception.Message)"
            }
            finally { if ($document) { $document.Close(0) | Out-Null; [Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null } }
            if ($nativePath -and (Test-Path -LiteralPath $nativePath)) {
                Check-Document $case $nativePath 'tracked' $false
                Check-Document $case $nativePath 'accepted' $false
                Check-Document $case $nativePath 'rejected' $false
            }
        }
    }
    foreach ($control in @($manifest.controls | Where-Object { $null -ne $_ })) {
        $document = $null
        $opened = $false
        try { $document = Open-WithoutRepair (Resolve-FixturePath $control.file); $opened = $true } catch { }
        finally { if ($document) { $document.Close(0) | Out-Null; [Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null } }
        $status = if ($opened -eq [bool]$control.expectOpen) { 'passed' } else { 'failed' }
        $checks.Add([pscustomobject]@{ case = $control.file; view = 'negative-control'; status = $status; opened = $opened })
        Write-Output "$($status.ToUpper()) $($control.file) (opened=$opened)"
    }
    $transportEvidence = if ($FixtureManifest) { 'External collected packages verified; consult the collector report for transport provenance.' } else { 'not exercised; native Word InsertXML is a separate package parser check' }
    $report = [ordered]@{ checkedAt = [DateTime]::UtcNow.ToString('o'); wordVersion = [string]$word.Version; wordBuild = [string]$word.Build; checks = @($checks.ToArray()); renders = @($renders.ToArray()); renderSkipped = [bool]$SkipRender; nativeInsertSkipped = [bool]$SkipNativeInsert; officeJsTransport = $transportEvidence; fixtureManifest = $FixtureManifest }
    $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $outputDir 'word-report.json') -Encoding UTF8
}
finally { if ($word) { $word.Quit() | Out-Null; [Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null } }
if (@($checks | Where-Object status -eq 'failed').Count) { exit 1 }
Write-Output "Requested Word checks passed. Render skipped=$SkipRender; native InsertXML skipped=$SkipNativeInsert. Any rendered PDFs require visual inspection. Report: $(Join-Path $outputDir 'word-report.json')"

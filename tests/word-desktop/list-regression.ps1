param()

$ErrorActionPreference = 'Stop'

function Normalize-Text {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return (($Text -replace '\s+', ' ').Trim()).ToLowerInvariant()
}

function Find-RowBySnippet {
    param(
        [array]$Rows,
        [string]$Snippet
    )
    $target = Normalize-Text -Text $Snippet
    foreach ($row in $Rows) {
        if ((Normalize-Text -Text $row.text).Contains($target)) {
            return $row
        }
    }
    throw "Word inspector row not found for snippet: `"$Snippet`""
}

function Assert-ListMarker {
    param(
        [array]$Rows,
        [string]$Snippet,
        [string]$ExpectedPrefix
    )
    $row = Find-RowBySnippet -Rows $Rows -Snippet $Snippet
    if (-not $row.isList) {
        throw "Expected list paragraph for `"$Snippet`""
    }
    $marker = [string]$row.listString
    if (-not $marker.Trim().StartsWith($ExpectedPrefix)) {
        throw "Expected marker `"$ExpectedPrefix`" for `"$Snippet`", got `"$marker`""
    }
}

function Assert-NonEmptyMarker {
    param(
        [array]$Rows,
        [string]$Snippet
    )
    $row = Find-RowBySnippet -Rows $Rows -Snippet $Snippet
    $marker = [string]$row.listString
    if ([string]::IsNullOrWhiteSpace($marker)) {
        throw "Expected non-empty list marker for `"$Snippet`""
    }
}

function Invoke-InspectorWithRetry {
    param(
        [string]$InspectorScript,
        [string]$DocxPath,
        [int]$MaxAttempts = 3
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            return (& $InspectorScript -DocxPath $DocxPath -OnlyLists | ConvertFrom-Json)
        } catch {
            if ($attempt -ge $MaxAttempts) { throw }
            Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
            Start-Sleep -Seconds 2
        }
    }
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Resolve-Path (Join-Path $scriptRoot '..\..')
$nodeScript = Join-Path $scriptRoot 'list-regression.mjs'
$manifestPath = Join-Path $scriptRoot '.tmp\list-regression-paths.json'
$inspectorScript = Join-Path $scriptRoot 'list-inspector.ps1'

Push-Location $projectRoot
try {
    node $nodeScript | Out-Host
} finally {
    Pop-Location
}

if (-not (Test-Path $manifestPath)) {
    throw "Missing regression manifest: $manifestPath"
}

$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
$workFolder = [string]$manifest.workFolder
$outputDocx = [string]$manifest.outputDocx
$outputInspectorJson = [string]$manifest.outputInspectorJson
$zipPath = [System.IO.Path]::ChangeExtension($outputDocx, '.zip')

if (Test-Path $outputDocx) {
    Remove-Item $outputDocx -Force
}
if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}

Push-Location $workFolder
try {
    Compress-Archive -Path * -DestinationPath $zipPath
} finally {
    Pop-Location
}
Move-Item -LiteralPath $zipPath -Destination $outputDocx -Force

$rows = Invoke-InspectorWithRetry -InspectorScript $inspectorScript -DocxPath $outputDocx
$rows | ConvertTo-Json -Depth 6 | Set-Content -Path $outputInspectorJson -Encoding UTF8

Assert-ListMarker -Rows $rows -Snippet 'DEFINITION OF CONFIDENTIAL INFORMATION' -ExpectedPrefix '1.'
Assert-ListMarker -Rows $rows -Snippet 'EXCLUSIONS' -ExpectedPrefix '2.'
Assert-ListMarker -Rows $rows -Snippet 'OBLIGATIONS OF RECEIVING PARTY' -ExpectedPrefix '3.'
Assert-ListMarker -Rows $rows -Snippet '4. TERM' -ExpectedPrefix '4.'
Assert-ListMarker -Rows $rows -Snippet 'REQUIRED DISCLOSURE' -ExpectedPrefix '5.'
Assert-ListMarker -Rows $rows -Snippet 'RETURN OF INFORMATION' -ExpectedPrefix '6.'
Assert-ListMarker -Rows $rows -Snippet 'REMEDIES' -ExpectedPrefix '7.'
Assert-ListMarker -Rows $rows -Snippet 'GOVERNING LAW' -ExpectedPrefix '8.'
Assert-ListMarker -Rows $rows -Snippet 'GENERAL PROVISIONS' -ExpectedPrefix '9.'

Assert-ListMarker -Rows $rows -Snippet 'is or becomes generally available to the public' -ExpectedPrefix '1.'
Assert-ListMarker -Rows $rows -Snippet 'was already known to the Receiving Party prior to disclosure' -ExpectedPrefix '2.'
Assert-ListMarker -Rows $rows -Snippet 'becomes available to the Receiving Party on a non-confidential basis' -ExpectedPrefix '3.'
Assert-ListMarker -Rows $rows -Snippet 'is independently developed by the Receiving Party' -ExpectedPrefix '4.'

Assert-ListMarker -Rows $rows -Snippet 'Use the Confidential Information solely for the Purpose' -ExpectedPrefix '1.'
Assert-ListMarker -Rows $rows -Snippet 'Protect the Confidential Information from unauthorized use' -ExpectedPrefix '2.'
Assert-ListMarker -Rows $rows -Snippet 'Disclose the Confidential Information only to its employees' -ExpectedPrefix '3.'
Assert-ListMarker -Rows $rows -Snippet 'Notify the Disclosing Party immediately upon discovery' -ExpectedPrefix '4.'

Assert-ListMarker -Rows $rows -Snippet 'Specifically, such retention must be legally required by the SEC or FCC.' -ExpectedPrefix '2.2.1'
Assert-NonEmptyMarker -Rows $rows -Snippet 'Handling of Confidential Materials'
Assert-NonEmptyMarker -Rows $rows -Snippet 'Action: Promptly return or destroy all documents'
Assert-NonEmptyMarker -Rows $rows -Snippet 'Scope: This includes the original documents and all copies thereof.'
Assert-NonEmptyMarker -Rows $rows -Snippet 'Archival Exception'
Assert-NonEmptyMarker -Rows $rows -Snippet 'legal counsel may retain one (1) copy of the materials.'
Assert-NonEmptyMarker -Rows $rows -Snippet 'This copy is to be used solely for archival purposes'
Assert-NonEmptyMarker -Rows $rows -Snippet 'Specifically, such retention must be legally required by the SEC or FCC.'
Assert-NonEmptyMarker -Rows $rows -Snippet 'Verification of Compliance'
Assert-NonEmptyMarker -Rows $rows -Snippet 'If requested, the Receiving Party must provide a written certification'

Write-Host "PASS: Word list regression checks passed."
Write-Host "Output docx: $outputDocx"
Write-Host "Inspector JSON: $outputInspectorJson"

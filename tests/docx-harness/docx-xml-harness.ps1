param(
    [ValidateSet('summary', 'list', 'extract', 'show', 'grep', 'query')]
    [string]$Action = 'summary',

    [string]$InputPath,

    [string]$Part,

    [string]$Pattern,

    [string]$XPath,

    [string]$OutDir,

    [switch]$Regex,

    [switch]$CaseSensitive,

    [int]$MaxResults = 0,

    [switch]$AsJson,

    [switch]$KeepExtracted
)

$ErrorActionPreference = 'Stop'

$modulePath = Join-Path $PSScriptRoot 'docx-xml-harness.psm1'
Import-Module $modulePath -Force

if ([string]::IsNullOrWhiteSpace($InputPath)) {
    $defaultDocx = Join-Path (Resolve-Path (Join-Path $PSScriptRoot '..')).Path 'Sample NDA.docx'
    if (Test-Path -LiteralPath $defaultDocx -PathType Leaf) {
        $InputPath = $defaultDocx
    } else {
        throw 'InputPath is required when tests/Sample NDA.docx is not present.'
    }
}

$package = Open-DocxXmlPackage -InputPath $InputPath

try {
    switch ($Action) {
        'summary' {
            $summary = Get-DocxPackageSummary -PackageRoot $package.PackageRoot
            if ($AsJson) {
                $summary | ConvertTo-Json -Depth 8
            } else {
                Write-Output "Source: $($package.SourcePath)"
                Write-Output "PackageRoot: $($summary.packageRoot)"
                Write-Output "Parts: $($summary.partCount) total, $($summary.xmlPartCount) xml/rels"
                Write-Output "Has word/document.xml: $($summary.hasDocumentXml)"
                Write-Output "Has word/numbering.xml: $($summary.hasNumberingXml)"
                if ($summary.document) {
                    Write-Output "Document: paragraphs=$($summary.document.paragraphCount), listParagraphs=$($summary.document.listParagraphCount), uniqueNumIds=$([string]::Join(',', $summary.document.uniqueNumIds))"
                }
                if ($summary.numbering) {
                    Write-Output "Numbering: abstractNum=$($summary.numbering.abstractNumCount), num=$($summary.numbering.numCount), schemaOrderValid=$($summary.numbering.schemaOrderValid)"
                }
            }
        }

        'list' {
            $parts = Get-DocxPackageParts -PackageRoot $package.PackageRoot -XmlOnly
            if ($AsJson) {
                $parts | ConvertTo-Json -Depth 6
            } else {
                $parts | ForEach-Object { Write-Output $_.part }
            }
        }

        'extract' {
            if ([string]::IsNullOrWhiteSpace($OutDir)) {
                $OutDir = Join-Path $PSScriptRoot ".tmp\extracted-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
            }
            if (Test-Path -LiteralPath $OutDir) {
                Remove-Item -LiteralPath $OutDir -Recurse -Force
            }
            New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
            Copy-Item -LiteralPath (Join-Path $package.PackageRoot '*') -Destination $OutDir -Recurse -Force
            if ($AsJson) {
                [pscustomobject]@{
                    extractedTo = (Resolve-Path -LiteralPath $OutDir).Path
                    source = $package.SourcePath
                } | ConvertTo-Json -Depth 4
            } else {
                Write-Output "Extracted to: $OutDir"
            }
        }

        'show' {
            if ([string]::IsNullOrWhiteSpace($Part)) {
                throw 'Part is required for Action=show (for example: word/document.xml).'
            }
            $text = Get-DocxPartText -PackageRoot $package.PackageRoot -Part $Part
            if ($AsJson) {
                [pscustomobject]@{
                    part = $Part
                    text = $text
                } | ConvertTo-Json -Depth 5
            } else {
                Write-Output $text
            }
        }

        'grep' {
            if ([string]::IsNullOrWhiteSpace($Pattern)) {
                throw 'Pattern is required for Action=grep.'
            }
            $matches = Find-DocxXmlPattern -PackageRoot $package.PackageRoot -Pattern $Pattern -Regex:$Regex -CaseSensitive:$CaseSensitive -Part $Part -MaxResults $MaxResults
            if ($AsJson) {
                $matches | ConvertTo-Json -Depth 8
            } else {
                foreach ($m in $matches) {
                    Write-Output "$($m.part):$($m.lineNumber):$($m.column) $($m.match)"
                }
                if ($matches.Count -eq 0) {
                    Write-Output 'No matches.'
                }
            }
        }

        'query' {
            if ([string]::IsNullOrWhiteSpace($Part)) {
                throw 'Part is required for Action=query.'
            }
            if ([string]::IsNullOrWhiteSpace($XPath)) {
                throw 'XPath is required for Action=query.'
            }
            $results = Invoke-DocxXPathQuery -PackageRoot $package.PackageRoot -Part $Part -XPath $XPath -MaxResults $MaxResults
            if ($AsJson) {
                $results | ConvertTo-Json -Depth 10
            } else {
                foreach ($r in $results) {
                    Write-Output "[$($r.index)] $($r.name): $($r.innerText)"
                }
                if ($results.Count -eq 0) {
                    Write-Output 'No nodes matched.'
                }
            }
        }
    }
}
finally {
    Close-DocxXmlPackage -Package $package -KeepExtracted:$KeepExtracted
}

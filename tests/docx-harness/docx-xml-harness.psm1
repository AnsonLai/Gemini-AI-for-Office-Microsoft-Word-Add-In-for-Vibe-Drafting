Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Resolve-InputAbsolutePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $resolved = Resolve-Path -LiteralPath $Path -ErrorAction SilentlyContinue
    if ($null -eq $resolved) {
        throw "Path not found: $Path"
    }
    return $resolved.Path
}

function Convert-ToPackageRelativePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackageRoot,
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    $root = (Resolve-InputAbsolutePath -Path $PackageRoot).TrimEnd('\') + '\'
    $file = Resolve-InputAbsolutePath -Path $FilePath
    $rootUri = [System.Uri]::new($root)
    $fileUri = [System.Uri]::new($file)
    $relative = [System.Uri]::UnescapeDataString($rootUri.MakeRelativeUri($fileUri).ToString())
    return $relative -replace '\\', '/'
}

function Resolve-DocxPartPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackageRoot,
        [Parameter(Mandatory = $true)]
        [string]$Part
    )

    $normalized = $Part.Trim() -replace '^[/\\]+', '' -replace '/', '\'
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        throw 'Part cannot be empty.'
    }

    $fullPath = Join-Path $PackageRoot $normalized
    if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
        throw "Part not found in package: $Part"
    }
    return $fullPath
}

function New-HarnessTempFolder {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root
    )

    if (-not (Test-Path -LiteralPath $Root)) {
        New-Item -ItemType Directory -Path $Root -Force | Out-Null
    }

    $folderName = "docx-" + (Get-Date -Format 'yyyyMMdd-HHmmss-fff') + "-" + ([guid]::NewGuid().ToString('N').Substring(0, 8))
    $path = Join-Path $Root $folderName
    New-Item -ItemType Directory -Path $path -Force | Out-Null
    return $path
}

function Test-IsDocxPackageFolder {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $contentTypes = Join-Path $Path '[Content_Types].xml'
    $wordDocument = Join-Path $Path 'word\document.xml'
    return (Test-Path -LiteralPath $contentTypes -PathType Leaf) -and (Test-Path -LiteralPath $wordDocument -PathType Leaf)
}

function Open-DocxXmlPackage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,
        [string]$WorkRoot = (Join-Path $PSScriptRoot '.tmp')
    )

    $resolved = Resolve-InputAbsolutePath -Path $InputPath
    $item = Get-Item -LiteralPath $resolved

    if ($item.PSIsContainer) {
        if (-not (Test-IsDocxPackageFolder -Path $resolved)) {
            throw "Input folder is not a valid unzipped .docx package: $resolved"
        }
        return [pscustomobject]@{
            SourcePath = $resolved
            SourceType = 'folder'
            PackageRoot = $resolved
            IsTemporary = $false
            WorkRoot = $null
        }
    }

    $extension = [System.IO.Path]::GetExtension($resolved)
    if ($extension -ne '.docx') {
        throw "Input file must be .docx or unzipped package folder. Got: $resolved"
    }

    $tempFolder = New-HarnessTempFolder -Root $WorkRoot
    $tempZipPath = Join-Path $tempFolder '__source.zip'
    Copy-Item -LiteralPath $resolved -Destination $tempZipPath -Force
    Expand-Archive -LiteralPath $tempZipPath -DestinationPath $tempFolder -Force
    Remove-Item -LiteralPath $tempZipPath -Force -ErrorAction SilentlyContinue

    if (-not (Test-IsDocxPackageFolder -Path $tempFolder)) {
        throw "Expanded archive is missing required OOXML parts: $resolved"
    }

    return [pscustomobject]@{
        SourcePath = $resolved
        SourceType = 'docx'
        PackageRoot = $tempFolder
        IsTemporary = $true
        WorkRoot = $WorkRoot
    }
}

function Close-DocxXmlPackage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Package,
        [switch]$KeepExtracted
    )

    if ($KeepExtracted) { return }
    if (-not $Package.IsTemporary) { return }
    if ([string]::IsNullOrWhiteSpace($Package.PackageRoot)) { return }
    if (Test-Path -LiteralPath $Package.PackageRoot) {
        Remove-Item -LiteralPath $Package.PackageRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Get-DocxPackageParts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackageRoot,
        [switch]$XmlOnly
    )

    $root = Resolve-InputAbsolutePath -Path $PackageRoot
    $allFiles = Get-ChildItem -LiteralPath $root -File -Recurse
    if ($XmlOnly) {
        $allFiles = $allFiles | Where-Object { $_.Extension -ieq '.xml' -or $_.Name -ieq '.rels' }
    }

    return $allFiles | ForEach-Object {
        [pscustomobject]@{
            part = (Convert-ToPackageRelativePath -PackageRoot $root -FilePath $_.FullName)
            fullPath = $_.FullName
            size = $_.Length
        }
    } | Sort-Object part
}

function Get-DocxPartText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackageRoot,
        [Parameter(Mandatory = $true)]
        [string]$Part
    )

    $partPath = Resolve-DocxPartPath -PackageRoot $PackageRoot -Part $Part
    return Get-Content -LiteralPath $partPath -Raw
}

function New-XmlNamespaceManagerDefault {
    param(
        [Parameter(Mandatory = $true)]
        [xml]$XmlDoc
    )

    $nsmgr = [System.Xml.XmlNamespaceManager]::new($XmlDoc.NameTable)
    $nsmgr.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    $nsmgr.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $nsmgr.AddNamespace('pkg', 'http://schemas.microsoft.com/office/2006/xmlPackage')
    return $nsmgr
}

function Get-DefaultNamespaces {
    return @{
        w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        pkg = 'http://schemas.microsoft.com/office/2006/xmlPackage'
    }
}

function Select-XmlNodesCompat {
    param(
        [Parameter(Mandatory = $true)]
        [xml]$XmlDoc,
        [Parameter(Mandatory = $true)]
        [string]$XPath
    )

    $matches = Select-Xml -Xml $XmlDoc -XPath $XPath -Namespace (Get-DefaultNamespaces)
    if ($null -eq $matches) { return @() }
    return @($matches | ForEach-Object { $_.Node })
}

function Get-DocxPackageSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackageRoot
    )

    $parts = Get-DocxPackageParts -PackageRoot $PackageRoot
    $xmlParts = Get-DocxPackageParts -PackageRoot $PackageRoot -XmlOnly
    $summary = [ordered]@{
        packageRoot = (Resolve-InputAbsolutePath -Path $PackageRoot)
        partCount = $parts.Count
        xmlPartCount = $xmlParts.Count
        hasDocumentXml = $false
        hasNumberingXml = $false
        document = $null
        numbering = $null
    }

    $documentPart = $parts | Where-Object { $_.part -ieq 'word/document.xml' } | Select-Object -First 1
    if ($null -ne $documentPart) {
        $summary.hasDocumentXml = $true
        $docXml = [xml](Get-Content -LiteralPath $documentPart.fullPath -Raw)
        $paragraphNodes = Select-XmlNodesCompat -XmlDoc $docXml -XPath '//w:p'
        $listParagraphNodes = Select-XmlNodesCompat -XmlDoc $docXml -XPath '//w:p[w:pPr/w:numPr]'
        $numIdNodes = Select-XmlNodesCompat -XmlDoc $docXml -XPath '//w:pPr/w:numPr/w:numId'
        $numIds = @()
        foreach ($numIdNode in $numIdNodes) {
            $raw = $numIdNode.GetAttribute('w:val')
            if ([string]::IsNullOrWhiteSpace($raw)) { $raw = $numIdNode.GetAttribute('val') }
            if (-not [string]::IsNullOrWhiteSpace($raw)) { $numIds += $raw }
        }
        $uniqueNumIds = $numIds | Sort-Object -Unique
        $summary.document = [ordered]@{
            paragraphCount = ($paragraphNodes | Measure-Object).Count
            listParagraphCount = ($listParagraphNodes | Measure-Object).Count
            numIdReferenceCount = $numIds.Count
            uniqueNumIds = $uniqueNumIds
        }
    }

    $numberingPart = $parts | Where-Object { $_.part -ieq 'word/numbering.xml' } | Select-Object -First 1
    if ($null -ne $numberingPart) {
        $summary.hasNumberingXml = $true
        $numXml = [xml](Get-Content -LiteralPath $numberingPart.fullPath -Raw)
        $abstractNumNodes = Select-XmlNodesCompat -XmlDoc $numXml -XPath '//w:abstractNum'
        $numNodes = Select-XmlNodesCompat -XmlDoc $numXml -XPath '//w:num'

        $orderValid = $true
        $firstInvalidChild = $null
        $sawNum = $false
        foreach ($child in $numXml.DocumentElement.ChildNodes) {
            if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
            if ($child.LocalName -eq 'num') {
                $sawNum = $true
                continue
            }
            if ($child.LocalName -eq 'abstractNum' -and $sawNum) {
                $orderValid = $false
                $firstInvalidChild = $child.LocalName
                break
            }
        }

        $summary.numbering = [ordered]@{
            abstractNumCount = ($abstractNumNodes | Measure-Object).Count
            numCount = ($numNodes | Measure-Object).Count
            schemaOrderValid = $orderValid
            firstInvalidChild = $firstInvalidChild
        }
    }

    return [pscustomobject]$summary
}

function Find-DocxXmlPattern {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackageRoot,
        [Parameter(Mandatory = $true)]
        [string]$Pattern,
        [switch]$Regex,
        [switch]$CaseSensitive,
        [string]$Part,
        [int]$MaxResults = 0
    )

    $parts = Get-DocxPackageParts -PackageRoot $PackageRoot -XmlOnly
    if (-not [string]::IsNullOrWhiteSpace($Part)) {
        $normalizedPart = $Part.Trim() -replace '^[/\\]+', '' -replace '\\', '/'
        $parts = $parts | Where-Object { $_.part -like $normalizedPart }
    }

    $regexOptions = [System.Text.RegularExpressions.RegexOptions]::Multiline
    if (-not $CaseSensitive) {
        $regexOptions = $regexOptions -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    }

    $rx = if ($Regex) {
        [regex]::new($Pattern, $regexOptions)
    } else {
        [regex]::new([regex]::Escape($Pattern), $regexOptions)
    }

    $results = New-Object System.Collections.Generic.List[object]
    foreach ($partEntry in $parts) {
        $lines = @(Get-Content -LiteralPath $partEntry.fullPath)
        for ($i = 0; $i -lt $lines.Length; $i++) {
            $line = [string]$lines[$i]
            $matches = $rx.Matches($line)
            if ($matches.Count -eq 0) { continue }
            foreach ($m in $matches) {
                $results.Add([pscustomobject]@{
                    part = $partEntry.part
                    lineNumber = $i + 1
                    column = $m.Index + 1
                    match = $m.Value
                    lineText = $line
                }) | Out-Null

                if ($MaxResults -gt 0 -and $results.Count -ge $MaxResults) {
                    return $results
                }
            }
        }
    }

    return $results
}

function Invoke-DocxXPathQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackageRoot,
        [Parameter(Mandatory = $true)]
        [string]$Part,
        [Parameter(Mandatory = $true)]
        [string]$XPath,
        [int]$MaxResults = 0
    )

    $partPath = Resolve-DocxPartPath -PackageRoot $PackageRoot -Part $Part
    $xmlDoc = [xml](Get-Content -LiteralPath $partPath -Raw)
    $nodes = Select-XmlNodesCompat -XmlDoc $xmlDoc -XPath $XPath

    $results = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $nodes.Length; $i++) {
        $node = $nodes[$i]
        $results.Add([pscustomobject]@{
            index = $i + 1
            name = $node.Name
            localName = $node.LocalName
            innerText = $node.InnerText
            outerXml = $node.OuterXml
        }) | Out-Null

        if ($MaxResults -gt 0 -and $results.Count -ge $MaxResults) {
            break
        }
    }

    return $results
}

Export-ModuleMember -Function `
    Open-DocxXmlPackage, `
    Close-DocxXmlPackage, `
    Get-DocxPackageParts, `
    Get-DocxPartText, `
    Get-DocxPackageSummary, `
    Find-DocxXmlPattern, `
    Invoke-DocxXPathQuery

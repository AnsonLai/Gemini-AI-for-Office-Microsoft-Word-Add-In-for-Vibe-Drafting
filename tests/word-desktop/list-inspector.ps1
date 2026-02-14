param(
    [Parameter(Mandatory = $true)]
    [string]$DocxPath,

    [switch]$OnlyLists,

    [int]$MaxRows = 0
)

$ErrorActionPreference = 'Stop'

function Normalize-ParagraphText {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    $clean = $Text -replace '[\r\a]+$', ''
    return $clean
}

$resolvedPath = Resolve-Path -LiteralPath $DocxPath
$fullPath = $resolvedPath.Path

$word = $null
$document = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $document = $word.Documents.Open(
        $fullPath,
        [ref]$false,  # ConfirmConversions
        [ref]$true    # ReadOnly
    )

    $rows = New-Object System.Collections.Generic.List[object]
    $wdListNoNumbering = 0

    for ($index = 1; $index -le $document.Paragraphs.Count; $index++) {
        $paragraph = $document.Paragraphs.Item($index)
        $text = Normalize-ParagraphText -Text $paragraph.Range.Text
        if ([string]::IsNullOrWhiteSpace($text)) { continue }

        $listFormat = $paragraph.Range.ListFormat
        $listType = [int]$listFormat.ListType
        $isList = $listType -ne $wdListNoNumbering

        if ($OnlyLists -and -not $isList) { continue }

        $listString = $null
        $listLevel = $null
        $listValue = $null
        $listId = $null
        $singleList = $null
        $singleListTemplate = $null

        if ($isList) {
            try { $listString = [string]$listFormat.ListString } catch {}
            try { $listLevel = [int]$listFormat.ListLevelNumber } catch {}
            try { $listValue = [int]$listFormat.ListValue } catch {}
            try { $singleList = [bool]$listFormat.SingleList } catch {}
            try { $singleListTemplate = [bool]$listFormat.SingleListTemplate } catch {}
            try {
                $listObj = $listFormat.List
                if ($null -ne $listObj) {
                    $listId = [int]$listObj.ID
                }
            } catch {}
        }

        $styleName = $null
        try { $styleName = [string]$paragraph.Range.Style } catch {}

        $rows.Add([pscustomobject]@{
            paragraphIndex = $index
            isList = $isList
            listType = $listType
            listId = $listId
            listString = $listString
            listValue = $listValue
            listLevel = $listLevel
            singleList = $singleList
            singleListTemplate = $singleListTemplate
            style = $styleName
            text = $text
        }) | Out-Null

        if ($MaxRows -gt 0 -and $rows.Count -ge $MaxRows) { break }
    }

    $rows | ConvertTo-Json -Depth 6
}
finally {
    if ($null -ne $document) {
        try { $document.Close([ref]$false) } catch {}
    }
    if ($null -ne $word) {
        try { $word.Quit() } catch {}
    }

    if ($null -ne $document) {
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) } catch {}
    }
    if ($null -ne $word) {
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) } catch {}
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

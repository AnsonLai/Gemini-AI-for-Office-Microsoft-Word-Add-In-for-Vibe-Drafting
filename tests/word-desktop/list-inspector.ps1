param(
    [Parameter(Mandatory = $true)]
    [string]$DocxPath,

    [switch]$OnlyLists,

    [int]$MaxRows = 0,

    [int]$RetryCount = 3,

    [switch]$KillWordBeforeStart,

    [bool]$AttemptWordRepair = $true,

    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

function Normalize-ParagraphText {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    $clean = $Text -replace '[\r\a]+$', ''
    return $clean
}

function Stop-WordProcesses {
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

function Ensure-WordWorkfileEnvironment {
    $localAppData = $env:LOCALAPPDATA
    if ([string]::IsNullOrWhiteSpace($localAppData)) {
        if (-not [string]::IsNullOrWhiteSpace($env:USERPROFILE)) {
            $localAppData = Join-Path $env:USERPROFILE 'AppData\Local'
        }
    }
    if ([string]::IsNullOrWhiteSpace($localAppData)) {
        throw "LOCALAPPDATA/USERPROFILE are unavailable; cannot prepare Word workfile environment."
    }

    $tempRoot = Join-Path $localAppData 'Temp'
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    $env:TEMP = $tempRoot
    $env:TMP = $tempRoot

    $wordWorkFolders = @(
        (Join-Path $localAppData 'Microsoft\Windows\INetCache\Content.Word'),
        (Join-Path $localAppData 'Microsoft\Windows\INetCache\Content.MSO'),
        (Join-Path $localAppData 'Microsoft\Windows\Temporary Internet Files\Content.Word')
    )
    foreach ($folder in $wordWorkFolders) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
}

function Release-ComObject {
    param([object]$ComObject)
    if ($null -eq $ComObject) { return }
    try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject) } catch {}
}

function Invoke-WordRegistrationRepair {
    Ensure-WordWorkfileEnvironment
    Stop-WordProcesses
    Start-Sleep -Milliseconds 500

    try {
        $wordExe = (Get-Command WINWORD.EXE -ErrorAction SilentlyContinue).Source
        if (-not [string]::IsNullOrWhiteSpace($wordExe)) {
            $proc = Start-Process -FilePath $wordExe -ArgumentList '/r','/q' -PassThru
            $null = $proc.WaitForExit(30000)
            if (-not $proc.HasExited) {
                try { $proc.Kill() } catch {}
            }
        }
    } catch {}

    Stop-WordProcesses
    Start-Sleep -Milliseconds 800
}

function Invoke-ListInspectionPass {
    param(
        [string]$Path,
        [switch]$OnlyListRows,
        [int]$RowLimit
    )

    $word = $null
    $document = $null

    try {
        $word = New-Object -ComObject Word.Application
        $null = $word.Version
        $word.Visible = $false
        $word.DisplayAlerts = 0
        try { $word.ScreenUpdating = $false } catch {}

        $openError = $null
        try {
            $document = $word.Documents.Open(
                $Path,
                [ref]$false,  # ConfirmConversions
                [ref]$true,   # ReadOnly
                [ref]$false,  # AddToRecentFiles
                [ref]'',      # PasswordDocument
                [ref]'',      # PasswordTemplate
                [ref]$false,  # Revert
                [ref]'',      # WritePasswordDocument
                [ref]'',      # WritePasswordTemplate
                [ref]0,       # Format
                [ref]0,       # Encoding
                [ref]$false,  # Visible
                [ref]$true,   # OpenAndRepair
                [ref]0,       # DocumentDirection
                [ref]$true    # NoEncodingDialog
            )
        } catch {
            $openError = $_
        }

        if ($null -eq $document -and $null -ne $openError) {
            $openMessage = ($openError.Exception.Message | Out-String).ToLowerInvariant()
            if ($openMessage.Contains('could not fire the event')) {
                # Word can throw an event error even when the file opened.
                try {
                    $activeDocument = $word.ActiveDocument
                    if ($null -ne $activeDocument) {
                        $activePath = [string]$activeDocument.FullName
                        if (-not [string]::IsNullOrWhiteSpace($activePath)) {
                            $resolvedActive = [System.IO.Path]::GetFullPath($activePath)
                            $resolvedTarget = [System.IO.Path]::GetFullPath($Path)
                            if ($resolvedActive -eq $resolvedTarget) {
                                $document = $activeDocument
                            }
                        }
                    }
                } catch {}
            }
        }

        if ($null -eq $document -and $null -ne $openError) {
            throw $openError
        }

        $rows = New-Object System.Collections.Generic.List[object]
        $wdListNoNumbering = 0
        $paragraphCount = [int]$document.Paragraphs.Count

        for ($index = 1; $index -le $paragraphCount; $index++) {
            $paragraph = $null
            $range = $null
            $listFormat = $null
            $listObj = $null
            try {
                $paragraph = $document.Paragraphs.Item($index)
                $range = $paragraph.Range
                $text = Normalize-ParagraphText -Text $range.Text
                if ([string]::IsNullOrWhiteSpace($text)) { continue }

                $listFormat = $range.ListFormat
                $listType = [int]$listFormat.ListType
                $isList = $listType -ne $wdListNoNumbering

                if ($OnlyListRows -and -not $isList) { continue }

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
                try { $styleName = [string]$range.Style } catch {}

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

                if ($RowLimit -gt 0 -and $rows.Count -ge $RowLimit) { break }
            } finally {
                Release-ComObject -ComObject $listObj
                Release-ComObject -ComObject $listFormat
                Release-ComObject -ComObject $range
                Release-ComObject -ComObject $paragraph
            }
        }

        return $rows
    }
    finally {
        if ($null -ne $document) {
            try { $document.Close([ref]$false) } catch {}
        }
        if ($null -ne $word) {
            try { $word.Quit() } catch {}
        }
        Release-ComObject -ComObject $document
        Release-ComObject -ComObject $word
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

$resolvedPath = Resolve-Path -LiteralPath $DocxPath
$fullPath = $resolvedPath.Path

$attempts = [Math]::Max(1, $RetryCount)
$lastError = $null
$rows = $null
$repairAttempted = $false

for ($attempt = 1; $attempt -le $attempts; $attempt++) {
    try {
        Ensure-WordWorkfileEnvironment
        if ($KillWordBeforeStart) {
            Stop-WordProcesses
            Start-Sleep -Milliseconds 400
        }
        $rows = Invoke-ListInspectionPass -Path $fullPath -OnlyListRows:$OnlyLists -RowLimit $MaxRows
        $lastError = $null
        break
    } catch {
        $lastError = $_
        $message = ($lastError.Exception.Message | Out-String).ToLowerInvariant()
        $looksLikeWordStartupIssue = $message.Contains('could not create the work file') -or $message.Contains('experienced an error trying to open the file')
        if ($AttemptWordRepair -and -not $repairAttempted -and $looksLikeWordStartupIssue) {
            $repairAttempted = $true
            Invoke-WordRegistrationRepair
            continue
        }
        Stop-WordProcesses
        Start-Sleep -Seconds ([Math]::Min(3, $attempt))
    }
}

if ($null -eq $rows) {
    $message = if ($null -ne $lastError) { $lastError.Exception.Message } else { 'Unknown inspector failure.' }
    throw "list-inspector failed after $attempts attempt(s): $message"
}

$json = $rows | ConvertTo-Json -Depth 6
if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $outputParent = Split-Path -Parent $OutputPath
    if (-not [string]::IsNullOrWhiteSpace($outputParent)) {
        New-Item -ItemType Directory -Path $outputParent -Force | Out-Null
    }
    Set-Content -Path $OutputPath -Value $json -Encoding UTF8
}

$json

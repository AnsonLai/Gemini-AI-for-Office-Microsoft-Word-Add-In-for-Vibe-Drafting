# Authors the second golden fixture, a short privacy notice for a fictional clinic, through
# Word COM so every paragraph, style, list and table is Word's own markup. Uses its own hidden
# Word instance and quits only that instance; a Word window the user has open is never touched.
param([string]$Output = 'tests\fixtures\golden\privacy-notice.docx')
$ErrorActionPreference = 'Stop'
$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$target = [System.IO.Path]::GetFullPath((Join-Path $repoRoot $Output))
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $target) | Out-Null

$paragraphs = @(
    @{ text = 'PRIVACY NOTICE'; style = 'Title' },
    @{ text = 'Maple Street Dental Clinic'; style = 'Subtitle' },
    @{ text = 'Last updated: [Date]' },
    @{ text = 'This notice explains how Maple Street Dental Clinic collects, uses and protects your personal information when you visit or contact us.' },
    @{ text = 'Information We Collect'; style = 'Heading 1' },
    @{ text = 'Contact details such as your name, address, phone number and email.'; list = 'collect' },
    @{ text = 'Health information we need to provide your dental care.'; list = 'collect' },
    @{ text = 'Billing and insurance information.'; list = 'collect' },
    @{ text = 'How We Use Your Information'; style = 'Heading 1' },
    @{ text = 'To provide and coordinate your care.'; list = 'use' },
    @{ text = 'To process payments and insurance claims.'; list = 'use' },
    @{ text = 'To send appointment reminders.'; list = 'use' },
    @{ text = 'Sharing Your Information'; style = 'Heading 1' },
    @{ text = 'We share your information only with other health professionals involved in your care, your insurer, and service providers who help us run the clinic.'; footnote = 'Service providers are bound by written confidentiality agreements.' },
    @{ text = 'Your Rights'; style = 'Heading 1' },
    @{ text = 'Ask to see the information we hold about you.'; list = 'rights' },
    @{ text = 'Ask us to correct information that is wrong.'; list = 'rights' },
    @{ text = 'Withdraw your consent where we rely on it.'; list = 'rights' },
    @{ text = 'Contact Us'; style = 'Heading 1' },
    @{ text = "Privacy Officer`t[Name]"; table = $true },
    @{ text = "Email`tprivacy@maplestreetdental.example"; table = $true },
    @{ text = "Phone`t555-0100"; table = $true },
    @{ text = 'Acknowledgement'; style = 'Heading 1' },
    @{ text = 'I have read and understood this privacy notice.' },
    @{ text = 'Patient signature: ____________________   Date: __________' }
)

$existing = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$word = New-Object -ComObject Word.Application
$owned = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $_.Id -notin $existing })
if ($owned.Count -ne 1) {
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
    throw "Could not start one fixture-owned Word process (new: $($owned.Count)); no user Word session was changed."
}
$document = $null
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Add()
    $document.Range(0, 0).Text = ($paragraphs | ForEach-Object { $_.text }) -join "`r"
    for ($i = 0; $i -lt $paragraphs.Count; $i++) {
        if ($paragraphs[$i].style) { $document.Paragraphs.Item($i + 1).Style = $paragraphs[$i].style }
    }
    # Lists: two separate bullet lists and one numbered list, each its own Word list.
    $bullet = $word.ListGalleries.Item(1).ListTemplates.Item(1)
    $number = $word.ListGalleries.Item(2).ListTemplates.Item(1)
    foreach ($name in @('collect', 'use', 'rights')) {
        $indexes = @(for ($i = 0; $i -lt $paragraphs.Count; $i++) { if ($paragraphs[$i].list -eq $name) { $i + 1 } })
        $start = $document.Paragraphs.Item($indexes[0]).Range.Start
        $end = $document.Paragraphs.Item($indexes[-1]).Range.End - 1
        $template = if ($name -eq 'rights') { $number } else { $bullet }
        $document.Range($start, $end).ListFormat.ApplyListTemplateWithLevel($template, $false, 0, 2, 1)
    }
    # One footnote at the end of the sharing paragraph (the shape of library issue #6). Step 04
    # rewrites this paragraph, so the scenario checks that the reference survives Accept All.
    $sharingIndex = @(for ($i = 0; $i -lt $paragraphs.Count; $i++) { if ($paragraphs[$i].footnote) { $i + 1 } })[0]
    $sharingEnd = $document.Paragraphs.Item($sharingIndex).Range.End - 1
    $document.Footnotes.Add($document.Range($sharingEnd, $sharingEnd), [Type]::Missing, $paragraphs[$sharingIndex - 1].footnote) | Out-Null
    # Contact table last, so earlier paragraph indexes stay valid while authoring.
    $tableIndexes = @(for ($i = 0; $i -lt $paragraphs.Count; $i++) { if ($paragraphs[$i].table) { $i + 1 } })
    $tableRange = $document.Range($document.Paragraphs.Item($tableIndexes[0]).Range.Start, $document.Paragraphs.Item($tableIndexes[-1]).Range.End - 1)
    $table = $tableRange.ConvertToTable(1, $tableIndexes.Count, 2)
    $table.Borders.Enable = $true
    [string]$savePath = $target
    [int]$format = 16
    $document.SaveAs2([ref]$savePath, [ref]$format)
    $document.Close(0) | Out-Null
    $document = $null

    # Reopen and report what Word itself sees.
    $check = $word.Documents.OpenNoRepairDialog($target)
    $lines = for ($i = 1; $i -le $check.Paragraphs.Count; $i++) {
        $paragraph = $check.Paragraphs.Item($i)
        $text = $paragraph.Range.Text.TrimEnd([char]13, [char]7)
        $label = $paragraph.Range.ListFormat.ListString
        "{0,2} {1,-10} {2}{3}" -f $i, $paragraph.Style.NameLocal, $(if ($label) { "[$label] " } else { '' }), $text
    }
    $check.Close(0) | Out-Null
    $lines
    Write-Output "Saved $target"
} finally {
    if ($document) { $document.Close(0) | Out-Null }
    # Quit's SaveChanges argument is by-reference in COM; with no argument, nothing is saved.
    $word.Quit()
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
    Start-Sleep -Milliseconds 500
    $leftover = Get-Process -Id $owned[0].Id -ErrorAction SilentlyContinue
    if ($leftover) { Stop-Process -Id $owned[0].Id -Force }
}

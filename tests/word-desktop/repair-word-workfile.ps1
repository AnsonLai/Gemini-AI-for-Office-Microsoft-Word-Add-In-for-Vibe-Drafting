param()

$ErrorActionPreference = 'Stop'

function Stop-WordProcesses {
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

function Ensure-Directory {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    New-Item -ItemType Directory -Path $Path -Force | Out-Null
}

function Try-SetExpandString {
    param(
        [string]$Path,
        [string]$Name,
        [string]$Value
    )
    try {
        New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType ExpandString -Force | Out-Null
    } catch {
        Write-Warning "Could not set registry value $Path\\$Name ($($_.Exception.Message))"
    }
}

$localAppData = $env:LOCALAPPDATA
if ([string]::IsNullOrWhiteSpace($localAppData) -and -not [string]::IsNullOrWhiteSpace($env:USERPROFILE)) {
    $localAppData = Join-Path $env:USERPROFILE 'AppData\Local'
}
if ([string]::IsNullOrWhiteSpace($localAppData)) {
    throw "LOCALAPPDATA/USERPROFILE unavailable."
}

$cacheExpand = '%USERPROFILE%\AppData\Local\Microsoft\Windows\INetCache'
$cookiesExpand = '%USERPROFILE%\AppData\Local\Microsoft\Windows\INetCookies'
$historyExpand = '%USERPROFILE%\AppData\Local\Microsoft\Windows\History'

$userShellFolders = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
Try-SetExpandString -Path $userShellFolders -Name Cache -Value $cacheExpand
Try-SetExpandString -Path $userShellFolders -Name Cookies -Value $cookiesExpand
Try-SetExpandString -Path $userShellFolders -Name History -Value $historyExpand

$tempRoot = Join-Path $localAppData 'Temp'
$env:TEMP = $tempRoot
$env:TMP = $tempRoot

Ensure-Directory -Path $tempRoot
Ensure-Directory -Path (Join-Path $localAppData 'Microsoft\Windows\INetCache')
Ensure-Directory -Path (Join-Path $localAppData 'Microsoft\Windows\INetCache\Content.Word')
Ensure-Directory -Path (Join-Path $localAppData 'Microsoft\Windows\INetCache\Content.MSO')
Ensure-Directory -Path (Join-Path $localAppData 'Microsoft\Windows\INetCookies')
Ensure-Directory -Path (Join-Path $localAppData 'Microsoft\Windows\History')

Stop-WordProcesses
Start-Sleep -Milliseconds 600

try {
    $wordExe = (Get-Command WINWORD.EXE -ErrorAction SilentlyContinue).Source
    if (-not [string]::IsNullOrWhiteSpace($wordExe)) {
        $proc = Start-Process -FilePath $wordExe -ArgumentList '/r','/q' -PassThru
        $null = $proc.WaitForExit(30000)
        if (-not $proc.HasExited) {
            try { $proc.Kill() } catch {}
        }
    }
} finally {
    Stop-WordProcesses
}

Write-Host "Word workfile repair steps completed."

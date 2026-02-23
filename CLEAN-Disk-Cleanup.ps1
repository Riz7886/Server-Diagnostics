param(
    [int]$TempFileAgeDays = 7,
    [int]$LogRetentionDays = 30,
    [int]$DiskThreshold = 70
)

$ErrorActionPreference = "SilentlyContinue"
$ReportPath = "C:\Temp\DiskCleanup_Reports"
$LogFile = Join-Path $ReportPath "Cleanup_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

if (!(Test-Path $ReportPath)) {
    New-Item -Path $ReportPath -ItemType Directory -Force | Out-Null
}

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $LogFile -Append
    Write-Host $Message
}

function Get-DiskSpace {
    $disk = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'"
    return @{
        TotalGB = [math]::Round($disk.Size / 1GB, 2)
        FreeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
        UsedPercent = [math]::Round((($disk.Size - $disk.FreeSpace) / $disk.Size) * 100, 2)
    }
}

function Get-FolderSize {
    param([string]$Path)
    if (!(Test-Path $Path)) { return 0 }
    $size = (Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
    if ($size) { return [math]::Round($size / 1MB, 2) } else { return 0 }
}

function Remove-OldFiles {
    param([string]$Path, [int]$Days, [string]$Description)
    if (!(Test-Path $Path)) { return 0 }
    $cutoffDate = (Get-Date).AddDays(-$Days)
    $sizeBefore = Get-FolderSize -Path $Path
    $files = Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue | Where-Object { !$_.PSIsContainer -and $_.LastWriteTime -lt $cutoffDate }
    $fileCount = 0
    foreach ($file in $files) {
        Remove-Item -Path $file.FullName -Force -ErrorAction SilentlyContinue
        if ($?) { $fileCount++ }
    }
    $sizeAfter = Get-FolderSize -Path $Path
    $spaceFreed = $sizeBefore - $sizeAfter
    Write-Log "Cleaned $Description - Deleted $fileCount files - Freed $spaceFreed MB"
    return $spaceFreed
}

Write-Log "Starting disk cleanup on $env:COMPUTERNAME"

$diskBefore = Get-DiskSpace
Write-Log "Disk before: $($diskBefore.FreeGB) GB free ($($diskBefore.UsedPercent)% used)"

if ($diskBefore.UsedPercent -lt $DiskThreshold) {
    Write-Log "Disk usage below threshold. Exiting."
    exit 0
}

$totalSpaceFreed = 0

$totalSpaceFreed += Remove-OldFiles -Path "C:\Windows\Temp" -Days $TempFileAgeDays -Description "Windows Temp"
$totalSpaceFreed += Remove-OldFiles -Path "C:\Windows\Prefetch" -Days 30 -Description "Prefetch"

$userProfiles = Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue
foreach ($profile in $userProfiles) {
    $tempPath = Join-Path $profile.FullName "AppData\Local\Temp"
    if (Test-Path $tempPath) {
        $totalSpaceFreed += Remove-OldFiles -Path $tempPath -Days $TempFileAgeDays -Description "User Temp $($profile.Name)"
    }
}

if (Test-Path "C:\inetpub\logs\LogFiles") {
    $totalSpaceFreed += Remove-OldFiles -Path "C:\inetpub\logs\LogFiles" -Days $LogRetentionDays -Description "IIS Logs"
}

if (Test-Path "C:\Windows\SoftwareDistribution\Download") {
    Stop-Service wuauserv -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    $totalSpaceFreed += Remove-OldFiles -Path "C:\Windows\SoftwareDistribution\Download" -Days 30 -Description "Windows Update Cache"
    Start-Service wuauserv -ErrorAction SilentlyContinue
}

if (Test-Path "C:\ProgramData\Microsoft\Windows\WER") {
    $totalSpaceFreed += Remove-OldFiles -Path "C:\ProgramData\Microsoft\Windows\WER" -Days $TempFileAgeDays -Description "Error Reports"
}

$sizeBefore = Get-FolderSize -Path 'C:\$Recycle.Bin'
$shell = New-Object -ComObject Shell.Application
$recycleBin = $shell.Namespace(0xA)
$itemsDeleted = 0
$cutoffDate = (Get-Date).AddDays(-7)
foreach ($item in $recycleBin.Items()) {
    $itemDate = [datetime]$item.ExtendedProperty("System.DateModified")
    if ($itemDate -lt $cutoffDate) {
        Remove-Item -Path $item.Path -Recurse -Force -ErrorAction SilentlyContinue
        if ($?) { $itemsDeleted++ }
    }
}
$sizeAfter = Get-FolderSize -Path 'C:\$Recycle.Bin'
$spaceFreed = $sizeBefore - $sizeAfter
Write-Log "Cleaned Recycle Bin - Deleted $itemsDeleted items - Freed $spaceFreed MB"
$totalSpaceFreed += $spaceFreed

$browserPaths = @(
    "AppData\Local\Microsoft\Windows\INetCache",
    "AppData\Local\Microsoft\Windows\Temporary Internet Files",
    "AppData\Local\Google\Chrome\User Data\Default\Cache",
    "AppData\Local\Mozilla\Firefox\Profiles\*\cache2"
)
foreach ($profile in $userProfiles) {
    foreach ($browserPath in $browserPaths) {
        $fullPath = Join-Path $profile.FullName $browserPath
        if (Test-Path $fullPath) {
            $totalSpaceFreed += Remove-OldFiles -Path $fullPath -Days $TempFileAgeDays -Description "Browser Cache $($profile.Name)"
        }
    }
}

if (Test-Path "C:\Windows\Installer\`$PatchCache`$") {
    $totalSpaceFreed += Remove-OldFiles -Path "C:\Windows\Installer\`$PatchCache`$" -Days 90 -Description "Installer Cache"
}

$volumeCachesKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches"
$safeCategories = @(
    "Active Setup Temp Folders",
    "Downloaded Program Files",
    "Internet Cache Files",
    "Memory Dump Files",
    "Offline Pages Files",
    "Old ChkDsk Files",
    "Recycle Bin",
    "Setup Log Files",
    "Temporary Files",
    "Temporary Setup Files",
    "Thumbnail Cache",
    "Windows Error Reporting Files"
)
foreach ($category in $safeCategories) {
    $keyPath = Join-Path $volumeCachesKey $category
    if (Test-Path $keyPath) {
        Set-ItemProperty -Path $keyPath -Name "StateFlags0064" -Value 2 -ErrorAction SilentlyContinue
    }
}
Start-Process cleanmgr.exe -ArgumentList "/sagerun:64" -Wait -WindowStyle Hidden
Write-Log "Ran Disk Cleanup utility"

$diskAfter = Get-DiskSpace
$actualSpaceFreedGB = [math]::Round($diskAfter.FreeGB - $diskBefore.FreeGB, 2)

Write-Log "Cleanup completed"
Write-Log "Space freed: $actualSpaceFreedGB GB"
Write-Log "Free space now: $($diskAfter.FreeGB) GB ($($diskAfter.UsedPercent)% used)"

$htmlReport = @"
<!DOCTYPE html>
<html>
<head>
<style>
body{font-family:Arial;margin:20px;background:#f5f5f5}
.container{max-width:800px;margin:0 auto;background:white;padding:30px;box-shadow:0 0 10px rgba(0,0,0,0.1)}
h1{color:#2E75B6;border-bottom:3px solid #2E75B6;padding-bottom:10px}
table{width:100%;border-collapse:collapse;margin:20px 0}
th{background:#2E75B6;color:white;padding:12px;text-align:left}
td{padding:10px;border:1px solid #ddd}
tr:nth-child(even){background:#f9f9f9}
.success{background:#d4edda;color:#155724;padding:15px;border-radius:5px;margin:20px 0;font-size:24px;font-weight:bold}
</style>
</head>
<body>
<div class="container">
<h1>Disk Cleanup Report</h1>
<p><strong>Server:</strong> $env:COMPUTERNAME</p>
<p><strong>Date:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
<div class="success">Space Recovered: $actualSpaceFreedGB GB</div>
<table>
<tr><th></th><th>Before</th><th>After</th><th>Change</th></tr>
<tr><td>Total Space</td><td>$($diskBefore.TotalGB) GB</td><td>$($diskAfter.TotalGB) GB</td><td>-</td></tr>
<tr><td>Free Space</td><td>$($diskBefore.FreeGB) GB</td><td>$($diskAfter.FreeGB) GB</td><td style="color:green;font-weight:bold">+$actualSpaceFreedGB GB</td></tr>
<tr><td>Used Percentage</td><td>$($diskBefore.UsedPercent)%</td><td>$($diskAfter.UsedPercent)%</td><td>$([math]::Round($diskAfter.UsedPercent - $diskBefore.UsedPercent, 2))%</td></tr>
</table>
<h2>Cleaned Areas</h2>
<ul>
<li>Windows Temp files (older than $TempFileAgeDays days)</li>
<li>User Temp files (older than $TempFileAgeDays days)</li>
<li>IIS Logs (older than $LogRetentionDays days)</li>
<li>Windows Update cache</li>
<li>Windows Error Reports</li>
<li>Recycle Bin (older than 7 days)</li>
<li>Browser caches</li>
<li>Windows Disk Cleanup utility</li>
</ul>
<p><strong>Log:</strong> $LogFile</p>
</div>
</body>
</html>
"@

$htmlReportPath = Join-Path $ReportPath "Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
$htmlReport | Out-File -FilePath $htmlReportPath -Encoding UTF8
Start-Process $htmlReportPath

Write-Host ""
Write-Host "CLEANUP COMPLETED" -ForegroundColor Green
Write-Host "Space freed: $actualSpaceFreedGB GB" -ForegroundColor Green
Write-Host "Free space: $($diskAfter.FreeGB) GB ($($diskAfter.UsedPercent)% used)" -ForegroundColor Green
Write-Host "Report: $htmlReportPath" -ForegroundColor Cyan
Write-Host ""

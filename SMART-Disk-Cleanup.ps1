param(
    [int]$LargeFileSizeMB = 100,
    [int]$FileAgeDays = 90,
    [switch]$AutoClean
)

$ErrorActionPreference = "SilentlyContinue"
$ReportPath = "C:\Temp\DiskAnalysis"
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$LogFile = Join-Path $ReportPath "Analysis_$timestamp.log"
$ResultsFile = Join-Path $ReportPath "Results_$timestamp.csv"

if (!(Test-Path $ReportPath)) { New-Item -Path $ReportPath -ItemType Directory -Force | Out-Null }

function Write-Log {
    param([string]$Message, [string]$Color = "White")
    $msg = "$(Get-Date -Format 'HH:mm:ss') - $Message"
    Write-Host $msg -ForegroundColor $Color
    $msg | Out-File -FilePath $LogFile -Append
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  AGGRESSIVE DISK SPACE ANALYZER & CLEANUP" -ForegroundColor Cyan
Write-Host "  Finding: Large files, Old files, Wasted space" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

Write-Log "Starting analysis on $env:COMPUTERNAME" "Cyan"

$disk = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'"
$totalGB = [math]::Round($disk.Size / 1GB, 2)
$freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
$usedGB = [math]::Round(($disk.Size - $disk.FreeSpace) / 1GB, 2)
$usedPercent = [math]::Round((($disk.Size - $disk.FreeSpace) / $disk.Size) * 100, 2)

Write-Host "DISK STATUS:" -ForegroundColor Yellow
Write-Host "  Total: $totalGB GB" -ForegroundColor White
Write-Host "  Used:  $usedGB GB ($usedPercent%)" -ForegroundColor $(if ($usedPercent -gt 90) { "Red" } else { "Yellow" })
Write-Host "  Free:  $freeGB GB" -ForegroundColor $(if ($freeGB -lt 5) { "Red" } else { "Green" })
Write-Host ""

$findings = @()
$totalWastedGB = 0

Write-Host "ANALYSIS PHASE 1: Scanning for large files (>$LargeFileSizeMB MB)..." -ForegroundColor Yellow
Write-Host "This may take 5-10 minutes..." -ForegroundColor Gray
Write-Host ""

$extensionsToCheck = @{
    "Old Installers" = @("*.msi", "*.exe", "*.msp")
    "ISO Images" = @("*.iso")
    "Video Files" = @("*.mp4", "*.avi", "*.mkv", "*.mov", "*.wmv")
    "Old Backups" = @("*.bak", "*.backup", "*.old")
    "Log Files" = @("*.log", "*.txt")
    "Dump Files" = @("*.dmp", "*.mdmp")
    "CAB Files" = @("*.cab")
    "Compressed" = @("*.zip", "*.rar", "*.7z", "*.tar", "*.gz")
}

$searchPaths = @("C:\Windows\Temp", "C:\Temp", "C:\Users", "C:\inetpub", "C:\Windows\SoftwareDistribution")

foreach ($path in $searchPaths) {
    if (!(Test-Path $path)) { continue }
    
    Write-Host "Scanning: $path" -ForegroundColor Gray
    
    foreach ($category in $extensionsToCheck.Keys) {
        $patterns = $extensionsToCheck[$category]
        
        foreach ($pattern in $patterns) {
            $files = Get-ChildItem -Path $path -Filter $pattern -Recurse -File -ErrorAction SilentlyContinue | 
                     Where-Object { ($_.Length / 1MB) -ge $LargeFileSizeMB }
            
            foreach ($file in $files) {
                $ageInDays = ((Get-Date) - $file.LastAccessTime).Days
                $sizeMB = [math]::Round($file.Length / 1MB, 2)
                $sizeGB = [math]::Round($file.Length / 1GB, 2)
                
                $canDelete = $false
                $reason = ""
                
                if ($ageInDays -gt $FileAgeDays) {
                    $canDelete = $true
                    $reason = "Not accessed in $ageInDays days"
                } elseif ($category -eq "Old Installers" -and $ageInDays -gt 30) {
                    $canDelete = $true
                    $reason = "Old installer"
                } elseif ($category -eq "Dump Files") {
                    $canDelete = $true
                    $reason = "Crash dump file"
                } elseif ($category -eq "Old Backups" -and $ageInDays -gt 30) {
                    $canDelete = $true
                    $reason = "Old backup file"
                } elseif ($file.FullName -like "*\Windows\SoftwareDistribution\*") {
                    $canDelete = $true
                    $reason = "Old Windows Update file"
                } elseif ($file.FullName -like "*\Windows\Temp\*" -and $ageInDays -gt 7) {
                    $canDelete = $true
                    $reason = "Old temp file"
                }
                
                if ($canDelete) {
                    $totalWastedGB += $sizeGB
                    
                    $findings += [PSCustomObject]@{
                        Category = $category
                        FilePath = $file.FullName
                        FileName = $file.Name
                        SizeMB = $sizeMB
                        SizeGB = $sizeGB
                        LastAccessed = $file.LastAccessTime
                        AgeInDays = $ageInDays
                        Reason = $reason
                        CanDelete = $canDelete
                    }
                }
            }
        }
    }
}

Write-Host ""
Write-Host "ANALYSIS PHASE 2: Checking common waste areas..." -ForegroundColor Yellow

$wasteAreas = @(
    @{Path="C:\Windows\SoftwareDistribution\Download"; Name="Windows Update Cache"},
    @{Path="C:\Windows\Temp"; Name="Windows Temp"},
    @{Path="C:\ProgramData\Microsoft\Windows\WER"; Name="Error Reports"},
    @{Path="C:\inetpub\logs"; Name="IIS Logs"},
    @{Path="C:\Windows\Logs"; Name="Windows Logs"},
    @{Path="C:\Windows\Panther"; Name="Windows Setup Logs"}
)

foreach ($area in $wasteAreas) {
    if (Test-Path $area.Path) {
        $size = (Get-ChildItem $area.Path -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
        if ($size) {
            $sizeGB = [math]::Round($size / 1GB, 2)
            if ($sizeGB -gt 0.5) {
                Write-Host "  Found: $($area.Name) - $sizeGB GB" -ForegroundColor Yellow
            }
        }
    }
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Green
Write-Host "  ANALYSIS COMPLETE" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Green
Write-Host ""

$sortedFindings = $findings | Sort-Object SizeGB -Descending

Write-Host "TOP 20 SPACE WASTERS:" -ForegroundColor Cyan
Write-Host ""

$top20 = $sortedFindings | Select-Object -First 20
$displayNum = 1
foreach ($item in $top20) {
    $color = if ($item.SizeGB -gt 1) { "Red" } elseif ($item.SizeMB -gt 500) { "Yellow" } else { "White" }
    Write-Host "$displayNum. $($item.Category) - $($item.SizeGB) GB - $($item.FileName)" -ForegroundColor $color
    Write-Host "   Path: $($item.FilePath)" -ForegroundColor Gray
    Write-Host "   Age: $($item.AgeInDays) days | Reason: $($item.Reason)" -ForegroundColor Gray
    Write-Host ""
    $displayNum++
}

Write-Host "SUMMARY:" -ForegroundColor Cyan
Write-Host "  Total files found: $($findings.Count)" -ForegroundColor White
Write-Host "  Total wasted space: $([math]::Round($totalWastedGB, 2)) GB" -ForegroundColor Red
Write-Host "  Potential savings: $([math]::Round($totalWastedGB, 2)) GB" -ForegroundColor Green
Write-Host ""

$findings | Export-Csv -Path $ResultsFile -NoTypeInformation
Write-Host "Full results saved to: $ResultsFile" -ForegroundColor Cyan
Write-Host ""

if ($AutoClean) {
    Write-Host "AUTO-CLEAN MODE: Deleting files..." -ForegroundColor Yellow
    Write-Host ""
    
    $deletedCount = 0
    $deletedGB = 0
    
    foreach ($item in $sortedFindings) {
        try {
            if (Test-Path $item.FilePath) {
                Remove-Item -Path $item.FilePath -Force -ErrorAction Stop
                $deletedCount++
                $deletedGB += $item.SizeGB
                Write-Host "Deleted: $($item.FileName) - $($item.SizeGB) GB" -ForegroundColor Green
            }
        } catch {
            Write-Host "Skipped (in use): $($item.FileName)" -ForegroundColor Yellow
        }
    }
    
    Write-Host ""
    Write-Host "CLEANUP COMPLETE:" -ForegroundColor Green
    Write-Host "  Files deleted: $deletedCount" -ForegroundColor White
    Write-Host "  Space freed: $([math]::Round($deletedGB, 2)) GB" -ForegroundColor Green
    Write-Host ""
} else {
    Write-Host "TO DELETE THESE FILES, RUN:" -ForegroundColor Yellow
    Write-Host "  .\SMART-Disk-Cleanup.ps1 -AutoClean" -ForegroundColor White
    Write-Host ""
}

$htmlReport = @"
<!DOCTYPE html>
<html>
<head>
<style>
body{font-family:Arial;margin:20px;background:#f5f5f5}
.container{max-width:1200px;margin:0 auto;background:white;padding:30px;box-shadow:0 0 10px rgba(0,0,0,0.1)}
h1{color:#dc2626;border-bottom:3px solid #dc2626;padding-bottom:10px}
.summary{background:#fef3c7;border-left:5px solid #f59e0b;padding:20px;margin:20px 0}
.summary h2{margin:0 0 10px 0;color:#92400e}
.metric{font-size:32px;font-weight:bold;color:#dc2626}
table{width:100%;border-collapse:collapse;margin:20px 0}
th{background:#dc2626;color:white;padding:12px;text-align:left}
td{padding:10px;border:1px solid #ddd;font-size:13px}
tr:nth-child(even){background:#f9f9f9}
.category{font-weight:bold;color:#2563eb}
.large{background:#fee2e2;font-weight:bold}
.medium{background:#fef3c7}
</style>
</head>
<body>
<div class="container">
<h1>Disk Space Analysis Report</h1>
<p><strong>Server:</strong> $env:COMPUTERNAME</p>
<p><strong>Analysis Date:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>

<div class="summary">
<h2>Critical Findings</h2>
<p>Total Wasted Space: <span class="metric">$([math]::Round($totalWastedGB, 2)) GB</span></p>
<p>Files Found: <span style="font-size:24px;font-weight:bold">$($findings.Count)</span></p>
<p>Current Free Space: <span style="font-size:24px;font-weight:bold;color:#059669">$freeGB GB</span></p>
<p>Potential Free Space After Cleanup: <span style="font-size:24px;font-weight:bold;color:#059669">$([math]::Round($freeGB + $totalWastedGB, 2)) GB</span></p>
</div>

<h2>Top Space Wasters</h2>
<table>
<tr>
<th>#</th>
<th>Category</th>
<th>File Name</th>
<th>Size (GB)</th>
<th>Age (Days)</th>
<th>Reason</th>
<th>Location</th>
</tr>
"@

$rowNum = 1
foreach ($item in ($sortedFindings | Select-Object -First 50)) {
    $rowClass = if ($item.SizeGB -gt 1) { "large" } elseif ($item.SizeMB -gt 500) { "medium" } else { "" }
    $htmlReport += @"
<tr class="$rowClass">
<td>$rowNum</td>
<td class="category">$($item.Category)</td>
<td>$($item.FileName)</td>
<td style="font-weight:bold">$($item.SizeGB)</td>
<td>$($item.AgeInDays)</td>
<td>$($item.Reason)</td>
<td style="font-size:11px">$($item.FilePath)</td>
</tr>
"@
    $rowNum++
}

$htmlReport += @"
</table>

<h2>Cleanup Recommendation</h2>
<p style="background:#d1fae5;border-left:5px solid #059669;padding:15px">
<strong>Action Required:</strong> Run the cleanup script with -AutoClean flag to delete these files and free up $([math]::Round($totalWastedGB, 2)) GB of space.
</p>

<p><strong>Command:</strong><br>
<code style="background:#f3f4f6;padding:10px;display:block;margin:10px 0">
.\SMART-Disk-Cleanup.ps1 -AutoClean
</code>
</p>

</div>
</body>
</html>
"@

$htmlReportPath = Join-Path $ReportPath "Report_$timestamp.html"
$htmlReport | Out-File -FilePath $htmlReportPath -Encoding UTF8
Start-Process $htmlReportPath

Write-Host "HTML Report opened in browser!" -ForegroundColor Green
Write-Host "Report location: $htmlReportPath" -ForegroundColor Cyan
Write-Host ""
Write-Log "Analysis complete. Report generated." "Green"

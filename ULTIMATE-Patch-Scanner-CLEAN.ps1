# ULTIMATE PATCH COMPLIANCE SCANNER v2.0
# Universal Multi-Format Patch Compliance Analyzer
# Author: Syed Ahmad
# Date: February 27, 2026
# NO MODIFICATIONS NEEDED - JUST RUN IT!

[CmdletBinding()]
param(
    [string]$ScanPath = "C:\Report-Alert",
    [string]$ServerListFile = "",
    [string]$OutputPath = "",
    [switch]$ExportCSV = $true
)

# Setup output path
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $ScanPath "Reports"
}

if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  ULTIMATE PATCH COMPLIANCE SCANNER v2.0" -ForegroundColor Yellow
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Scan Path: $ScanPath" -ForegroundColor White
Write-Host "Output Path: $OutputPath" -ForegroundColor White
Write-Host ""

# Function to read Excel files
function Read-ExcelFile {
    param([string]$FilePath)
    
    Write-Host "Reading Excel file: $([System.IO.Path]::GetFileName($FilePath))" -ForegroundColor Cyan
    
    try {
        $Excel = New-Object -ComObject Excel.Application
        $Excel.Visible = $false
        $Excel.DisplayAlerts = $false
        
        $Workbook = $Excel.Workbooks.Open($FilePath)
        $Worksheet = $Workbook.Sheets.Item(1)
        
        $Range = $Worksheet.UsedRange
        $RowCount = $Range.Rows.Count
        $ColCount = $Range.Columns.Count
        
        $Headers = @()
        for ($col = 1; $col -le $ColCount; $col++) {
            $Headers += $Worksheet.Cells.Item(1, $col).Text
        }
        
        $Data = @()
        for ($row = 2; $row -le $RowCount; $row++) {
            $RowData = @{}
            for ($col = 1; $col -le $Headers.Count; $col++) {
                $RowData[$Headers[$col - 1]] = $Worksheet.Cells.Item($row, $col).Text
            }
            $Data += [PSCustomObject]$RowData
        }
        
        $Workbook.Close($false)
        $Excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
        
        Write-Host "  Loaded $($Data.Count) rows" -ForegroundColor Green
        return $Data
        
    } catch {
        Write-Host "  ERROR: $_" -ForegroundColor Red
        return @()
    }
}

# Function to find IPs in text
function Find-IPAddresses {
    param([string]$Text)
    
    $IPPattern = '\b(?:\d{1,3}\.){3}\d{1,3}\b'
    $Matches = [regex]::Matches($Text, $IPPattern)
    return $Matches | ForEach-Object { $_.Value } | Select-Object -Unique
}

# Function to analyze reports
function Analyze-Reports {
    param([string]$Path)
    
    Write-Host "Scanning for reports..." -ForegroundColor Cyan
    
    $ReportFiles = Get-ChildItem -Path $Path -Include "*.html","*.htm","*.txt","*.csv" -Recurse -ErrorAction SilentlyContinue
    
    Write-Host "  Found $($ReportFiles.Count) report files" -ForegroundColor Yellow
    
    $AllServers = @()
    
    foreach ($File in $ReportFiles) {
        Write-Host "  Processing: $($File.Name)" -ForegroundColor Gray
        
        try {
            $Content = Get-Content $File.FullName -Raw -ErrorAction Stop
            $IPs = Find-IPAddresses -Text $Content
            
            foreach ($IP in $IPs) {
                $IPIndex = $Content.IndexOf($IP)
                if ($IPIndex -ge 0) {
                    $ContextStart = [Math]::Max(0, $IPIndex - 500)
                    $ContextEnd = [Math]::Min($Content.Length, $IPIndex + 500)
                    $Context = $Content.Substring($ContextStart, $ContextEnd - $ContextStart)
                    
                    $Status = "UNKNOWN"
                    if ($Context -match "Compliant") { $Status = "COMPLIANT" }
                    if ($Context -match "NonCompliant|Non-Compliant") { $Status = "NON-COMPLIANT" }
                    
                    $AllServers += [PSCustomObject]@{
                        IP = $IP
                        Status = $Status
                        SourceReport = $File.Name
                        LastChecked = $File.LastWriteTime
                    }
                }
            }
        } catch {
            Write-Host "    ERROR: $_" -ForegroundColor Red
        }
    }
    
    $UniqueServers = $AllServers | Group-Object IP | ForEach-Object {
        $_.Group | Sort-Object LastChecked -Descending | Select-Object -First 1
    }
    
    Write-Host "  Found $($UniqueServers.Count) unique servers" -ForegroundColor Green
    return $UniqueServers
}

# Main execution
Write-Host "Step 1: Analyzing compliance reports..." -ForegroundColor Cyan
$ScannedServers = Analyze-Reports -Path $ScanPath

if ($ScannedServers.Count -eq 0) {
    Write-Host "ERROR: No servers found in reports!" -ForegroundColor Red
    exit 1
}

# Calculate statistics
$TotalServers = $ScannedServers.Count
$Compliant = ($ScannedServers | Where-Object { $_.Status -eq "COMPLIANT" }).Count
$NonCompliant = ($ScannedServers | Where-Object { $_.Status -eq "NON-COMPLIANT" }).Count
$Unknown = $TotalServers - $Compliant - $NonCompliant

Write-Host ""
Write-Host "Step 2: Generating HTML report..." -ForegroundColor Cyan

$CompliancePercent = if ($TotalServers -gt 0) { [math]::Round(($Compliant / $TotalServers) * 100, 1) } else { 0 }

$HTML = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Patch Compliance Report</title>
    <style>
        body { font-family: Arial; background: #f0f0f0; padding: 20px; }
        .container { max-width: 1400px; margin: 0 auto; background: white; border-radius: 10px; box-shadow: 0 0 20px rgba(0,0,0,0.1); }
        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 40px; text-align: center; border-radius: 10px 10px 0 0; }
        .header h1 { margin: 0; font-size: 36px; }
        .stats { display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; padding: 30px; background: #f8f9fa; }
        .stat-card { background: white; padding: 25px; border-radius: 10px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
        .stat-label { font-size: 14px; color: #666; margin-bottom: 10px; text-transform: uppercase; }
        .stat-value { font-size: 36px; font-weight: bold; color: #667eea; }
        .stat-card.success .stat-value { color: #10b981; }
        .stat-card.danger .stat-value { color: #ef4444; }
        .content { padding: 30px; }
        table { width: 100%; border-collapse: collapse; margin: 20px 0; }
        th { background: #667eea; color: white; padding: 15px; text-align: left; }
        td { padding: 12px 15px; border-bottom: 1px solid #e0e0e0; }
        tr:hover { background: #f5f5f5; }
        .badge { padding: 6px 12px; border-radius: 20px; font-weight: bold; font-size: 12px; display: inline-block; }
        .badge.success { background: #d1fae5; color: #065f46; }
        .badge.danger { background: #fee2e2; color: #991b1b; }
        .badge.unknown { background: #e5e7eb; color: #374151; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Patch Compliance Report</h1>
            <p>Generated: $(Get-Date -Format 'MMMM dd, yyyy - hh:mm:ss tt')</p>
        </div>
        
        <div class="stats">
            <div class="stat-card">
                <div class="stat-label">Total Servers</div>
                <div class="stat-value">$TotalServers</div>
            </div>
            <div class="stat-card success">
                <div class="stat-label">Compliant</div>
                <div class="stat-value">$Compliant</div>
            </div>
            <div class="stat-card danger">
                <div class="stat-label">Non-Compliant</div>
                <div class="stat-value">$NonCompliant</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Unknown</div>
                <div class="stat-value">$Unknown</div>
            </div>
        </div>
        
        <div class="content">
            <h2>Server Details</h2>
            <table>
                <thead>
                    <tr>
                        <th>IP Address</th>
                        <th>Status</th>
                        <th>Source Report</th>
                        <th>Last Checked</th>
                    </tr>
                </thead>
                <tbody>
"@

foreach ($Server in $ScannedServers) {
    $BadgeClass = switch ($Server.Status) {
        "COMPLIANT" { "success" }
        "NON-COMPLIANT" { "danger" }
        default { "unknown" }
    }
    
    $HTML += @"
                    <tr>
                        <td><strong>$($Server.IP)</strong></td>
                        <td><span class="badge $BadgeClass">$($Server.Status)</span></td>
                        <td>$($Server.SourceReport)</td>
                        <td>$($Server.LastChecked)</td>
                    </tr>
"@
}

$HTML += @"
                </tbody>
            </table>
        </div>
    </div>
</body>
</html>
"@

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$HTMLFile = Join-Path $OutputPath "PatchCompliance_$Timestamp.html"
$HTML | Out-File -FilePath $HTMLFile -Encoding UTF8 -Force

Write-Host "  HTML Report: $HTMLFile" -ForegroundColor Green

# Export CSV
if ($ExportCSV) {
    $CSVFile = Join-Path $OutputPath "PatchCompliance_$Timestamp.csv"
    $ScannedServers | Export-Csv -Path $CSVFile -NoTypeInformation -Encoding UTF8
    Write-Host "  CSV Export: $CSVFile" -ForegroundColor Green
}

# Summary
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  SCAN COMPLETE" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Total Servers: $TotalServers" -ForegroundColor White
Write-Host "Compliant: $Compliant" -ForegroundColor Green
Write-Host "Non-Compliant: $NonCompliant" -ForegroundColor Yellow
Write-Host "Unknown: $Unknown" -ForegroundColor Gray
Write-Host ""
Write-Host "Opening HTML report..." -ForegroundColor Yellow
Start-Process $HTMLFile

Write-Host ""
Write-Host "DONE! Check your report! " -ForegroundColor Green

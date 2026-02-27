# PATCH SCANNER - SIMPLE ROBUST VERSION
# Handles any Excel column names
# Author: Syed Ahmad

param(
    [string]$ScanPath = "C:\Reports-alerts",
    [string]$OutputPath = "C:\Reports"
)

Write-Host ""
Write-Host "PATCH SCANNER - SIMPLE VERSION" -ForegroundColor Yellow
Write-Host ""

if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

# Read Excel - Simple method
function Read-ServerList {
    param([string]$Path)
    
    Write-Host "Looking for server list..." -ForegroundColor Cyan
    
    $ExcelFiles = Get-ChildItem -Path (Split-Path $Path -Parent) -Filter "*.xlsx" -ErrorAction SilentlyContinue
    $ServerListFile = $ExcelFiles | Where-Object { $_.Name -match "rick|patch|server|list" } | Select-Object -First 1
    
    if (-not $ServerListFile) {
        Write-Host "  No server list found - will scan all IPs in reports" -ForegroundColor Yellow
        return @()
    }
    
    Write-Host "  Found: $($ServerListFile.Name)" -ForegroundColor Green
    
    try {
        $Excel = New-Object -ComObject Excel.Application
        $Excel.Visible = $false
        $Excel.DisplayAlerts = $false
        
        $Workbook = $Excel.Workbooks.Open($ServerListFile.FullName)
        $Worksheet = $Workbook.Sheets.Item(1)
        $Range = $Worksheet.UsedRange
        
        # Just get all IPs from first column
        $IPs = @()
        for ($row = 1; $row -le $Range.Rows.Count; $row++) {
            $Value = $Worksheet.Cells.Item($row, 1).Text
            if ($Value -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
                $IPs += $Value
            }
        }
        
        $Workbook.Close($false)
        $Excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
        
        Write-Host "  Loaded $($IPs.Count) IPs from list" -ForegroundColor Green
        return $IPs
        
    } catch {
        Write-Host "  Error: $_" -ForegroundColor Red
        return @()
    }
}

# Scan reports
function Scan-Reports {
    param([string]$Path)
    
    Write-Host ""
    Write-Host "Scanning reports in: $Path" -ForegroundColor Cyan
    
    $Files = Get-ChildItem -Path $Path -Include "*.msg","*.html","*.txt" -Recurse -ErrorAction SilentlyContinue
    Write-Host "  Found $($Files.Count) files" -ForegroundColor Green
    
    $Results = @{}
    
    foreach ($File in $Files) {
        Write-Host "  Processing: $($File.Name)" -ForegroundColor Gray
        
        try {
            # Read content
            $Content = ""
            if ($File.Extension -eq ".msg") {
                try {
                    $Outlook = New-Object -ComObject Outlook.Application
                    $Msg = $Outlook.Session.OpenSharedItem($File.FullName)
                    $Content = $Msg.Body
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
                } catch {
                    $Content = [System.IO.File]::ReadAllText($File.FullName)
                }
            } else {
                $Content = Get-Content $File.FullName -Raw
            }
            
            # Find IPs
            $IPMatches = [regex]::Matches($Content, '\b(?:\d{1,3}\.){3}\d{1,3}\b')
            
            foreach ($Match in $IPMatches) {
                $IP = $Match.Value
                
                # Get context
                $Index = $Content.IndexOf($IP)
                $Start = [Math]::Max(0, $Index - 500)
                $End = [Math]::Min($Content.Length, $Index + 500)
                $Context = $Content.Substring($Start, $End - $Start)
                
                # Check status
                $Status = "UNKNOWN"
                if ($Context -match "Compliant" -and $Context -notmatch "NonCompliant") {
                    $Status = "COMPLIANT"
                } elseif ($Context -match "NonCompliant|Non-Compliant") {
                    $Status = "NON-COMPLIANT"
                }
                
                if (-not $Results.ContainsKey($IP)) {
                    $Results[$IP] = @{
                        IP = $IP
                        Status = $Status
                        Source = $File.Name
                    }
                }
            }
            
        } catch {
            Write-Host "    Error: $_" -ForegroundColor Red
        }
    }
    
    Write-Host "  Found $($Results.Count) unique IPs" -ForegroundColor Green
    return $Results
}

# Main
Write-Host "Step 1: Load server list" -ForegroundColor Yellow
$MyIPs = Read-ServerList -Path $ScanPath

Write-Host ""
Write-Host "Step 2: Scan reports" -ForegroundColor Yellow
$ReportData = Scan-Reports -Path $ScanPath

Write-Host ""
Write-Host "Step 3: Generate report" -ForegroundColor Yellow

# Compare if we have a list
if ($MyIPs.Count -gt 0) {
    Write-Host "  Comparing your $($MyIPs.Count) IPs against reports..." -ForegroundColor Cyan
    
    $Compliant = 0
    $NonCompliant = 0
    $NotFound = 0
    
    $HTML = "<html><head><title>Patch Comparison</title></head><body style='font-family:Arial;padding:20px;'>"
    $HTML += "<h1>Patch Compliance Report</h1>"
    $HTML += "<p>Your Server List: $($MyIPs.Count) servers</p>"
    $HTML += "<table border='1' cellpadding='10' style='border-collapse:collapse;'>"
    $HTML += "<tr style='background:#667eea;color:white;'><th>IP</th><th>Status</th><th>Source</th></tr>"
    
    foreach ($IP in $MyIPs) {
        if ($ReportData.ContainsKey($IP)) {
            $Status = $ReportData[$IP].Status
            $Source = $ReportData[$IP].Source
            
            if ($Status -eq "COMPLIANT") { $Compliant++ }
            elseif ($Status -eq "NON-COMPLIANT") { $NonCompliant++ }
            
            $Color = if ($Status -eq "COMPLIANT") { "#d1fae5" } elseif ($Status -eq "NON-COMPLIANT") { "#fee2e2" } else { "#fef3c7" }
            $HTML += "<tr style='background:$Color;'><td>$IP</td><td>$Status</td><td>$Source</td></tr>"
        } else {
            $NotFound++
            $HTML += "<tr style='background:#fef3c7;'><td>$IP</td><td>NOT IN REPORTS</td><td>N/A</td></tr>"
        }
    }
    
    $HTML += "</table>"
    $HTML += "<h2>Summary:</h2>"
    $HTML += "<p>Compliant: $Compliant | Non-Compliant: $NonCompliant | Not Found: $NotFound</p>"
    $HTML += "</body></html>"
    
    Write-Host ""
    Write-Host "RESULTS:" -ForegroundColor Green
    Write-Host "  Compliant: $Compliant" -ForegroundColor Green
    Write-Host "  Non-Compliant: $NonCompliant" -ForegroundColor Red
    Write-Host "  Not in Reports: $NotFound" -ForegroundColor Yellow
    
} else {
    Write-Host "  No server list - showing all IPs found in reports..." -ForegroundColor Yellow
    
    $HTML = "<html><head><title>Patch Scan</title></head><body style='font-family:Arial;padding:20px;'>"
    $HTML += "<h1>Patch Scan Results</h1>"
    $HTML += "<p>Total IPs found: $($ReportData.Count)</p>"
    $HTML += "<table border='1' cellpadding='10' style='border-collapse:collapse;'>"
    $HTML += "<tr style='background:#667eea;color:white;'><th>IP</th><th>Status</th><th>Source</th></tr>"
    
    foreach ($IP in $ReportData.Keys) {
        $Status = $ReportData[$IP].Status
        $Source = $ReportData[$IP].Source
        $Color = if ($Status -eq "COMPLIANT") { "#d1fae5" } elseif ($Status -eq "NON-COMPLIANT") { "#fee2e2" } else { "#fef3c7" }
        $HTML += "<tr style='background:$Color;'><td>$IP</td><td>$Status</td><td>$Source</td></tr>"
    }
    
    $HTML += "</table></body></html>"
}

# Save and open
$ReportFile = Join-Path $OutputPath "PatchReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
$HTML | Out-File -FilePath $ReportFile -Encoding UTF8

Write-Host ""
Write-Host "Report saved: $ReportFile" -ForegroundColor Green
Write-Host ""
Write-Host "Opening report..." -ForegroundColor Yellow

Start-Process $ReportFile

Write-Host ""
Write-Host "DONE!" -ForegroundColor Green

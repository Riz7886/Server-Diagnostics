<#
.SYNOPSIS
    Patch Compliance Checker - Compares Rick-patch-List against Weekly Patching Compliance Reports
    
.DESCRIPTION
    This script:
    1. Reads Rick-patch-List.xlsx from C:\Report-Alert
    2. Scans all Weekly Patching Compliance Report files (HTML/MSG)
    3. Compares server IPs and patch status
    4. Generates detailed HTML report showing:
       - Servers fully patched
       - Servers with missing patches
       - Servers not found in reports
       - Agent compliance status
    
.NOTES
    Author: Syed Ahmad
    Date: February 27, 2026
    Version: 1.0
#>

# ============================================================================
# CONFIGURATION
# ============================================================================

$ReportAlertFolder = "C:\Report-Alert"
$RickPatchListFile = "Rick-patch-List.xlsx"
$OutputReportPath = "C:\Report-Alert\Patch-Compliance-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

function Read-RickPatchList {
    param([string]$FilePath)
    
    Write-Host "Reading Rick-patch-List from: $FilePath" -ForegroundColor Cyan
    
    try {
        $Excel = New-Object -ComObject Excel.Application
        $Excel.Visible = $false
        $Excel.DisplayAlerts = $false
        
        $Workbook = $Excel.Workbooks.Open($FilePath)
        $Worksheet = $Workbook.Sheets.Item(1)
        
        $UsedRange = $Worksheet.UsedRange
        $RowCount = $UsedRange.Rows.Count
        
        $ServerList = @()
        
        # Read headers from first row
        $Headers = @()
        for ($col = 1; $col -le $UsedRange.Columns.Count; $col++) {
            $Headers += $Worksheet.Cells.Item(1, $col).Text
        }
        
        # Read data rows
        for ($row = 2; $row -le $RowCount; $row++) {
            $ServerObj = @{}
            for ($col = 1; $col -le $Headers.Count; $col++) {
                $HeaderName = $Headers[$col - 1]
                $CellValue = $Worksheet.Cells.Item($row, $col).Text
                $ServerObj[$HeaderName] = $CellValue
            }
            
            if ($ServerObj['server IP'] -or $ServerObj['ServerIP'] -or $ServerObj['IP']) {
                $ServerList += [PSCustomObject]$ServerObj
            }
        }
        
        $Workbook.Close($false)
        $Excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
        
        Write-Host "  Loaded $($ServerList.Count) servers from Rick-patch-List" -ForegroundColor Green
        return $ServerList
        
    } catch {
        Write-Host "  ERROR reading Excel file: $_" -ForegroundColor Red
        return @()
    }
}

function Parse-ComplianceReports {
    param([string]$FolderPath)
    
    Write-Host "`nScanning Weekly Patching Compliance Reports..." -ForegroundColor Cyan
    
    # Get all report files (HTML, MSG, TXT)
    $ReportFiles = Get-ChildItem -Path $FolderPath -Include "*.html","*.htm","*.msg","*.txt" -Recurse
    
    Write-Host "  Found $($ReportFiles.Count) report files" -ForegroundColor Yellow
    
    $AllServersFromReports = @()
    
    foreach ($File in $ReportFiles) {
        Write-Host "  Processing: $($File.Name)" -ForegroundColor Gray
        
        try {
            # Read file content
            if ($File.Extension -eq ".msg") {
                # For MSG files, extract using Outlook
                $Outlook = New-Object -ComObject Outlook.Application
                $MailItem = $Outlook.CreateItemFromTemplate($File.FullName)
                $Content = $MailItem.HTMLBody
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
            } else {
                $Content = Get-Content $File.FullName -Raw
            }
            
            # Extract server information using regex patterns
            # Pattern 1: IP addresses
            $IPMatches = [regex]::Matches($Content, '\b(?:\d{1,3}\.){3}\d{1,3}\b')
            
            # Pattern 2: Server names
            $ServerMatches = [regex]::Matches($Content, '(?:EC2AMAZ-|[a-z]\d{3}app\d{2}[a-z]{3})[^\s<>,]+', 
                [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            
            # Pattern 3: Compliance status
            $CompliantMatches = [regex]::Matches($Content, '(Compliant|NonCompliant|Not Reachable)', 
                [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            
            # Pattern 4: Agent status
            $AgentMatches = [regex]::Matches($Content, 'Connected to (nessus-manager[^\s<]+)')
            
            # Build server objects from extracted data
            $IPList = $IPMatches | ForEach-Object { $_.Value } | Select-Object -Unique
            
            foreach ($IP in $IPList) {
                # Find surrounding context for this IP
                $IPIndex = $Content.IndexOf($IP)
                $ContextStart = [Math]::Max(0, $IPIndex - 500)
                $ContextEnd = [Math]::Min($Content.Length, $IPIndex + 500)
                $Context = $Content.Substring($ContextStart, $ContextEnd - $ContextStart)
                
                # Extract server name from context
                $ServerName = "Unknown"
                $NameMatch = [regex]::Match($Context, '(?:EC2AMAZ-|[a-z]\d{3}app\d{2}[a-z]{3})[^\s<>,]+', 
                    [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                if ($NameMatch.Success) {
                    $ServerName = $NameMatch.Value
                }
                
                # Extract compliance status
                $TrendMicro = "Unknown"
                $TrellixAgent = "Unknown"
                $CrowdStrike = "Unknown"
                $CloudWatch = "Unknown"
                
                if ($Context -match 'Trend\s*Micro[^\w]*(\w+)') { $TrendMicro = $Matches[1] }
                if ($Context -match 'Trellix\s*Agent[^\w]*(\w+)') { $TrellixAgent = $Matches[1] }
                if ($Context -match 'CrowdStrike[^\w]*(\w+)') { $CrowdStrike = $Matches[1] }
                if ($Context -match 'CloudWatch[^\w]*(\w+)') { $CloudWatch = $Matches[1] }
                
                # Extract Defender signature
                $DefenderSig = "Unknown"
                if ($Context -match '(\d+\.\d+\.\d+)') {
                    $DefenderSig = $Matches[1]
                }
                
                # Extract Nessus status
                $NessusStatus = "Unknown"
                if ($Context -match 'Connected to (nessus-manager[^\s<,]+)') {
                    $NessusStatus = "Connected: $($Matches[1])"
                } elseif ($Context -match 'secs ago,(\d+) of (\d+) limit') {
                    $NessusStatus = "Connected ($($Matches[1]) of $($Matches[2]) limit)"
                }
                
                $AllServersFromReports += [PSCustomObject]@{
                    IP = $IP
                    ServerName = $ServerName
                    SourceReport = $File.Name
                    TrendMicro = $TrendMicro
                    TrellixAgent = $TrellixAgent
                    CrowdStrike = $CrowdStrike
                    CloudWatch = $CloudWatch
                    DefenderSignature = $DefenderSig
                    NessusAgentStatus = $NessusStatus
                    LastChecked = $File.LastWriteTime
                }
            }
            
        } catch {
            Write-Host "    ERROR processing $($File.Name): $_" -ForegroundColor Red
        }
    }
    
    Write-Host "  Extracted $($AllServersFromReports.Count) server entries from reports" -ForegroundColor Green
    return $AllServersFromReports
}

function Compare-PatchStatus {
    param(
        [array]$RickList,
        [array]$ReportData
    )
    
    Write-Host "`nComparing Rick-patch-List against Reports..." -ForegroundColor Cyan
    
    $ComparisonResults = @()
    
    foreach ($Server in $RickList) {
        # Get IP from Rick-patch-List (handle different column names)
        $ServerIP = $Server.'server IP' ?? $Server.'ServerIP' ?? $Server.'IP' ?? $Server.'A'
        $PatchName = $Server.'patch name' ?? $Server.'PatchName' ?? $Server.'B'
        $TicketNumber = $Server.'ticket number' ?? $Server.'TicketNumber' ?? $Server.'D'
        
        if ([string]::IsNullOrWhiteSpace($ServerIP)) {
            continue
        }
        
        # Find matching server in reports
        $MatchedServers = $ReportData | Where-Object { $_.IP -eq $ServerIP }
        
        if ($MatchedServers) {
            foreach ($Match in $MatchedServers) {
                # Determine overall patch status
                $PatchStatus = "PATCHED ✓"
                $StatusColor = "success"
                
                $IssuesList = @()
                
                if ($Match.TrendMicro -match 'NonCompliant|Not') {
                    $IssuesList += "Trend Micro: $($Match.TrendMicro)"
                    $PatchStatus = "MISSING PATCHES ⚠"
                    $StatusColor = "warning"
                }
                
                if ($Match.TrellixAgent -match 'NonCompliant|Not') {
                    $IssuesList += "Trellix: $($Match.TrellixAgent)"
                    $PatchStatus = "MISSING PATCHES ⚠"
                    $StatusColor = "warning"
                }
                
                if ($Match.CrowdStrike -match 'NonCompliant|Not') {
                    $IssuesList += "CrowdStrike: $($Match.CrowdStrike)"
                    $PatchStatus = "MISSING PATCHES ⚠"
                    $StatusColor = "warning"
                }
                
                if ($Match.CloudWatch -match 'NonCompliant|Not') {
                    $IssuesList += "CloudWatch: $($Match.CloudWatch)"
                    $PatchStatus = "MISSING PATCHES ⚠"
                    $StatusColor = "warning"
                }
                
                $ComparisonResults += [PSCustomObject]@{
                    ServerIP = $ServerIP
                    ServerName = $Match.ServerName
                    PatchName = $PatchName
                    TicketNumber = $TicketNumber
                    PatchStatus = $PatchStatus
                    StatusColor = $StatusColor
                    TrendMicro = $Match.TrendMicro
                    TrellixAgent = $Match.TrellixAgent
                    CrowdStrike = $Match.CrowdStrike
                    CloudWatch = $Match.CloudWatch
                    DefenderSignature = $Match.DefenderSignature
                    NessusAgentStatus = $Match.NessusAgentStatus
                    Issues = ($IssuesList -join "; ")
                    SourceReport = $Match.SourceReport
                    LastChecked = $Match.LastChecked
                    FoundInReport = "Yes"
                }
            }
        } else {
            # Server NOT found in any report
            $ComparisonResults += [PSCustomObject]@{
                ServerIP = $ServerIP
                ServerName = "Unknown"
                PatchName = $PatchName
                TicketNumber = $TicketNumber
                PatchStatus = "NOT IN REPORTS ❌"
                StatusColor = "danger"
                TrendMicro = "N/A"
                TrellixAgent = "N/A"
                CrowdStrike = "N/A"
                CloudWatch = "N/A"
                DefenderSignature = "N/A"
                NessusAgentStatus = "N/A"
                Issues = "Server not found in any compliance report"
                SourceReport = "N/A"
                LastChecked = "N/A"
                FoundInReport = "No"
            }
        }
    }
    
    Write-Host "  Comparison complete: $($ComparisonResults.Count) entries" -ForegroundColor Green
    return $ComparisonResults
}

function Generate-HTMLReport {
    param(
        [array]$Data,
        [string]$OutputPath
    )
    
    Write-Host "`nGenerating HTML Report..." -ForegroundColor Cyan
    
    # Calculate statistics
    $TotalServers = $Data.Count
    $FullyPatched = ($Data | Where-Object { $_.PatchStatus -eq "PATCHED ✓" }).Count
    $MissingPatches = ($Data | Where-Object { $_.PatchStatus -eq "MISSING PATCHES ⚠" }).Count
    $NotInReports = ($Data | Where-Object { $_.PatchStatus -eq "NOT IN REPORTS ❌" }).Count
    
    $PatchedPercent = if ($TotalServers -gt 0) { [math]::Round(($FullyPatched / $TotalServers) * 100, 1) } else { 0 }
    
    $HTML = @"
<!DOCTYPE html>
<html>
<head>
    <title>Patch Compliance Report - $(Get-Date -Format 'MMMM dd, yyyy')</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            color: #333;
        }
        
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #2E75B6 0%, #1a4d7a 100%);
            color: white;
            padding: 40px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 36px;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }
        
        .header p {
            font-size: 16px;
            opacity: 0.9;
        }
        
        .stats-container {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            padding: 30px;
            background: #f8f9fa;
        }
        
        .stat-card {
            background: white;
            padding: 25px;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            border-left: 5px solid #2E75B6;
            transition: transform 0.3s ease;
        }
        
        .stat-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 6px 12px rgba(0,0,0,0.15);
        }
        
        .stat-card.success {
            border-left-color: #28a745;
        }
        
        .stat-card.warning {
            border-left-color: #ffc107;
        }
        
        .stat-card.danger {
            border-left-color: #dc3545;
        }
        
        .stat-label {
            font-size: 14px;
            color: #666;
            margin-bottom: 8px;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        
        .stat-value {
            font-size: 42px;
            font-weight: bold;
            color: #2E75B6;
        }
        
        .stat-card.success .stat-value {
            color: #28a745;
        }
        
        .stat-card.warning .stat-value {
            color: #ffc107;
        }
        
        .stat-card.danger .stat-value {
            color: #dc3545;
        }
        
        .content {
            padding: 30px;
        }
        
        .section-title {
            font-size: 24px;
            color: #2E75B6;
            margin: 30px 0 20px 0;
            padding-bottom: 10px;
            border-bottom: 3px solid #2E75B6;
        }
        
        .filter-buttons {
            margin: 20px 0;
            display: flex;
            gap: 10px;
            flex-wrap: wrap;
        }
        
        .filter-btn {
            padding: 10px 20px;
            border: 2px solid #2E75B6;
            background: white;
            color: #2E75B6;
            border-radius: 5px;
            cursor: pointer;
            font-weight: bold;
            transition: all 0.3s ease;
        }
        
        .filter-btn:hover {
            background: #2E75B6;
            color: white;
        }
        
        .filter-btn.active {
            background: #2E75B6;
            color: white;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            background: white;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        
        th {
            background: linear-gradient(135deg, #2E75B6 0%, #1a4d7a 100%);
            color: white;
            padding: 15px;
            text-align: left;
            font-weight: bold;
            text-transform: uppercase;
            font-size: 12px;
            letter-spacing: 1px;
        }
        
        td {
            padding: 12px 15px;
            border-bottom: 1px solid #e0e0e0;
        }
        
        tr:hover {
            background: #f5f5f5;
        }
        
        .status-badge {
            padding: 6px 12px;
            border-radius: 20px;
            font-weight: bold;
            font-size: 12px;
            display: inline-block;
            text-align: center;
            min-width: 150px;
        }
        
        .status-badge.success {
            background: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
        }
        
        .status-badge.warning {
            background: #fff3cd;
            color: #856404;
            border: 1px solid #ffeaa7;
        }
        
        .status-badge.danger {
            background: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
        }
        
        .compliance-badge {
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: bold;
        }
        
        .compliance-badge.compliant {
            background: #28a745;
            color: white;
        }
        
        .compliance-badge.noncompliant {
            background: #dc3545;
            color: white;
        }
        
        .footer {
            background: #f8f9fa;
            padding: 20px;
            text-align: center;
            color: #666;
            font-size: 14px;
        }
        
        .search-box {
            padding: 12px;
            width: 100%;
            max-width: 400px;
            border: 2px solid #2E75B6;
            border-radius: 5px;
            font-size: 14px;
            margin-bottom: 20px;
        }
        
        .progress-bar {
            width: 100%;
            height: 30px;
            background: #e0e0e0;
            border-radius: 15px;
            overflow: hidden;
            margin: 20px 0;
        }
        
        .progress-fill {
            height: 100%;
            background: linear-gradient(90deg, #28a745 0%, #20c997 100%);
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-weight: bold;
            transition: width 1s ease;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🔒 Patch Compliance Report</h1>
            <p>Generated on $(Get-Date -Format 'MMMM dd, yyyy - hh:mm:ss tt')</p>
        </div>
        
        <div class="stats-container">
            <div class="stat-card">
                <div class="stat-label">Total Servers</div>
                <div class="stat-value">$TotalServers</div>
            </div>
            <div class="stat-card success">
                <div class="stat-label">Fully Patched ✓</div>
                <div class="stat-value">$FullyPatched</div>
            </div>
            <div class="stat-card warning">
                <div class="stat-label">Missing Patches ⚠</div>
                <div class="stat-value">$MissingPatches</div>
            </div>
            <div class="stat-card danger">
                <div class="stat-label">Not in Reports ❌</div>
                <div class="stat-value">$NotInReports</div>
            </div>
        </div>
        
        <div class="content">
            <div class="progress-bar">
                <div class="progress-fill" style="width: $PatchedPercent%">
                    $PatchedPercent% Compliant
                </div>
            </div>
            
            <div class="section-title">Detailed Server Status</div>
            
            <input type="text" class="search-box" id="searchBox" placeholder="🔍 Search by IP, Server Name, or Ticket Number..." onkeyup="filterTable()">
            
            <div class="filter-buttons">
                <button class="filter-btn active" onclick="filterStatus('all')">All Servers</button>
                <button class="filter-btn" onclick="filterStatus('PATCHED')">Fully Patched Only</button>
                <button class="filter-btn" onclick="filterStatus('MISSING')">Missing Patches</button>
                <button class="filter-btn" onclick="filterStatus('NOT IN')">Not in Reports</button>
            </div>
            
            <table id="serverTable">
                <thead>
                    <tr>
                        <th>Server IP</th>
                        <th>Server Name</th>
                        <th>Patch Name</th>
                        <th>Ticket #</th>
                        <th>Status</th>
                        <th>Trend Micro</th>
                        <th>Trellix</th>
                        <th>CrowdStrike</th>
                        <th>CloudWatch</th>
                        <th>Defender Sig</th>
                        <th>Nessus Status</th>
                        <th>Issues</th>
                        <th>Source Report</th>
                        <th>Last Checked</th>
                    </tr>
                </thead>
                <tbody>
"@

    foreach ($Row in $Data) {
        $StatusBadge = "<span class='status-badge $($Row.StatusColor)'>$($Row.PatchStatus)</span>"
        
        $TrendMicroBadge = if ($Row.TrendMicro -match 'Compliant') {
            "<span class='compliance-badge compliant'>$($Row.TrendMicro)</span>"
        } elseif ($Row.TrendMicro -match 'NonCompliant') {
            "<span class='compliance-badge noncompliant'>$($Row.TrendMicro)</span>"
        } else {
            $Row.TrendMicro
        }
        
        $TrellixBadge = if ($Row.TrellixAgent -match 'Compliant') {
            "<span class='compliance-badge compliant'>$($Row.TrellixAgent)</span>"
        } elseif ($Row.TrellixAgent -match 'NonCompliant') {
            "<span class='compliance-badge noncompliant'>$($Row.TrellixAgent)</span>"
        } else {
            $Row.TrellixAgent
        }
        
        $CrowdStrikeBadge = if ($Row.CrowdStrike -match 'Compliant') {
            "<span class='compliance-badge compliant'>$($Row.CrowdStrike)</span>"
        } elseif ($Row.CrowdStrike -match 'NonCompliant') {
            "<span class='compliance-badge noncompliant'>$($Row.CrowdStrike)</span>"
        } else {
            $Row.CrowdStrike
        }
        
        $CloudWatchBadge = if ($Row.CloudWatch -match 'Compliant') {
            "<span class='compliance-badge compliant'>$($Row.CloudWatch)</span>"
        } elseif ($Row.CloudWatch -match 'NonCompliant') {
            "<span class='compliance-badge noncompliant'>$($Row.CloudWatch)</span>"
        } else {
            $Row.CloudWatch
        }
        
        $HTML += @"
                    <tr class="data-row" data-status="$($Row.PatchStatus)">
                        <td><strong>$($Row.ServerIP)</strong></td>
                        <td>$($Row.ServerName)</td>
                        <td>$($Row.PatchName)</td>
                        <td>$($Row.TicketNumber)</td>
                        <td>$StatusBadge</td>
                        <td>$TrendMicroBadge</td>
                        <td>$TrellixBadge</td>
                        <td>$CrowdStrikeBadge</td>
                        <td>$CloudWatchBadge</td>
                        <td>$($Row.DefenderSignature)</td>
                        <td>$($Row.NessusAgentStatus)</td>
                        <td>$($Row.Issues)</td>
                        <td>$($Row.SourceReport)</td>
                        <td>$($Row.LastChecked)</td>
                    </tr>
"@
    }

    $HTML += @"
                </tbody>
            </table>
        </div>
        
        <div class="footer">
            <p><strong>Report Generated by:</strong> Patch Compliance Checker v1.0</p>
            <p><strong>Author:</strong> Syed Ahmad | <strong>Date:</strong> $(Get-Date -Format 'MMMM dd, yyyy')</p>
            <p>Report Alert Folder: $ReportAlertFolder</p>
        </div>
    </div>
    
    <script>
        function filterStatus(status) {
            const rows = document.querySelectorAll('.data-row');
            const buttons = document.querySelectorAll('.filter-btn');
            
            buttons.forEach(btn => btn.classList.remove('active'));
            event.target.classList.add('active');
            
            rows.forEach(row => {
                const rowStatus = row.getAttribute('data-status');
                if (status === 'all') {
                    row.style.display = '';
                } else {
                    row.style.display = rowStatus.includes(status) ? '' : 'none';
                }
            });
        }
        
        function filterTable() {
            const input = document.getElementById('searchBox');
            const filter = input.value.toUpperCase();
            const table = document.getElementById('serverTable');
            const rows = table.getElementsByTagName('tr');
            
            for (let i = 1; i < rows.length; i++) {
                const row = rows[i];
                const cells = row.getElementsByTagName('td');
                let found = false;
                
                for (let j = 0; j < cells.length; j++) {
                    const cell = cells[j];
                    if (cell) {
                        const textValue = cell.textContent || cell.innerText;
                        if (textValue.toUpperCase().indexOf(filter) > -1) {
                            found = true;
                            break;
                        }
                    }
                }
                
                row.style.display = found ? '' : 'none';
            }
        }
    </script>
</body>
</html>
"@

    $HTML | Out-File -FilePath $OutputPath -Encoding UTF8
    Write-Host "  HTML Report saved to: $OutputPath" -ForegroundColor Green
    
    return $OutputPath
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  PATCH COMPLIANCE CHECKER" -ForegroundColor Yellow
Write-Host "  Author: Syed Ahmad" -ForegroundColor Gray
Write-Host "  Date: $(Get-Date -Format 'MMMM dd, yyyy')" -ForegroundColor Gray
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Step 1: Verify Report-Alert folder exists
if (-not (Test-Path $ReportAlertFolder)) {
    Write-Host "ERROR: Report-Alert folder not found at: $ReportAlertFolder" -ForegroundColor Red
    Write-Host "Please verify the path and try again." -ForegroundColor Yellow
    exit 1
}

# Step 2: Read Rick-patch-List
$RickPatchListPath = Join-Path $ReportAlertFolder $RickPatchListFile
if (-not (Test-Path $RickPatchListPath)) {
    Write-Host "ERROR: Rick-patch-List.xlsx not found at: $RickPatchListPath" -ForegroundColor Red
    Write-Host "Please ensure the file exists in the Report-Alert folder." -ForegroundColor Yellow
    exit 1
}

$ServerList = Read-RickPatchList -FilePath $RickPatchListPath

if ($ServerList.Count -eq 0) {
    Write-Host "ERROR: No servers found in Rick-patch-List" -ForegroundColor Red
    exit 1
}

# Step 3: Parse all compliance reports
$ReportData = Parse-ComplianceReports -FolderPath $ReportAlertFolder

if ($ReportData.Count -eq 0) {
    Write-Host "WARNING: No server data extracted from reports" -ForegroundColor Yellow
}

# Step 4: Compare and generate results
$ComparisonResults = Compare-PatchStatus -RickList $ServerList -ReportData $ReportData

# Step 5: Generate HTML report
$ReportPath = Generate-HTMLReport -Data $ComparisonResults -OutputPath $OutputReportPath

# Step 6: Display summary
Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host "  SUMMARY" -ForegroundColor Yellow
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Total Servers Checked: $($ComparisonResults.Count)" -ForegroundColor White
Write-Host "Fully Patched: $(($ComparisonResults | Where-Object { $_.PatchStatus -eq 'PATCHED ✓' }).Count)" -ForegroundColor Green
Write-Host "Missing Patches: $(($ComparisonResults | Where-Object { $_.PatchStatus -eq 'MISSING PATCHES ⚠' }).Count)" -ForegroundColor Yellow
Write-Host "Not in Reports: $(($ComparisonResults | Where-Object { $_.PatchStatus -eq 'NOT IN REPORTS ❌' }).Count)" -ForegroundColor Red
Write-Host ""
Write-Host "HTML Report: $ReportPath" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

# Step 7: Open report in browser
Write-Host "`nOpening report in browser..." -ForegroundColor Yellow
Start-Process $ReportPath

Write-Host "`n✓ DONE BRO! Check out your detailed HTML report! 🔥" -ForegroundColor Green

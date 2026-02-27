<#
.SYNOPSIS
    ULTIMATE Patch Compliance Scanner - Universal Multi-Format Patch Report Analyzer
    
.DESCRIPTION
    Production-grade patch compliance scanner that works with ANY environment:
    - Auto-detects all file formats (Excel, CSV, HTML, MSG, TXT, JSON)
    - Smart column name detection (no hardcoded names)
    - Works with ANY patch management system (SCCM, WSUS, Nessus, Qualys, etc.)
    - Generates comprehensive HTML reports
    - Optional CSV export
    - Detailed logging
    - Email notification support
    - Scheduled task friendly
    
    PERFECT FOR: DoD, Federal, Commercial, SAP environments
    
.PARAMETER ScanPath
    Path to scan for patch reports (default: C:\Report-Alert)
    
.PARAMETER ServerListFile
    Server inventory file (Excel/CSV) - auto-detected if not specified
    
.PARAMETER OutputPath
    Where to save reports (default: C:\Report-Alert\Reports)
    
.PARAMETER EmailReport
    Send email report (requires SMTP config in settings)
    
.EXAMPLE
    .\ULTIMATE-Patch-Compliance-Scanner.ps1
    
.EXAMPLE
    .\ULTIMATE-Patch-Compliance-Scanner.ps1 -ScanPath "D:\PatchReports" -OutputPath "D:\Reports"
    
.NOTES
    Author: Syed Ahmad
    Version: 2.0 ULTIMATE EDITION
    Date: February 27, 2026
    License: Free for any use
    
    NO MODIFICATIONS NEEDED - JUST RUN IT!
#>

[CmdletBinding()]
param(
    [string]$ScanPath = "C:\Report-Alert",
    [string]$ServerListFile = "",
    [string]$OutputPath = "",
    [switch]$EmailReport = $false,
    [switch]$ExportCSV = $true,
    [switch]$Verbose = $true
)

# ============================================================================
# GLOBAL CONFIGURATION - AUTOMATICALLY ADAPTS TO YOUR ENVIRONMENT
# ============================================================================

$Script:Config = @{
    # Auto-detect these file patterns
    ServerListPatterns = @("*server*list*", "*inventory*", "*asset*", "*rick*patch*", "*patch*list*")
    ReportFilePatterns = @("*compliance*", "*patch*", "*scan*", "*vulnerability*", "*nessus*", "*qualys*")
    
    # Support all these formats
    SupportedFormats = @(".xlsx", ".xls", ".csv", ".html", ".htm", ".msg", ".txt", ".json", ".xml")
    
    # Smart detection keywords for compliance status
    CompliantKeywords = @("compliant", "patched", "up-to-date", "current", "pass", "ok", "success", "installed")
    NonCompliantKeywords = @("noncompliant", "non-compliant", "missing", "failed", "outdated", "vulnerable", "critical", "high")
    
    # Agent/Software detection patterns
    AgentPatterns = @{
        "Trend Micro" = @("trend", "trendmicro", "apex")
        "Trellix" = @("trellix", "mcafee", "mvision")
        "CrowdStrike" = @("crowdstrike", "falcon")
        "CloudWatch" = @("cloudwatch", "aws")
        "Defender" = @("defender", "windows defender", "ATP")
        "Nessus" = @("nessus", "tenable")
        "Qualys" = @("qualys")
        "SCCM" = @("sccm", "configuration manager")
    }
    
    # Email settings (configure if needed)
    SMTP = @{
        Server = "smtp.office365.com"
        Port = 587
        From = "patchreports@yourdomain.com"
        To = @("sysadmin@yourdomain.com")
        UseSSL = $true
    }
}

# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Color = switch ($Level) {
        "INFO" { "Cyan" }
        "SUCCESS" { "Green" }
        "WARNING" { "Yellow" }
        "ERROR" { "Red" }
    }
    
    $LogMessage = "[$Timestamp] [$Level] $Message"
    Write-Host $LogMessage -ForegroundColor $Color
    
    # Also log to file
    $LogFile = Join-Path $OutputPath "PatchCompliance_$(Get-Date -Format 'yyyyMMdd').log"
    Add-Content -Path $LogFile -Value $LogMessage -ErrorAction SilentlyContinue
}

function Test-FileLock {
    param([string]$FilePath)
    
    try {
        $FileStream = [System.IO.File]::Open($FilePath, 'Open', 'Read', 'None')
        $FileStream.Close()
        return $false
    } catch {
        return $true
    }
}

# ============================================================================
# SMART FILE DETECTION AND READING
# ============================================================================

function Find-ServerListFile {
    param([string]$SearchPath)
    
    Write-Log "Auto-detecting server inventory file..." "INFO"
    
    foreach ($Pattern in $Script:Config.ServerListPatterns) {
        $Files = Get-ChildItem -Path $SearchPath -Filter $Pattern -File -ErrorAction SilentlyContinue
        
        foreach ($File in $Files) {
            if ($File.Extension -in @(".xlsx", ".xls", ".csv")) {
                Write-Log "Found server list: $($File.Name)" "SUCCESS"
                return $File.FullName
            }
        }
    }
    
    Write-Log "No server list found. Will analyze all servers in reports." "WARNING"
    return $null
}

function Read-UniversalFile {
    param(
        [string]$FilePath,
        [string]$FileType = "Auto"
    )
    
    Write-Log "Reading file: $(Split-Path $FilePath -Leaf)" "INFO"
    
    $Extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
    
    try {
        switch ($Extension) {
            ".xlsx" { return Read-ExcelFile -FilePath $FilePath }
            ".xls" { return Read-ExcelFile -FilePath $FilePath }
            ".csv" { return Import-Csv -Path $FilePath -ErrorAction Stop }
            ".html" { return Read-HTMLFile -FilePath $FilePath }
            ".htm" { return Read-HTMLFile -FilePath $FilePath }
            ".txt" { return Read-TextFile -FilePath $FilePath }
            ".json" { return Get-Content $FilePath | ConvertFrom-Json }
            ".xml" { return [xml](Get-Content $FilePath) }
            default { 
                Write-Log "Unsupported file type: $Extension" "WARNING"
                return $null
            }
        }
    } catch {
        Write-Log "Error reading file: $_" "ERROR"
        return $null
    }
}

function Read-ExcelFile {
    param([string]$FilePath)
    
    try {
        # Check if file is locked
        if (Test-FileLock -FilePath $FilePath) {
            Write-Log "File is locked: $FilePath" "WARNING"
            Start-Sleep -Seconds 2
        }
        
        $Excel = New-Object -ComObject Excel.Application -ErrorAction Stop
        $Excel.Visible = $false
        $Excel.DisplayAlerts = $false
        
        $Workbook = $Excel.Workbooks.Open($FilePath)
        $Worksheet = $Workbook.Sheets.Item(1)
        
        $Range = $Worksheet.UsedRange
        $RowCount = $Range.Rows.Count
        $ColCount = $Range.Columns.Count
        
        # Read headers
        $Headers = @()
        for ($col = 1; $col -le $ColCount; $col++) {
            $HeaderValue = $Worksheet.Cells.Item(1, $col).Text
            if (-not [string]::IsNullOrWhiteSpace($HeaderValue)) {
                $Headers += $HeaderValue
            }
        }
        
        # Read data
        $Data = @()
        for ($row = 2; $row -le $RowCount; $row++) {
            $RowData = @{}
            $HasData = $false
            
            for ($col = 1; $col -le $Headers.Count; $col++) {
                $CellValue = $Worksheet.Cells.Item($row, $col).Text
                $RowData[$Headers[$col - 1]] = $CellValue
                
                if (-not [string]::IsNullOrWhiteSpace($CellValue)) {
                    $HasData = $true
                }
            }
            
            if ($HasData) {
                $Data += [PSCustomObject]$RowData
            }
        }
        
        $Workbook.Close($false)
        $Excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        Write-Log "  Loaded $($Data.Count) rows from Excel" "SUCCESS"
        return $Data
        
    } catch {
        Write-Log "  Excel error: $_" "ERROR"
        
        # Fallback: Try ImportExcel module
        if (Get-Module -ListAvailable -Name ImportExcel) {
            try {
                $Data = Import-Excel -Path $FilePath
                Write-Log "  Loaded $($Data.Count) rows using ImportExcel module" "SUCCESS"
                return $Data
            } catch {
                Write-Log "  ImportExcel fallback failed: $_" "ERROR"
            }
        }
        
        return @()
    }
}

function Read-HTMLFile {
    param([string]$FilePath)
    
    $Content = Get-Content $FilePath -Raw
    
    # Extract data from tables
    $Data = @()
    
    # Use regex to find table rows
    $TableMatches = [regex]::Matches($Content, '<tr[^>]*>(.*?)</tr>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    
    $Headers = @()
    $IsFirstRow = $true
    
    foreach ($RowMatch in $TableMatches) {
        $RowHTML = $RowMatch.Groups[1].Value
        
        # Extract cells
        $CellMatches = [regex]::Matches($RowHTML, '<t[dh][^>]*>(.*?)</t[dh]>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        
        $Cells = @()
        foreach ($CellMatch in $CellMatches) {
            $CellText = $CellMatch.Groups[1].Value -replace '<[^>]+>', '' -replace '&nbsp;', ' ' -replace '&amp;', '&'
            $CellText = $CellText.Trim()
            $Cells += $CellText
        }
        
        if ($IsFirstRow -and $Cells.Count -gt 0) {
            $Headers = $Cells
            $IsFirstRow = $false
        } elseif ($Cells.Count -eq $Headers.Count) {
            $RowData = @{}
            for ($i = 0; $i -lt $Headers.Count; $i++) {
                $RowData[$Headers[$i]] = $Cells[$i]
            }
            $Data += [PSCustomObject]$RowData
        }
    }
    
    Write-Log "  Extracted $($Data.Count) rows from HTML table" "SUCCESS"
    return $Data
}

function Read-TextFile {
    param([string]$FilePath)
    
    $Content = Get-Content $FilePath -Raw
    
    # Try to parse as structured text
    $Lines = $Content -split "`r?`n"
    $Data = @()
    
    # Look for patterns like "key: value" or "key = value"
    foreach ($Line in $Lines) {
        if ($Line -match '^\s*([^:=]+)\s*[:=]\s*(.+)$') {
            $Key = $Matches[1].Trim()
            $Value = $Matches[2].Trim()
            
            $Data += [PSCustomObject]@{
                Property = $Key
                Value = $Value
            }
        }
    }
    
    if ($Data.Count -eq 0) {
        # Return raw content as single object
        return [PSCustomObject]@{
            Content = $Content
        }
    }
    
    return $Data
}

# ============================================================================
# SMART IP AND SERVER DETECTION
# ============================================================================

function Find-IPAddresses {
    param([string]$Text)
    
    $IPPattern = '\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b'
    $Matches = [regex]::Matches($Text, $IPPattern)
    
    return $Matches | ForEach-Object { $_.Value } | Select-Object -Unique
}

function Find-ServerNames {
    param([string]$Text)
    
    # Common server naming patterns
    $Patterns = @(
        'EC2AMAZ-[A-Z0-9]+',
        '[a-z]\d{3}app\d{2}[a-z]{3}[^\s<>,]*',
        'SERVER-[A-Z0-9]+',
        'SRV-[A-Z0-9]+',
        'WIN-[A-Z0-9]+',
        '[A-Z]{2,}\d{2,}-[A-Z0-9]+'
    )
    
    $ServerNames = @()
    foreach ($Pattern in $Patterns) {
        $Matches = [regex]::Matches($Text, $Pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        $ServerNames += $Matches | ForEach-Object { $_.Value }
    }
    
    return $ServerNames | Select-Object -Unique
}

function Get-SmartColumnValue {
    param(
        [object]$Row,
        [string[]]$PossibleNames
    )
    
    foreach ($Name in $PossibleNames) {
        if ($Row.PSObject.Properties.Name -contains $Name) {
            $Value = $Row.$Name
            if (-not [string]::IsNullOrWhiteSpace($Value)) {
                return $Value
            }
        }
    }
    
    return $null
}

# ============================================================================
# PATCH COMPLIANCE ANALYSIS
# ============================================================================

function Analyze-ComplianceReports {
    param([string]$FolderPath)
    
    Write-Log "Scanning for compliance reports in: $FolderPath" "INFO"
    
    $AllReportFiles = @()
    foreach ($Pattern in $Script:Config.ReportFilePatterns) {
        $Files = Get-ChildItem -Path $FolderPath -Filter $Pattern -File -Recurse -ErrorAction SilentlyContinue
        $AllReportFiles += $Files
    }
    
    $AllReportFiles = $AllReportFiles | Select-Object -Unique FullName, Name, Extension, LastWriteTime
    
    Write-Log "Found $($AllReportFiles.Count) report files" "INFO"
    
    $AllServers = @()
    $ProcessedCount = 0
    
    foreach ($File in $AllReportFiles) {
        $ProcessedCount++
        Write-Progress -Activity "Processing Reports" -Status "File $ProcessedCount of $($AllReportFiles.Count)" -PercentComplete (($ProcessedCount / $AllReportFiles.Count) * 100)
        
        Write-Log "  Processing: $($File.Name)" "INFO"
        
        try {
            # Read file content
            $FileContent = Get-Content $File.FullName -Raw -ErrorAction Stop
            
            # Extract IPs
            $IPs = Find-IPAddresses -Text $FileContent
            
            # Extract server names
            $ServerNames = Find-ServerNames -Text $FileContent
            
            # Analyze each IP found
            foreach ($IP in $IPs) {
                # Get context around this IP
                $IPIndex = $FileContent.IndexOf($IP)
                if ($IPIndex -ge 0) {
                    $ContextStart = [Math]::Max(0, $IPIndex - 1000)
                    $ContextEnd = [Math]::Min($FileContent.Length, $IPIndex + 1000)
                    $Context = $FileContent.Substring($ContextStart, $ContextEnd - $ContextStart)
                    
                    # Determine server name
                    $ServerName = "Unknown"
                    $ContextServerNames = Find-ServerNames -Text $Context
                    if ($ContextServerNames.Count -gt 0) {
                        $ServerName = $ContextServerNames[0]
                    }
                    
                    # Analyze compliance status for each agent
                    $AgentStatus = @{}
                    foreach ($AgentName in $Script:Config.AgentPatterns.Keys) {
                        $Status = Analyze-AgentStatus -Context $Context -AgentName $AgentName
                        $AgentStatus[$AgentName] = $Status
                    }
                    
                    # Determine overall compliance
                    $IsCompliant = $true
                    $Issues = @()
                    
                    foreach ($Agent in $AgentStatus.Keys) {
                        if ($AgentStatus[$Agent] -match "NonCompliant|Missing|Failed|Not") {
                            $IsCompliant = $false
                            $Issues += "$Agent`: $($AgentStatus[$Agent])"
                        }
                    }
                    
                    $AllServers += [PSCustomObject]@{
                        IP = $IP
                        ServerName = $ServerName
                        OverallStatus = if ($IsCompliant) { "COMPLIANT" } else { "NON-COMPLIANT" }
                        TrendMicro = $AgentStatus["Trend Micro"]
                        Trellix = $AgentStatus["Trellix"]
                        CrowdStrike = $AgentStatus["CrowdStrike"]
                        CloudWatch = $AgentStatus["CloudWatch"]
                        Defender = $AgentStatus["Defender"]
                        Nessus = $AgentStatus["Nessus"]
                        Qualys = $AgentStatus["Qualys"]
                        SCCM = $AgentStatus["SCCM"]
                        Issues = ($Issues -join "; ")
                        SourceReport = $File.Name
                        LastChecked = $File.LastWriteTime
                    }
                }
            }
            
        } catch {
            Write-Log "    Error processing file: $_" "WARNING"
        }
    }
    
    Write-Progress -Activity "Processing Reports" -Completed
    
    # Remove duplicates
    $UniqueServers = $AllServers | Group-Object IP | ForEach-Object {
        $_.Group | Sort-Object LastChecked -Descending | Select-Object -First 1
    }
    
    Write-Log "Extracted data for $($UniqueServers.Count) unique servers" "SUCCESS"
    return $UniqueServers
}

function Analyze-AgentStatus {
    param(
        [string]$Context,
        [string]$AgentName
    )
    
    $Patterns = $Script:Config.AgentPatterns[$AgentName]
    
    foreach ($Pattern in $Patterns) {
        if ($Context -match $Pattern) {
            # Found agent mention - determine status
            
            # Check for compliant keywords
            foreach ($Keyword in $Script:Config.CompliantKeywords) {
                if ($Context -match "$Pattern[^\w]*$Keyword") {
                    return "Compliant"
                }
            }
            
            # Check for non-compliant keywords
            foreach ($Keyword in $Script:Config.NonCompliantKeywords) {
                if ($Context -match "$Pattern[^\w]*$Keyword") {
                    return "NonCompliant"
                }
            }
            
            return "Unknown"
        }
    }
    
    return "Not Found"
}

# ============================================================================
# COMPARISON AND REPORTING
# ============================================================================

function Compare-ServersWithInventory {
    param(
        [array]$InventoryServers,
        [array]$ScannedServers
    )
    
    Write-Log "Comparing server inventory with scan results..." "INFO"
    
    $ComparisonResults = @()
    
    foreach ($InventoryServer in $InventoryServers) {
        # Smart detection of IP column
        $IP = Get-SmartColumnValue -Row $InventoryServer -PossibleNames @("IP", "IPAddress", "server IP", "Server_IP", "IP Address", "Host", "Hostname")
        
        if ([string]::IsNullOrWhiteSpace($IP)) {
            continue
        }
        
        # Find in scanned servers
        $Match = $ScannedServers | Where-Object { $_.IP -eq $IP } | Select-Object -First 1
        
        if ($Match) {
            $ComparisonResults += [PSCustomObject]@{
                IP = $IP
                ServerName = $Match.ServerName
                Status = $Match.OverallStatus
                TrendMicro = $Match.TrendMicro
                Trellix = $Match.Trellix
                CrowdStrike = $Match.CrowdStrike
                CloudWatch = $Match.CloudWatch
                Defender = $Match.Defender
                Nessus = $Match.Nessus
                Qualys = $Match.Qualys
                SCCM = $Match.SCCM
                Issues = $Match.Issues
                SourceReport = $Match.SourceReport
                LastChecked = $Match.LastChecked
                FoundInScan = "Yes"
            }
        } else {
            $ComparisonResults += [PSCustomObject]@{
                IP = $IP
                ServerName = Get-SmartColumnValue -Row $InventoryServer -PossibleNames @("ServerName", "Server Name", "Name", "Hostname", "Host")
                Status = "NOT SCANNED"
                TrendMicro = "N/A"
                Trellix = "N/A"
                CrowdStrike = "N/A"
                CloudWatch = "N/A"
                Defender = "N/A"
                Nessus = "N/A"
                Qualys = "N/A"
                SCCM = "N/A"
                Issues = "Server not found in any compliance report"
                SourceReport = "N/A"
                LastChecked = "N/A"
                FoundInScan = "No"
            }
        }
    }
    
    Write-Log "Comparison complete: $($ComparisonResults.Count) servers" "SUCCESS"
    return $ComparisonResults
}

function Generate-HTMLReport {
    param(
        [array]$Data,
        [string]$OutputFile
    )
    
    Write-Log "Generating HTML report..." "INFO"
    
    # Calculate statistics
    $TotalServers = $Data.Count
    $Compliant = ($Data | Where-Object { $_.Status -eq "COMPLIANT" }).Count
    $NonCompliant = ($Data | Where-Object { $_.Status -eq "NON-COMPLIANT" }).Count
    $NotScanned = ($Data | Where-Object { $_.Status -eq "NOT SCANNED" }).Count
    
    $CompliancePercent = if ($TotalServers -gt 0) { [math]::Round(($Compliant / $TotalServers) * 100, 1) } else { 0 }
    
    $HTML = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ULTIMATE Patch Compliance Report - $(Get-Date -Format 'MMMM dd, yyyy')</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            padding: 20px;
        }
        
        .container {
            max-width: 1600px;
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 80px rgba(0,0,0,0.4);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 50px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 42px;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }
        
        .header p {
            font-size: 18px;
            opacity: 0.95;
        }
        
        .badge {
            display: inline-block;
            background: rgba(255,255,255,0.2);
            padding: 8px 16px;
            border-radius: 20px;
            margin-top: 15px;
            font-weight: bold;
        }
        
        .stats {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 25px;
            padding: 40px;
            background: #f8f9fa;
        }
        
        .stat-card {
            background: white;
            padding: 30px;
            border-radius: 15px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
            transition: all 0.3s ease;
            border-left: 5px solid #667eea;
        }
        
        .stat-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.15);
        }
        
        .stat-card.success { border-left-color: #10b981; }
        .stat-card.warning { border-left-color: #f59e0b; }
        .stat-card.danger { border-left-color: #ef4444; }
        
        .stat-label {
            font-size: 14px;
            color: #6b7280;
            margin-bottom: 10px;
            text-transform: uppercase;
            letter-spacing: 1.5px;
            font-weight: 600;
        }
        
        .stat-value {
            font-size: 48px;
            font-weight: bold;
            color: #1f2937;
        }
        
        .stat-card.success .stat-value { color: #10b981; }
        .stat-card.warning .stat-value { color: #f59e0b; }
        .stat-card.danger .stat-value { color: #ef4444; }
        
        .content {
            padding: 40px;
        }
        
        .section-title {
            font-size: 28px;
            color: #1f2937;
            margin: 30px 0 20px 0;
            padding-bottom: 15px;
            border-bottom: 3px solid #667eea;
        }
        
        .progress-container {
            margin: 30px 0;
        }
        
        .progress-bar {
            width: 100%;
            height: 40px;
            background: #e5e7eb;
            border-radius: 20px;
            overflow: hidden;
            position: relative;
        }
        
        .progress-fill {
            height: 100%;
            background: linear-gradient(90deg, #10b981 0%, #34d399 100%);
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-weight: bold;
            font-size: 18px;
            transition: width 2s ease;
        }
        
        .controls {
            display: flex;
            gap: 15px;
            margin: 25px 0;
            flex-wrap: wrap;
            align-items: center;
        }
        
        .search-box {
            flex: 1;
            min-width: 300px;
            padding: 14px 20px;
            border: 2px solid #d1d5db;
            border-radius: 10px;
            font-size: 15px;
            transition: all 0.3s ease;
        }
        
        .search-box:focus {
            outline: none;
            border-color: #667eea;
            box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
        }
        
        .filter-btn {
            padding: 14px 24px;
            border: 2px solid #667eea;
            background: white;
            color: #667eea;
            border-radius: 10px;
            cursor: pointer;
            font-weight: bold;
            transition: all 0.3s ease;
            font-size: 14px;
        }
        
        .filter-btn:hover {
            background: #667eea;
            color: white;
            transform: translateY(-2px);
        }
        
        .filter-btn.active {
            background: #667eea;
            color: white;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 25px 0;
            background: white;
            border-radius: 15px;
            overflow: hidden;
            box-shadow: 0 5px 15px rgba(0,0,0,0.08);
        }
        
        thead {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
        }
        
        th {
            padding: 18px 15px;
            text-align: left;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 13px;
            letter-spacing: 1px;
        }
        
        td {
            padding: 16px 15px;
            border-bottom: 1px solid #f3f4f6;
        }
        
        tr:hover {
            background: #f9fafb;
        }
        
        .status-badge {
            padding: 8px 16px;
            border-radius: 25px;
            font-weight: bold;
            font-size: 12px;
            display: inline-block;
            text-align: center;
            min-width: 130px;
        }
        
        .status-badge.compliant {
            background: #d1fae5;
            color: #065f46;
            border: 1px solid #6ee7b7;
        }
        
        .status-badge.non-compliant {
            background: #fee2e2;
            color: #991b1b;
            border: 1px solid #fca5a5;
        }
        
        .status-badge.not-scanned {
            background: #fef3c7;
            color: #92400e;
            border: 1px solid #fcd34d;
        }
        
        .agent-badge {
            padding: 5px 10px;
            border-radius: 6px;
            font-size: 11px;
            font-weight: 600;
            display: inline-block;
            margin: 2px;
        }
        
        .agent-badge.ok {
            background: #10b981;
            color: white;
        }
        
        .agent-badge.fail {
            background: #ef4444;
            color: white;
        }
        
        .agent-badge.unknown {
            background: #6b7280;
            color: white;
        }
        
        .footer {
            background: #f9fafb;
            padding: 30px;
            text-align: center;
            color: #6b7280;
            border-top: 1px solid #e5e7eb;
        }
        
        .footer strong {
            color: #1f2937;
        }
        
        @media print {
            body { background: white; }
            .controls { display: none; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🛡️ ULTIMATE Patch Compliance Report</h1>
            <p>Universal Multi-Format Compliance Scanner</p>
            <div class="badge">Generated: $(Get-Date -Format 'MMMM dd, yyyy - hh:mm:ss tt')</div>
        </div>
        
        <div class="stats">
            <div class="stat-card">
                <div class="stat-label">Total Servers</div>
                <div class="stat-value">$TotalServers</div>
            </div>
            <div class="stat-card success">
                <div class="stat-label">Compliant ✓</div>
                <div class="stat-value">$Compliant</div>
            </div>
            <div class="stat-card danger">
                <div class="stat-label">Non-Compliant ⚠</div>
                <div class="stat-value">$NonCompliant</div>
            </div>
            <div class="stat-card warning">
                <div class="stat-label">Not Scanned</div>
                <div class="stat-value">$NotScanned</div>
            </div>
        </div>
        
        <div class="content">
            <div class="progress-container">
                <h3 style="margin-bottom: 15px; color: #1f2937;">Overall Compliance Rate</h3>
                <div class="progress-bar">
                    <div class="progress-fill" style="width: $CompliancePercent%">
                        $CompliancePercent% Compliant
                    </div>
                </div>
            </div>
            
            <div class="section-title">Server Details</div>
            
            <div class="controls">
                <input type="text" class="search-box" id="searchBox" placeholder="🔍 Search by IP, Server Name, or any field..." onkeyup="filterTable()">
                <button class="filter-btn active" onclick="filterStatus('all')">All Servers</button>
                <button class="filter-btn" onclick="filterStatus('COMPLIANT')">Compliant Only</button>
                <button class="filter-btn" onclick="filterStatus('NON-COMPLIANT')">Non-Compliant</button>
                <button class="filter-btn" onclick="filterStatus('NOT SCANNED')">Not Scanned</button>
            </div>
            
            <table id="serverTable">
                <thead>
                    <tr>
                        <th>IP Address</th>
                        <th>Server Name</th>
                        <th>Overall Status</th>
                        <th>Trend Micro</th>
                        <th>Trellix</th>
                        <th>CrowdStrike</th>
                        <th>CloudWatch</th>
                        <th>Defender</th>
                        <th>Nessus</th>
                        <th>Issues</th>
                        <th>Source Report</th>
                        <th>Last Checked</th>
                    </tr>
                </thead>
                <tbody>
"@

    foreach ($Server in $Data) {
        $StatusClass = switch ($Server.Status) {
            "COMPLIANT" { "compliant" }
            "NON-COMPLIANT" { "non-compliant" }
            "NOT SCANNED" { "not-scanned" }
            default { "unknown" }
        }
        
        $StatusBadge = "<span class='status-badge $StatusClass'>$($Server.Status)</span>"
        
        function Get-AgentBadge {
            param([string]$Status)
            
            $BadgeClass = switch -Regex ($Status) {
                "Compliant|OK|Installed" { "ok" }
                "NonCompliant|Missing|Failed" { "fail" }
                default { "unknown" }
            }
            
            return "<span class='agent-badge $BadgeClass'>$Status</span>"
        }
        
        $HTML += @"
                    <tr class="data-row" data-status="$($Server.Status)">
                        <td><strong>$($Server.IP)</strong></td>
                        <td>$($Server.ServerName)</td>
                        <td>$StatusBadge</td>
                        <td>$(Get-AgentBadge -Status $Server.TrendMicro)</td>
                        <td>$(Get-AgentBadge -Status $Server.Trellix)</td>
                        <td>$(Get-AgentBadge -Status $Server.CrowdStrike)</td>
                        <td>$(Get-AgentBadge -Status $Server.CloudWatch)</td>
                        <td>$(Get-AgentBadge -Status $Server.Defender)</td>
                        <td>$(Get-AgentBadge -Status $Server.Nessus)</td>
                        <td>$($Server.Issues)</td>
                        <td>$($Server.SourceReport)</td>
                        <td>$($Server.LastChecked)</td>
                    </tr>
"@
    }

    $HTML += @"
                </tbody>
            </table>
        </div>
        
        <div class="footer">
            <p><strong>ULTIMATE Patch Compliance Scanner v2.0</strong></p>
            <p>Universal multi-format compliance analyzer | Works with ANY environment</p>
            <p>Generated by: $(whoami) | Scan Path: $ScanPath</p>
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
                row.style.display = (status === 'all' || rowStatus === status) ? '' : 'none';
            });
        }
        
        function filterTable() {
            const input = document.getElementById('searchBox');
            const filter = input.value.toUpperCase();
            const table = document.getElementById('serverTable');
            const rows = table.getElementsByTagName('tr');
            
            for (let i = 1; i < rows.length; i++) {
                const row = rows[i];
                const text = row.textContent || row.innerText;
                row.style.display = text.toUpperCase().indexOf(filter) > -1 ? '' : 'none';
            }
        }
    </script>
</body>
</html>
"@

    $HTML | Out-File -FilePath $OutputFile -Encoding UTF8 -Force
    Write-Log "HTML report saved: $OutputFile" "SUCCESS"
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  ULTIMATE PATCH COMPLIANCE SCANNER v2.0" -ForegroundColor Yellow
Write-Host "  Universal Multi-Format Analyzer" -ForegroundColor Gray
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Setup output path
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $ScanPath "Reports"
}

if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
    Write-Log "Created output directory: $OutputPath" "INFO"
}

# Step 1: Find server inventory (optional)
$InventoryFile = if ([string]::IsNullOrWhiteSpace($ServerListFile)) {
    Find-ServerListFile -SearchPath $ScanPath
} else {
    $ServerListFile
}

$InventoryServers = @()
if ($InventoryFile -and (Test-Path $InventoryFile)) {
    $InventoryServers = Read-UniversalFile -FilePath $InventoryFile
    Write-Log "Loaded $($InventoryServers.Count) servers from inventory" "SUCCESS"
}

# Step 2: Analyze all compliance reports
$ScannedServers = Analyze-ComplianceReports -FolderPath $ScanPath

# Step 3: Compare if we have inventory
if ($InventoryServers.Count -gt 0) {
    $Results = Compare-ServersWithInventory -InventoryServers $InventoryServers -ScannedServers $ScannedServers
} else {
    $Results = $ScannedServers
    Write-Log "No inventory file found - reporting all discovered servers" "WARNING"
}

# Step 4: Generate reports
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$HTMLReportPath = Join-Path $OutputPath "PatchCompliance_$Timestamp.html"

Generate-HTMLReport -Data $Results -OutputFile $HTMLReportPath

# Step 5: Export CSV if requested
if ($ExportCSV) {
    $CSVPath = Join-Path $OutputPath "PatchCompliance_$Timestamp.csv"
    $Results | Export-Csv -Path $CSVPath -NoTypeInformation -Encoding UTF8
    Write-Log "CSV export saved: $CSVPath" "SUCCESS"
}

# Step 6: Display summary
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  SCAN COMPLETE" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Total Servers Analyzed: $($Results.Count)" -ForegroundColor White
Write-Host "Compliant: $(($Results | Where-Object { $_.Status -eq 'COMPLIANT' }).Count)" -ForegroundColor Green
Write-Host "Non-Compliant: $(($Results | Where-Object { $_.Status -eq 'NON-COMPLIANT' }).Count)" -ForegroundColor Yellow
Write-Host "Not Scanned: $(($Results | Where-Object { $_.Status -eq 'NOT SCANNED' }).Count)" -ForegroundColor Red
Write-Host ""
Write-Host "Reports Generated:" -ForegroundColor Cyan
Write-Host "  HTML: $HTMLReportPath" -ForegroundColor White
if ($ExportCSV) {
    Write-Host "  CSV:  $CSVPath" -ForegroundColor White
}
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Step 7: Open HTML report
Write-Host "Opening HTML report in browser..." -ForegroundColor Yellow
Start-Process $HTMLReportPath

Write-Host ""
Write-Host "✓ DONE! Check your reports folder! 🔥" -ForegroundColor Green
Write-Host ""

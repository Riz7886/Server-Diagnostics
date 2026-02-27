param(
    [string]$TargetServers = "",
    [string]$DownloadPath = "C:\PatchRepository",
    [switch]$CheckOnly = $false,
    [string[]]$Components = @("Windows", "Chrome", "Edge", "Defender", "CloudWatch", "CrowdStrike", "Nessus")
)

$ErrorActionPreference = "SilentlyContinue"

if (-not (Test-Path $DownloadPath)) {
    New-Item -Path $DownloadPath -ItemType Directory -Force | Out-Null
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Enterprise Patch Management System" -ForegroundColor White
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

function Get-Servers {
    param([string]$Source)
    
    if ([string]::IsNullOrWhiteSpace($Source)) {
        return @($env:COMPUTERNAME)
    }
    
    if (Test-Path $Source) {
        if ($Source -match "\.txt$") {
            return Get-Content $Source | Where-Object { $_ -match '\S' }
        } elseif ($Source -match "\.xlsx$") {
            $Excel = New-Object -ComObject Excel.Application
            $Excel.Visible = $false
            $Excel.DisplayAlerts = $false
            $Workbook = $Excel.Workbooks.Open($Source)
            $Worksheet = $Workbook.Sheets.Item(1)
            $Range = $Worksheet.UsedRange
            
            $List = @()
            for ($i = 1; $i -le $Range.Rows.Count; $i++) {
                $Value = $Worksheet.Cells.Item($i, 1).Text
                if ($Value -match '\S') { $List += $Value }
            }
            
            $Workbook.Close($false)
            $Excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            
            return $List
        }
    }
    
    return $Source -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '\S' }
}

function Get-PatchStatus {
    param([string]$Computer)
    
    try {
        $Session = New-CimSession -ComputerName $Computer -ErrorAction Stop
        
        $Hotfixes = Get-CimInstance -CimSession $Session -ClassName Win32_QuickFixEngineering -ErrorAction Stop |
                    Sort-Object -Property InstalledOn -Descending
        
        $LastPatch = if ($Hotfixes -and $Hotfixes[0].InstalledOn) {
            $Hotfixes[0].InstalledOn.ToString("yyyy-MM-dd")
        } else {
            "Unknown"
        }
        
        Remove-CimSession -CimSession $Session
        
        return @{
            Server = $Computer
            Status = "Online"
            LastPatch = $LastPatch
            PatchCount = $Hotfixes.Count
        }
    } catch {
        return @{
            Server = $Computer
            Status = "Unreachable"
            LastPatch = "N/A"
            PatchCount = 0
        }
    }
}

function Get-AvailableUpdates {
    
    try {
        $UpdateSession = New-Object -ComObject Microsoft.Update.Session
        $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
        $SearchResult = $UpdateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
        
        $Updates = @()
        foreach ($Update in $SearchResult.Updates) {
            $KB = "Unknown"
            if ($Update.KBArticleIDs.Count -gt 0) {
                $KB = "KB" + $Update.KBArticleIDs[0]
            }
            
            $Updates += @{
                Title = $Update.Title
                KB = $KB
                Size = [math]::Round($Update.MaxDownloadSize / 1MB, 2)
                Severity = $Update.MsrcSeverity
                UpdateObject = $Update
            }
        }
        
        return $Updates
    } catch {
        return @()
    }
}

function Download-Update {
    param($UpdateObject, $Path)
    
    try {
        $Session = New-Object -ComObject Microsoft.Update.Session
        $Downloader = $Session.CreateUpdateDownloader()
        $Downloader.Updates = New-Object -ComObject Microsoft.Update.UpdateColl
        $Downloader.Updates.Add($UpdateObject) | Out-Null
        $Result = $Downloader.Download()
        
        return $Result.ResultCode -eq 2
    } catch {
        return $false
    }
}

function Get-ChromeVersion {
    
    try {
        $Response = Invoke-WebRequest -Uri "https://chromereleases.googleblog.com/search/label/Stable%20updates" -UseBasicParsing -TimeoutSec 10
        if ($Response.Content -match "(\d+\.\d+\.\d+\.\d+)") {
            return $Matches[1]
        }
    } catch { }
    
    return "Latest"
}

function Download-Chrome {
    param([string]$Path)
    
    $URL = "https://dl.google.com/chrome/install/googlechromestandaloneenterprise64.msi"
    $Output = Join-Path $Path "ChromeEnterprise.msi"
    
    try {
        Invoke-WebRequest -Uri $URL -OutFile $Output -UseBasicParsing -TimeoutSec 60
        if (Test-Path $Output) {
            return $Output
        }
    } catch { }
    
    return $null
}

function Get-EdgeVersion {
    
    try {
        $URL = "https://edgeupdates.microsoft.com/api/products"
        $Response = Invoke-RestMethod -Uri $URL -TimeoutSec 10
        $Stable = $Response | Where-Object { $_.Product -eq "Stable" } | Select-Object -First 1
        
        if ($Stable -and $Stable.Releases) {
            $Latest = $Stable.Releases | Where-Object { $_.Platform -eq "Windows" -and $_.Architecture -eq "x64" } |
                      Sort-Object -Property ProductVersion -Descending | Select-Object -First 1
            
            if ($Latest) {
                return $Latest.ProductVersion
            }
        }
    } catch { }
    
    return "Latest"
}

function Download-Edge {
    param([string]$Path)
    
    $URL = "https://msedge.sf.dl.delivery.mp.microsoft.com/filestreamingservice/files/latest/MicrosoftEdgeEnterpriseX64.msi"
    $Output = Join-Path $Path "EdgeEnterprise.msi"
    
    try {
        Invoke-WebRequest -Uri $URL -OutFile $Output -UseBasicParsing -TimeoutSec 60
        if (Test-Path $Output) {
            return $Output
        }
    } catch { }
    
    return $null
}

function Update-Defender {
    
    try {
        $MpCmdRun = Join-Path $env:ProgramFiles "Windows Defender\MpCmdRun.exe"
        
        if (Test-Path $MpCmdRun) {
            Start-Process -FilePath $MpCmdRun -ArgumentList "-SignatureUpdate" -Wait -NoNewWindow
            
            $Status = Get-MpComputerStatus
            
            return @{
                Version = $Status.AntivirusSignatureVersion
                Age = $Status.AntivirusSignatureAge
                Engine = $Status.AMEngineVersion
                Updated = $Status.AntivirusSignatureLastUpdated
            }
        }
    } catch { }
    
    return @{
        Version = "N/A"
        Age = "N/A"
        Engine = "N/A"
        Updated = "N/A"
    }
}

function Get-CloudWatchVersion {
    
    try {
        $AgentPath = "C:\Program Files\Amazon\AmazonCloudWatchAgent\amazon-cloudwatch-agent.exe"
        
        if (Test-Path $AgentPath) {
            $VersionInfo = (Get-Item $AgentPath).VersionInfo
            return $VersionInfo.ProductVersion
        }
    } catch { }
    
    return "Not Installed"
}

function Download-CloudWatch {
    param([string]$Path)
    
    $URL = "https://s3.amazonaws.com/amazoncloudwatch-agent/windows/amd64/latest/amazon-cloudwatch-agent.msi"
    $Output = Join-Path $Path "AmazonCloudWatchAgent.msi"
    
    try {
        Invoke-WebRequest -Uri $URL -OutFile $Output -UseBasicParsing -TimeoutSec 120
        if (Test-Path $Output) {
            return $Output
        }
    } catch { }
    
    return $null
}

function Get-CrowdStrikeVersion {
    
    try {
        $ServiceName = "CSFalconService"
        $Service = Get-Service -Name $ServiceName -ErrorAction Stop
        
        if ($Service) {
            $AgentPath = "C:\Program Files\CrowdStrike\CSFalconService.exe"
            
            if (Test-Path $AgentPath) {
                $VersionInfo = (Get-Item $AgentPath).VersionInfo
                return $VersionInfo.FileVersion
            }
            
            return "Installed"
        }
    } catch { }
    
    return "Not Installed"
}

function Download-CrowdStrike {
    param([string]$Path)
    
    $ManualNote = Join-Path $Path "CrowdStrike_MANUAL_DOWNLOAD.txt"
    $Note = @"
CrowdStrike Falcon Sensor Download Instructions

Requires manual download from CrowdStrike Falcon console.

Steps:
1. Log in to CrowdStrike Falcon console
2. Navigate to Host Setup and Management - Sensor Downloads
3. Download Windows Sensor 64-bit installer
4. Place installer in: $Path
5. Rename to: CrowdStrikeFalcon.exe

Contact CrowdStrike support for assistance.
"@
    
    $Note | Out-File -FilePath $ManualNote -Encoding UTF8 -Force
    
    return $null
}

function Get-NessusVersion {
    
    try {
        $AgentPath = "C:\Program Files\Tenable\Nessus Agent\nessusagent.exe"
        
        if (Test-Path $AgentPath) {
            $VersionInfo = (Get-Item $AgentPath).VersionInfo
            return $VersionInfo.ProductVersion
        }
    } catch { }
    
    return "Not Installed"
}

function Download-Nessus {
    param([string]$Path)
    
    $ManualNote = Join-Path $Path "Nessus_MANUAL_DOWNLOAD.txt"
    $Note = @"
Nessus Agent Download Instructions

Requires manual download from Tenable portal.

Steps:
1. Log in to Tenable.io or Tenable.sc
2. Navigate to Settings - Agents - Download Agents
3. Download Nessus Agent for Windows x64
4. Place installer in: $Path
5. Rename to: NessusAgent.msi

Required information:
- Linking key from Tenable console
- Tenable server address

Contact Tenable support for assistance.
"@
    
    $Note | Out-File -FilePath $ManualNote -Encoding UTF8 -Force
    
    return $null
}

$ServerList = Get-Servers -Source $TargetServers

Write-Host "Analyzing $($ServerList.Count) servers..." -ForegroundColor Yellow
Write-Host ""

$ServerData = @()
foreach ($Server in $ServerList) {
    Write-Host "  $Server..." -ForegroundColor Gray -NoNewline
    
    $Data = Get-PatchStatus -Computer $Server
    $ServerData += $Data
    
    if ($Data.Status -eq "Online") {
        Write-Host " Online (Last: $($Data.LastPatch))" -ForegroundColor Green
    } else {
        Write-Host " Unreachable" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "Checking available updates..." -ForegroundColor Yellow
Write-Host ""

$WindowsUpdates = @()
$ChromeVersion = "N/A"
$EdgeVersion = "N/A"
$DefenderInfo = @{}
$CloudWatchVersion = "N/A"
$CrowdStrikeVersion = "N/A"
$NessusVersion = "N/A"

if ($Components -contains "Windows") {
    Write-Host "  Windows updates..." -ForegroundColor Gray
    $WindowsUpdates = Get-AvailableUpdates
    Write-Host "  Found: $($WindowsUpdates.Count) updates" -ForegroundColor Green
}

if ($Components -contains "Chrome") {
    Write-Host "  Chrome Enterprise..." -ForegroundColor Gray
    $ChromeVersion = Get-ChromeVersion
    Write-Host "  Version: $ChromeVersion" -ForegroundColor Green
}

if ($Components -contains "Edge") {
    Write-Host "  Microsoft Edge..." -ForegroundColor Gray
    $EdgeVersion = Get-EdgeVersion
    Write-Host "  Version: $EdgeVersion" -ForegroundColor Green
}

if ($Components -contains "Defender") {
    Write-Host "  Windows Defender..." -ForegroundColor Gray
    $DefenderInfo = Update-Defender
    Write-Host "  Version: $($DefenderInfo.Version)" -ForegroundColor Green
}

if ($Components -contains "CloudWatch") {
    Write-Host "  Amazon CloudWatch Agent..." -ForegroundColor Gray
    $CloudWatchVersion = Get-CloudWatchVersion
    Write-Host "  Current: $CloudWatchVersion" -ForegroundColor Green
}

if ($Components -contains "CrowdStrike") {
    Write-Host "  CrowdStrike Falcon..." -ForegroundColor Gray
    $CrowdStrikeVersion = Get-CrowdStrikeVersion
    Write-Host "  Current: $CrowdStrikeVersion" -ForegroundColor Green
}

if ($Components -contains "Nessus") {
    Write-Host "  Nessus Agent..." -ForegroundColor Gray
    $NessusVersion = Get-NessusVersion
    Write-Host "  Current: $NessusVersion" -ForegroundColor Green
}

Write-Host ""

$Downloaded = @()

if (-not $CheckOnly) {
    Write-Host "Downloading to: $DownloadPath" -ForegroundColor Yellow
    Write-Host ""
    
    if ($Components -contains "Windows") {
        $Priority = $WindowsUpdates | Where-Object { $_.Severity -in @("Critical", "Important") } | Select-Object -First 10
        
        foreach ($Update in $Priority) {
            Write-Host "  $($Update.KB)..." -ForegroundColor Gray -NoNewline
            
            $Success = Download-Update -UpdateObject $Update.UpdateObject -Path $DownloadPath
            
            if ($Success) {
                Write-Host " Downloaded" -ForegroundColor Green
                $Downloaded += $Update.KB
            } else {
                Write-Host " Skipped" -ForegroundColor Yellow
            }
        }
    }
    
    if ($Components -contains "Chrome") {
        Write-Host "  Chrome Enterprise..." -ForegroundColor Gray -NoNewline
        $ChromePath = Download-Chrome -Path $DownloadPath
        
        if ($ChromePath) {
            Write-Host " Downloaded" -ForegroundColor Green
            $Downloaded += "Chrome"
        } else {
            Write-Host " Failed" -ForegroundColor Red
        }
    }
    
    if ($Components -contains "Edge") {
        Write-Host "  Microsoft Edge..." -ForegroundColor Gray -NoNewline
        $EdgePath = Download-Edge -Path $DownloadPath
        
        if ($EdgePath) {
            Write-Host " Downloaded" -ForegroundColor Green
            $Downloaded += "Edge"
        } else {
            Write-Host " Failed" -ForegroundColor Red
        }
    }
    
    if ($Components -contains "CloudWatch") {
        Write-Host "  Amazon CloudWatch Agent..." -ForegroundColor Gray -NoNewline
        $CloudWatchPath = Download-CloudWatch -Path $DownloadPath
        
        if ($CloudWatchPath) {
            Write-Host " Downloaded" -ForegroundColor Green
            $Downloaded += "CloudWatch"
        } else {
            Write-Host " Failed" -ForegroundColor Red
        }
    }
    
    if ($Components -contains "CrowdStrike") {
        Write-Host "  CrowdStrike Falcon..." -ForegroundColor Gray -NoNewline
        Download-CrowdStrike -Path $DownloadPath | Out-Null
        Write-Host " Manual download required" -ForegroundColor Yellow
    }
    
    if ($Components -contains "Nessus") {
        Write-Host "  Nessus Agent..." -ForegroundColor Gray -NoNewline
        Download-Nessus -Path $DownloadPath | Out-Null
        Write-Host " Manual download required" -ForegroundColor Yellow
    }
    
    Write-Host ""
    Write-Host "  Total: $($Downloaded.Count) items downloaded" -ForegroundColor Green
}

Write-Host ""
Write-Host "Generating report..." -ForegroundColor Yellow
Write-Host ""

$Report = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Patch Status Report</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}body{font-family:Arial,sans-serif;background:#f5f5f5;padding:20px}.container{max-width:1600px;margin:0 auto;background:#fff;border-radius:12px;box-shadow:0 4px 20px rgba(0,0,0,.15)}.header{background:linear-gradient(135deg,#667eea,#764ba2);color:#fff;padding:40px;text-align:center;border-radius:12px 12px 0 0}.header h1{font-size:32px;margin-bottom:8px}.header p{font-size:14px;opacity:.95}.content{padding:30px}h2{color:#333;margin:20px 0 15px;font-size:22px}table{width:100%;border-collapse:collapse;margin:20px 0;font-size:13px}thead{background:linear-gradient(135deg,#667eea,#764ba2);color:#fff}th{padding:14px 10px;text-align:left;font-weight:600;font-size:12px;text-transform:uppercase}td{padding:12px 10px;border-bottom:1px solid #eee}tr:hover{background:#f9f9f9}.status-online{color:#10b981;font-weight:600}.status-offline{color:#ef4444;font-weight:600}.footer{background:#fafafa;padding:20px;text-align:center;color:#666;font-size:13px;border-top:1px solid #eee}
</style>
</head>
<body>
<div class="container">
<div class="header">
<h1>Enterprise Patch Status Report</h1>
<p>Generated: $(Get-Date -Format 'MMMM dd, yyyy - hh:mm:ss tt')</p>
</div>
<div class="content">
<h2>Server Status</h2>
<table>
<thead>
<tr><th>Server</th><th>Status</th><th>Last Patch</th><th>Total Patches</th></tr>
</thead>
<tbody>
"@

foreach ($Server in $ServerData) {
    $StatusClass = if ($Server.Status -eq "Online") { "status-online" } else { "status-offline" }
    $Report += "<tr><td>$($Server.Server)</td><td class='$StatusClass'>$($Server.Status)</td><td>$($Server.LastPatch)</td><td>$($Server.PatchCount)</td></tr>"
}

$Report += @"
</tbody>
</table>
<h2>Available Updates</h2>
<table>
<thead>
<tr><th>Component</th><th>Version</th><th>Size MB</th><th>Status</th></tr>
</thead>
<tbody>
"@

if ($Components -contains "Chrome") {
    $Report += "<tr><td>Chrome Enterprise</td><td>$ChromeVersion</td><td>90</td><td>Available</td></tr>"
}

if ($Components -contains "Edge") {
    $Report += "<tr><td>Microsoft Edge</td><td>$EdgeVersion</td><td>150</td><td>Available</td></tr>"
}

if ($Components -contains "Defender") {
    $Report += "<tr><td>Windows Defender</td><td>$($DefenderInfo.Version)</td><td>50</td><td>Available</td></tr>"
}

if ($Components -contains "CloudWatch") {
    $Report += "<tr><td>Amazon CloudWatch</td><td>Latest</td><td>75</td><td>Available</td></tr>"
}

if ($Components -contains "CrowdStrike") {
    $Report += "<tr><td>CrowdStrike Falcon</td><td>Latest</td><td>100</td><td>Manual Required</td></tr>"
}

if ($Components -contains "Nessus") {
    $Report += "<tr><td>Nessus Agent</td><td>Latest</td><td>80</td><td>Manual Required</td></tr>"
}

foreach ($Update in ($WindowsUpdates | Select-Object -First 20)) {
    $Severity = if ($Update.Severity) { $Update.Severity } else { "Moderate" }
    $Report += "<tr><td>Windows Update</td><td>$($Update.KB)</td><td>$($Update.Size)</td><td>$Severity</td></tr>"
}

$Report += @"
</tbody>
</table>
</div>
<div class="footer">
Enterprise Patch Management System | Repository: $DownloadPath
</div>
</div>
</body>
</html>
"@

$ReportFile = Join-Path $DownloadPath "PatchStatus_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
$Report | Out-File -FilePath $ReportFile -Encoding UTF8 -Force

Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Process Complete" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Servers: $($ServerData.Count) | Downloaded: $($Downloaded.Count) items" -ForegroundColor White
Write-Host "Repository: $DownloadPath" -ForegroundColor White
Write-Host "Report: $ReportFile" -ForegroundColor White
Write-Host ""

Start-Process $ReportFile

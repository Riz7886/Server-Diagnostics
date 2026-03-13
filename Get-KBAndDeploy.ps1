#Requires -Version 5.0
<#
.SYNOPSIS
    Get-KBAndDeploy.ps1
    Downloads KB5078752 from Microsoft Update Catalog and copies it
    directly to C:\temp on remote host 10.70.20.181

.NOTES
    Author : Syed Rizvi
    Run As : Administrator on your local machine / jumphost
    Usage  : .\Get-KBAndDeploy.ps1
             .\Get-KBAndDeploy.ps1 -RemoteHost "10.70.20.181"
             .\Get-KBAndDeploy.ps1 -RemoteHost "10.70.20.181" -UseCredential
#>

param(
    [string]$RemoteHost   = '10.70.20.181',
    [string]$LocalSaveDir = 'C:\temp',
    [switch]$UseCredential
)

$ErrorActionPreference = 'Stop'
$KBNumber = 'KB5078752'

# ── LOGGING ────────────────────────────────────────────────────────────
function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "[$(Get-Date -Format 'HH:mm:ss')] [$Level] $Message"
    $color = switch ($Level) {
        'SUCCESS' { 'Green'  }
        'ERROR'   { 'Red'    }
        'WARN'    { 'Yellow' }
        'STEP'    { 'Magenta'}
        default   { 'Cyan'   }
    }
    Write-Host $line -ForegroundColor $color
}

function Write-Step { param([string]$T)
    Write-Log ("─" * 55) -Level 'STEP'
    Write-Log "  $T" -Level 'STEP'
    Write-Log ("─" * 55) -Level 'STEP'
}

# ═══════════════════════════════════════════════════════════
#  STEP 1 — ENSURE LOCAL TEMP FOLDER EXISTS
# ═══════════════════════════════════════════════════════════
Write-Step "Step 1 — Preparing local temp folder"

if (-not (Test-Path $LocalSaveDir)) {
    New-Item -ItemType Directory -Path $LocalSaveDir -Force | Out-Null
    Write-Log "Created: $LocalSaveDir" -Level 'SUCCESS'
}
else {
    Write-Log "Local folder ready: $LocalSaveDir" -Level 'SUCCESS'
}

# ═══════════════════════════════════════════════════════════
#  STEP 2 — GET DOWNLOAD URL FROM MICROSOFT UPDATE CATALOG
# ═══════════════════════════════════════════════════════════
Write-Step "Step 2 — Finding $KBNumber on Microsoft Update Catalog"

try {
    $searchUrl = "https://www.catalog.update.microsoft.com/Search.aspx?q=$KBNumber"
    Write-Log "Searching: $searchUrl" -Level 'INFO'

    # Search the catalog
    $searchPage = Invoke-WebRequest -Uri $searchUrl -UseBasicParsing -SessionVariable session
    Write-Log "Catalog page loaded." -Level 'SUCCESS'

    # Extract update IDs from the page
    $updateIds = [regex]::Matches($searchPage.Content, "goToDetails\('([a-f0-9\-]+)'\)") |
                 ForEach-Object { $_.Groups[1].Value } |
                 Select-Object -Unique

    if ($updateIds.Count -eq 0) {
        throw "No updates found for $KBNumber on Microsoft Update Catalog."
    }

    Write-Log "Found $($updateIds.Count) update package(s) for $KBNumber" -Level 'SUCCESS'

    # Get download links for each update ID
    $downloadLinks = @()
    foreach ($uid in $updateIds) {
        try {
            $dlPage = Invoke-WebRequest `
                -Uri "https://www.catalog.update.microsoft.com/DownloadDialog.aspx" `
                -Method Post `
                -Body "updateIDs=[{`"size`":0,`"languages`":`"`",`"uidInfo`":`"$uid`",`"updateID`":`"$uid`"}]&updateIDsBlockedForImport=&wsusApiPresent=&contentImport=&sku=&serverName=&ssl=&portNumber=&version=" `
                -UseBasicParsing `
                -WebSession $session

            $links = [regex]::Matches($dlPage.Content, 'https://[^"]+\.msu') |
                     ForEach-Object { $_.Value } |
                     Select-Object -Unique

            $downloadLinks += $links
        }
        catch {
            Write-Log "Could not get link for update ID $uid — skipping." -Level 'WARN'
        }
    }

    if ($downloadLinks.Count -eq 0) {
        throw "Could not extract any .msu download links from the catalog."
    }

    # Prefer x64
    $bestLink = $downloadLinks | Where-Object { $_ -like '*x64*' } | Select-Object -First 1
    if (-not $bestLink) {
        $bestLink = $downloadLinks | Select-Object -First 1
    }

    Write-Log "Download URL found: $bestLink" -Level 'SUCCESS'
}
catch {
    Write-Log "Catalog lookup failed: $_" -Level 'ERROR'
    Write-Log "" -Level 'INFO'
    Write-Log "MANUAL FALLBACK: Go to this URL in your browser and download manually:" -Level 'WARN'
    Write-Log "  https://www.catalog.update.microsoft.com/Search.aspx?q=KB5078752" -Level 'WARN'
    Write-Log "  Save the .msu file to: $LocalSaveDir" -Level 'WARN'
    Write-Log "  Then re-run this script with -SkipDownload flag" -Level 'WARN'
    exit 1
}

# ═══════════════════════════════════════════════════════════
#  STEP 3 — DOWNLOAD THE KB FILE
# ═══════════════════════════════════════════════════════════
Write-Step "Step 3 — Downloading $KBNumber"

$fileName  = [System.IO.Path]::GetFileName($bestLink)
$localPath = Join-Path $LocalSaveDir $fileName

if (Test-Path $localPath) {
    Write-Log "File already exists locally: $localPath — skipping download." -Level 'WARN'
}
else {
    Write-Log "Downloading to: $localPath" -Level 'INFO'
    Write-Log "This may take a few minutes depending on your connection..." -Level 'INFO'

    try {
        # Use BITS if available (faster, resumable)
        if (Get-Command Start-BitsTransfer -ErrorAction SilentlyContinue) {
            Write-Log "Using BITS transfer (faster)..." -Level 'INFO'
            Start-BitsTransfer `
                -Source $bestLink `
                -Destination $localPath `
                -DisplayName "Downloading $KBNumber" `
                -ErrorAction Stop
        }
        else {
            # Fall back to WebClient with progress
            Write-Log "Using WebClient transfer..." -Level 'INFO'
            $wc = [System.Net.WebClient]::new()
            $wc.DownloadFile($bestLink, $localPath)
        }

        $sizeMB = [math]::Round((Get-Item $localPath).Length / 1MB, 1)
        Write-Log "Download complete. File size: $sizeMB MB" -Level 'SUCCESS'
        Write-Log "Saved to: $localPath" -Level 'SUCCESS'
    }
    catch {
        throw "Download failed: $_"
    }
}

# ═══════════════════════════════════════════════════════════
#  STEP 4 — COPY FILE TO REMOTE HOST 10.70.20.181
# ═══════════════════════════════════════════════════════════
Write-Step "Step 4 — Copying to remote host $RemoteHost"

$remoteTempUNC = "\\$RemoteHost\c$\temp"
$remoteDest    = Join-Path $remoteTempUNC $fileName

Write-Log "Remote destination: $remoteDest" -Level 'INFO'

# Get credentials if needed
$cred = $null
if ($UseCredential) {
    Write-Log "Credential prompt requested. Enter credentials for $RemoteHost ..." -Level 'INFO'
    $cred = Get-Credential -Message "Enter admin credentials for $RemoteHost"
}

# Try METHOD 1 — Direct UNC copy (works if you have network access + admin share)
Write-Log "Trying Method 1: Direct UNC copy (\\$RemoteHost\c$\temp)..." -Level 'INFO'
$method1Success = $false

try {
    # Map drive with credentials if provided
    if ($cred) {
        New-PSDrive -Name 'RemoteTemp' -PSProvider FileSystem `
            -Root $remoteTempUNC -Credential $cred -ErrorAction Stop | Out-Null
    }

    # Ensure remote temp folder exists
    if (-not (Test-Path $remoteTempUNC)) {
        if ($cred) {
            New-Item -ItemType Directory -Path $remoteTempUNC -Force -ErrorAction Stop | Out-Null
        }
        else {
            New-Item -ItemType Directory -Path $remoteTempUNC -Force -ErrorAction Stop | Out-Null
        }
        Write-Log "Created remote temp folder." -Level 'INFO'
    }

    Write-Log "Copying file..." -Level 'INFO'
    Copy-Item -Path $localPath -Destination $remoteDest -Force -ErrorAction Stop

    $remoteSize = (Get-Item $remoteDest -ErrorAction Stop).Length / 1MB
    Write-Log "Method 1 SUCCESS. Remote file size: $([math]::Round($remoteSize,1)) MB" -Level 'SUCCESS'
    $method1Success = $true

    if (Get-PSDrive -Name 'RemoteTemp' -ErrorAction SilentlyContinue) {
        Remove-PSDrive -Name 'RemoteTemp' -Force -ErrorAction SilentlyContinue
    }
}
catch {
    Write-Log "Method 1 failed: $_" -Level 'WARN'
}

# Try METHOD 2 — PowerShell Remoting (WinRM)
if (-not $method1Success) {
    Write-Log "Trying Method 2: PowerShell Remoting (WinRM)..." -Level 'INFO'
    try {
        $sessionParams = @{
            ComputerName = $RemoteHost
            ErrorAction  = 'Stop'
        }
        if ($cred) { $sessionParams['Credential'] = $cred }

        $psSession = New-PSSession @sessionParams

        # Ensure remote temp folder exists
        Invoke-Command -Session $psSession -ScriptBlock {
            if (-not (Test-Path 'C:\temp')) {
                New-Item -ItemType Directory -Path 'C:\temp' -Force | Out-Null
            }
        }

        # Copy file using the PS session
        Write-Log "Transferring file via PS session..." -Level 'INFO'
        Copy-Item -Path $localPath -Destination 'C:\temp\' -ToSession $psSession -Force -ErrorAction Stop

        # Verify on remote
        $remoteCheck = Invoke-Command -Session $psSession -ScriptBlock {
            param($fn)
            $p = "C:\temp\$fn"
            if (Test-Path $p) { (Get-Item $p).Length }
            else { 0 }
        } -ArgumentList $fileName

        Remove-PSSession $psSession -ErrorAction SilentlyContinue

        if ($remoteCheck -gt 0) {
            $remoteSize = [math]::Round($remoteCheck / 1MB, 1)
            Write-Log "Method 2 SUCCESS. Remote file size: $remoteSize MB" -Level 'SUCCESS'
            $method1Success = $true
        }
        else {
            throw "File copy appeared to succeed but file not found on remote host."
        }
    }
    catch {
        Write-Log "Method 2 failed: $_" -Level 'WARN'
    }
}

# Try METHOD 3 — robocopy (very reliable on domain networks)
if (-not $method1Success) {
    Write-Log "Trying Method 3: robocopy..." -Level 'INFO'
    try {
        $roboDest = "\\$RemoteHost\c$\temp"
        $result = robocopy $LocalSaveDir $roboDest $fileName /R:3 /W:5 /NP /LOG+:$LocalSaveDir\robocopy_log.txt

        if ($LASTEXITCODE -le 1) {
            Write-Log "Method 3 SUCCESS via robocopy." -Level 'SUCCESS'
            $method1Success = $true
        }
        else {
            throw "robocopy exit code: $LASTEXITCODE"
        }
    }
    catch {
        Write-Log "Method 3 failed: $_" -Level 'WARN'
    }
}

# ═══════════════════════════════════════════════════════════
#  FINAL RESULT
# ═══════════════════════════════════════════════════════════
Write-Log ("=" * 55) -Level 'STEP'
if ($method1Success) {
    Write-Log "  SUCCESS — $KBNumber is now on $RemoteHost" -Level 'SUCCESS'
    Write-Log "  Location: C:\temp\$fileName" -Level 'SUCCESS'
    Write-Log "  Tell Mark: file is at C:\temp\$fileName" -Level 'SUCCESS'
    Write-Log ("=" * 55) -Level 'STEP'
    Write-Log "" -Level 'INFO'
    Write-Log "To install the patch on $RemoteHost run this:" -Level 'INFO'
    Write-Log "  wusa.exe C:\temp\$fileName /quiet /norestart" -Level 'INFO'
    Write-Log "Or via PowerShell remoting:" -Level 'INFO'
    Write-Log "  Invoke-Command -ComputerName $RemoteHost -ScriptBlock {" -Level 'INFO'
    Write-Log "    wusa.exe C:\temp\$fileName /quiet /norestart" -Level 'INFO'
    Write-Log "  }" -Level 'INFO'
}
else {
    Write-Log "  ALL 3 METHODS FAILED" -Level 'ERROR'
    Write-Log ("=" * 55) -Level 'STEP'
    Write-Log "" -Level 'INFO'
    Write-Log "MANUAL OPTION — do this instead:" -Level 'WARN'
    Write-Log "  1. File is downloaded to: $localPath" -Level 'WARN'
    Write-Log "  2. Open File Explorer, type in address bar:" -Level 'WARN'
    Write-Log "     \\$RemoteHost\c$\temp" -Level 'WARN'
    Write-Log "  3. Drag and drop the file there" -Level 'WARN'
    Write-Log "" -Level 'INFO'
    Write-Log "OR — Run this script again with -UseCredential flag:" -Level 'WARN'
    Write-Log "  .\Get-KBAndDeploy.ps1 -UseCredential" -Level 'WARN'
}

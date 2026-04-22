# ==============================================================================
#  Get-OTACC-Logs.ps1
#  Pull OpenText Core Archive Connector logs from a SAP server to the
#  current jump server, from a single command. Handles:
#    - Cross-domain creds (IL5 jump -> CRE target)
#    - Path with spaces ("Core Archive Connector")
#    - Live/locked log files (robocopy /B backup mode for otacc.log + http_access.log)
#    - Auto-cleanup of the SMB mount when done
#
#  Usage (from the jump server, run PowerShell as Administrator):
#     .\Get-OTACC-Logs.ps1
#
#  Optional parameters:
#     .\Get-OTACC-Logs.ps1 -Server 10.168.0.32 -Dest "$env:USERPROFILE\Desktop\otacc-logs"
# ==============================================================================

param(
    [string]$Server      = "10.168.0.32",
    [string]$RemotePath  = 'OTC\OpenText\Core Archive Connector\logs',
    [string]$Dest        = "$env:USERPROFILE\Desktop\otacc-logs",
    [string]$MountLetter = "Z"
)

$ErrorActionPreference = "Stop"

function Cleanup {
    cmd /c "net use ${MountLetter}: /delete /y" 2>$null | Out-Null
}

try {
    # ---- 1. Credentials for the CRE domain on the target --------------
    Write-Host "This server is on a different domain from your jump host." -ForegroundColor Yellow
    Write-Host "Enter creds that have READ on \\$Server\e$ (usually a CRE domain account)." -ForegroundColor Yellow
    $cred = Get-Credential -Message "Creds for \\$Server\e$"

    $user = $cred.UserName
    $pwd  = $cred.GetNetworkCredential().Password

    # ---- 2. Map the admin share -------------------------------------------
    Write-Host "`nMapping \\$Server\e$ as ${MountLetter}: ..." -ForegroundColor Cyan
    Cleanup   # in case a stale mount exists
    $mapOut = cmd /c "net use ${MountLetter}: \\$Server\e$ $pwd /user:$user /persistent:no" 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "net use failed: $mapOut"
    }
    Write-Host "  Mounted OK." -ForegroundColor Green

    # ---- 3. Destination ---------------------------------------------------
    New-Item -ItemType Directory -Path $Dest -Force | Out-Null
    Write-Host "Destination: $Dest" -ForegroundColor Cyan

    # ---- 4. Pull the logs ------------------------------------------------
    $remoteFull = "${MountLetter}:\$RemotePath"
    Write-Host "`nCopying from: $remoteFull" -ForegroundColor Cyan

    if (-not (Test-Path -LiteralPath $remoteFull)) {
        throw "Remote path not reachable: $remoteFull  (check firewall / share permissions)"
    }

    # /B  = backup mode, lets us read exclusively-locked active log files
    #       (otacc.log and http_access.log are live-written by the Java service)
    # /E  = copy subdirs including empty
    # /R:1 /W:1 = retry once, wait 1 sec  (fail fast in an interactive session)
    # /MT:8  = 8 threads
    # /NP    = no per-file progress (keeps the console clean)
    # /XJ    = skip junction points
    # /COPY:DAT = data+attributes+timestamps (skip ACLs — they won't map across domains)
    $args = @(
        "`"$remoteFull`"",
        "`"$Dest`"",
        "/E","/B","/R:1","/W:1","/MT:8","/NP","/XJ","/COPY:DAT"
    )

    $rc = (Start-Process -FilePath robocopy -ArgumentList $args -NoNewWindow -Wait -PassThru).ExitCode

    # Robocopy exit codes 0-7 are success (0=nothing, 1=copied, 2=extras, 3=both)
    if ($rc -ge 8) {
        Write-Warning "robocopy returned $rc — partial/failed. See console above."
    } else {
        Write-Host "`n  Copy complete (robocopy code $rc)." -ForegroundColor Green
    }

    # ---- 5. Show what we got ---------------------------------------------
    $total  = (Get-ChildItem -Path $Dest -Recurse -File -ErrorAction SilentlyContinue | Measure-Object Length -Sum)
    Write-Host ("`n  Files:   {0}" -f $total.Count) -ForegroundColor Green
    Write-Host ("  Size:    {0:N1} MB" -f ($total.Sum / 1MB)) -ForegroundColor Green
    Write-Host ("  Path:    {0}" -f $Dest) -ForegroundColor Green
    Write-Host ""
    explorer.exe $Dest
}
catch {
    Write-Host "`n  ERROR: $_" -ForegroundColor Red
    exit 1
}
finally {
    Write-Host "`nUnmapping ${MountLetter}:" -ForegroundColor DarkGray
    Cleanup
}

$ErrorActionPreference = "Continue"

Clear-Host
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  PREREQUISITE TEST - CHECK BEFORE RUNNING MAIN SCRIPT" -ForegroundColor Cyan
Write-Host "  This script tests if multi-server scanning will work" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

$serverList = @(
    "10.133.39.41",
    "10.116.20.98",
    "10.116.52.137",
    "10.174.8.24",
    "10.174.16.13",
    "10.133.7.16",
    "10.133.39.23",
    "10.116.33.62",
    "10.116.21.83",
    "10.116.52.11"
)

Write-Host "Testing connectivity to all servers..." -ForegroundColor Yellow
Write-Host ""

$results = @()

foreach ($server in $serverList) {
    Write-Host "Testing: $server..." -ForegroundColor Yellow -NoNewline
    
    $result = [PSCustomObject]@{
        Server = $server
        Ping = "Failed"
        WinRM = "Failed"
        Remote = "Failed"
        CanScan = "NO"
    }
    
    # Test 1: Ping
    try {
        $ping = Test-Connection -ComputerName $server -Count 1 -Quiet -ErrorAction Stop
        if ($ping) {
            $result.Ping = "Success"
            Write-Host " [PING: OK]" -ForegroundColor Green -NoNewline
        }
    } catch {
        Write-Host " [PING: FAIL]" -ForegroundColor Red -NoNewline
    }
    
    # Test 2: WinRM Port (5985)
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $connect = $tcp.BeginConnect($server, 5985, $null, $null)
        $wait = $connect.AsyncWaitHandle.WaitOne(1000, $false)
        if ($wait) {
            $tcp.EndConnect($connect)
            $result.WinRM = "Success"
            Write-Host " [WinRM: OK]" -ForegroundColor Green -NoNewline
        } else {
            Write-Host " [WinRM: FAIL]" -ForegroundColor Red -NoNewline
        }
        $tcp.Close()
    } catch {
        Write-Host " [WinRM: FAIL]" -ForegroundColor Red -NoNewline
    }
    
    # Test 3: Actual Remote Command
    try {
        $test = Invoke-Command -ComputerName $server -ScriptBlock { $env:COMPUTERNAME } -ErrorAction Stop
        if ($test) {
            $result.Remote = "Success"
            $result.CanScan = "YES"
            Write-Host " [REMOTE: OK]" -ForegroundColor Green
        }
    } catch {
        Write-Host " [REMOTE: FAIL]" -ForegroundColor Red
        $result.Remote = "Failed: $($_.Exception.Message)"
    }
    
    $results += $result
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Yellow
Write-Host "TEST RESULTS" -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Yellow
Write-Host ""

$canScanCount = ($results | Where-Object { $_.CanScan -eq "YES" }).Count
$cannotScanCount = $serverList.Count - $canScanCount

Write-Host "Servers that CAN be scanned remotely: $canScanCount" -ForegroundColor $(if($canScanCount -gt 0){"Green"}else{"Red"})
Write-Host "Servers that CANNOT be scanned: $cannotScanCount" -ForegroundColor $(if($cannotScanCount -gt 0){"Red"}else{"Green"})
Write-Host ""

Write-Host "DETAILED RESULTS:" -ForegroundColor Cyan
Write-Host ""
Write-Host "Server           Ping    WinRM   Remote  Can Scan?" -ForegroundColor White
Write-Host "------------------------------------------------------------" -ForegroundColor Gray

foreach ($r in $results) {
    $pingColor = if($r.Ping -eq "Success"){"Green"}else{"Red"}
    $winrmColor = if($r.WinRM -eq "Success"){"Green"}else{"Red"}
    $remoteColor = if($r.Remote -eq "Success"){"Green"}else{"Red"}
    $scanColor = if($r.CanScan -eq "YES"){"Green"}else{"Red"}
    
    Write-Host "$($r.Server.PadRight(16))" -NoNewline
    Write-Host " $($r.Ping.PadRight(7))" -ForegroundColor $pingColor -NoNewline
    Write-Host " $($r.WinRM.PadRight(7))" -ForegroundColor $winrmColor -NoNewline
    Write-Host " $($r.Remote.PadRight(7))" -ForegroundColor $remoteColor -NoNewline
    Write-Host " $($r.CanScan)" -ForegroundColor $scanColor
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Yellow
Write-Host "RECOMMENDATION" -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Yellow
Write-Host ""

if ($canScanCount -eq $serverList.Count) {
    Write-Host "ALL SERVERS ARE READY FOR REMOTE SCANNING!" -ForegroundColor Green
    Write-Host ""
    Write-Host "You can safely use: MultiServer-Agent-Report-READONLY.ps1" -ForegroundColor Green
    Write-Host ""
}
elseif ($canScanCount -gt 0) {
    Write-Host "PARTIAL SUCCESS: Some servers can be scanned remotely" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Option 1: Use MultiServer script for the $canScanCount working servers" -ForegroundColor Yellow
    Write-Host "Option 2: Use Single-Server script on the $cannotScanCount failed servers" -ForegroundColor Yellow
    Write-Host ""
}
else {
    Write-Host "REMOTE SCANNING NOT AVAILABLE" -ForegroundColor Red
    Write-Host ""
    Write-Host "REASONS:" -ForegroundColor Red
    Write-Host "  - PowerShell Remoting (WinRM) is not enabled" -ForegroundColor Red
    Write-Host "  - Firewall is blocking remote connections" -ForegroundColor Red
    Write-Host "  - You don't have admin credentials" -ForegroundColor Red
    Write-Host ""
    Write-Host "SOLUTION:" -ForegroundColor Yellow
    Write-Host "  Use the Single-Server script (Agent-Status-Report-READONLY.ps1)" -ForegroundColor Yellow
    Write-Host "  Run it ON each server individually" -ForegroundColor Yellow
    Write-Host ""
}

Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

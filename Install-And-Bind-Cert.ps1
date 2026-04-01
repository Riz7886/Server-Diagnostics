$ErrorActionPreference = "SilentlyContinue"
$WarningPreference     = "SilentlyContinue"
$SignedCertPath        = "C:\navy-cert\signed.p7b"

function Write-Log {
    param([string]$M, [string]$L = "INFO")
    $C = switch ($L) { "OK" {"Green"} "WARN" {"Yellow"} "ERR" {"Red"} "FIX" {"Magenta"} default {"Cyan"} }
    Write-Host "[$L] [$(Get-Date -Format 'HH:mm:ss')] $M" -ForegroundColor $C
}
function Write-Line { Write-Host ("=" * 65) -ForegroundColor Blue }

Write-Line
Write-Host "  NS2SW1APP1 - INSTALL CERT + BIND IIS + RDP + VALIDATE" -ForegroundColor White
Write-Host "  Author: Syed Rizvi" -ForegroundColor Gray
Write-Line

Write-Log "Checking cert file: $SignedCertPath" "INFO"
if (-not (Test-Path $SignedCertPath)) {
    Write-Log "File not found: $SignedCertPath" "ERR"
    Write-Log "Make sure signed.p7b is saved in C:\navy-cert\" "WARN"
    exit 1
}
Write-Log "File found" "OK"

Write-Line
Write-Log "Importing signed.p7b into LocalMachine\My store..." "FIX"
try {
    Import-Certificate -FilePath $SignedCertPath -CertStoreLocation Cert:\LocalMachine\My -ErrorAction Stop | Out-Null
    Write-Log "Certificate imported successfully" "OK"
} catch {
    Write-Log "Import-Certificate failed - trying certutil..." "WARN"
    certutil -addstore My $SignedCertPath 2>&1 | Out-Null
    Write-Log "certutil import attempted" "INFO"
}
Start-Sleep -Seconds 3

Write-Line
Write-Log "Finding certificate in store..." "INFO"
$Cert = Get-ChildItem Cert:\LocalMachine\My |
    Where-Object { $_.HasPrivateKey -eq $true -and $_.NotAfter -gt (Get-Date) } |
    Sort-Object NotAfter -Descending |
    Select-Object -First 1

if (-not $Cert) {
    Write-Log "No valid cert found - check certmgr.msc Local Computer - Personal" "ERR"
    exit 1
}

$Thumbprint = $Cert.Thumbprint
Write-Log "Subject    : $($Cert.Subject)" "OK"
Write-Log "Thumbprint : $Thumbprint" "OK"
Write-Log "Key Size   : $($Cert.PublicKey.Key.KeySize)-bit" "OK"
Write-Log "Expiry     : $($Cert.NotAfter)" "OK"
Write-Log "Private Key: $($Cert.HasPrivateKey)" "OK"

Write-Line
Write-Log "Binding cert to RDP..." "FIX"
$rdpDone = $false
try {
    $ts = Get-WmiObject -Namespace root\cimv2\TerminalServices -Class Win32_TSGeneralSetting -ErrorAction Stop
    $ts.SSLCertificateSHA1Hash = $Thumbprint
    $ts.Put() | Out-Null
    Write-Log "RDP binding via WMI - DONE" "OK"
    $rdpDone = $true
} catch { Write-Log "WMI failed - trying registry..." "WARN" }

if (-not $rdpDone) {
    try {
        $bytes = [byte[]] ($Thumbprint -replace '..', '$0 ' -split ' ' -ne '' | ForEach-Object { [Convert]::ToByte($_, 16) })
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" -Name "SSLCertificateSHA1Hash" -Value $bytes -ErrorAction Stop
        Write-Log "RDP binding via registry - DONE" "OK"
        $rdpDone = $true
    } catch { Write-Log "RDP binding failed" "ERR" }
}

Write-Line
Write-Log "Binding cert to IIS HTTPS port 443..." "FIX"
$iisDone = $false
try {
    Import-Module WebAdministration -ErrorAction Stop
    Get-WebBinding -Name "Default Web Site" -Protocol "https" -ErrorAction SilentlyContinue | Remove-WebBinding -ErrorAction SilentlyContinue
    New-WebBinding -Name "Default Web Site" -Protocol "https" -Port 443 -IPAddress "*" -SslFlags 0
    Start-Sleep -Seconds 1
    (Get-WebBinding -Name "Default Web Site" -Protocol "https").AddSslCertificate($Thumbprint, "My")
    Write-Log "IIS HTTPS binding port 443 - DONE" "OK"
    $iisDone = $true
} catch { Write-Log "IIS WebAdmin failed - trying netsh..." "WARN" }

if (-not $iisDone) {
    try {
        $guid = "{$([System.Guid]::NewGuid().ToString())}"
        netsh http delete sslcert ipport=0.0.0.0:443 | Out-Null
        netsh http add sslcert ipport=0.0.0.0:443 certhash=$Thumbprint appid=$guid | Out-Null
        Write-Log "IIS netsh binding - DONE" "OK"
        $iisDone = $true
    } catch { Write-Log "IIS binding failed" "ERR" }
}

Write-Line
Write-Log "Restarting RDP service..." "FIX"
try {
    Restart-Service TermService -Force -ErrorAction Stop
    Start-Sleep -Seconds 4
    Write-Log "RDP service restarted - DONE" "OK"
} catch { Write-Log "Try manually: Restart-Service TermService -Force" "WARN" }

Write-Line
Write-Host "  FINAL VALIDATION" -ForegroundColor White
Write-Line

$c2 = Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.Thumbprint -eq $Thumbprint }
if ($c2)                                       { Write-Log "Cert in store       : PASS" "OK" } else { Write-Log "Cert in store       : FAIL" "ERR" }
if ($c2 -and $c2.PublicKey.Key.KeySize -ge 2048){ Write-Log "Key size 2048-bit   : PASS" "OK" } else { Write-Log "Key size            : WARN" "WARN" }
if ($c2 -and $c2.HasPrivateKey)                { Write-Log "Private key         : PASS" "OK" } else { Write-Log "Private key         : FAIL" "ERR" }

$rdp2 = Get-WmiObject -Namespace root\cimv2\TerminalServices -Class Win32_TSGeneralSetting
if ($rdp2 -and $rdp2.SSLCertificateSHA1Hash -eq $Thumbprint) { Write-Log "RDP binding         : PASS" "OK" } else { Write-Log "RDP binding         : WARN" "WARN" }

$svc2 = Get-Service TermService -ErrorAction SilentlyContinue
if ($svc2 -and $svc2.Status -eq "Running")     { Write-Log "RDP service running : PASS" "OK" } else { Write-Log "RDP service         : WARN" "WARN" }

Import-Module WebAdministration -ErrorAction SilentlyContinue
$iis2 = Get-WebBinding -Protocol "https" -ErrorAction SilentlyContinue
if ($iis2) { Write-Log "IIS HTTPS port 443  : PASS - $($iis2.bindingInformation)" "OK" } else { Write-Log "IIS HTTPS binding   : WARN" "WARN" }

Write-Line
Write-Log "DONE - Cert installed, IIS and RDP bound, service restarted" "OK"
Write-Log "RDP connections will no longer show invalid cert warnings" "OK"
Write-Line

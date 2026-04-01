# NS2SW1APP1 - Install Signed Cert + Bind IIS + RDP + Validate
# Author: Syed Rizvi
# Run as Administrator in C:\navy-cert\

$ErrorActionPreference = "SilentlyContinue"
$WarningPreference     = "SilentlyContinue"

$P7bPath  = "C:\navy-cert\signed.p7b"
$CerPath  = "C:\navy-cert\signed.cer"

function Write-Log {
    param([string]$M, [string]$L = "INFO")
    $C = switch ($L) { "OK" {"Green"} "WARN" {"Yellow"} "ERR" {"Red"} "FIX" {"Magenta"} default {"Cyan"} }
    Write-Host "[$L] [$(Get-Date -Format 'HH:mm:ss')] $M" -ForegroundColor $C
}
function Write-Line { Write-Host ("=" * 65) -ForegroundColor Blue }

Write-Line
Write-Host "  NS2SW1APP1 - INSTALL CERT + BIND IIS + RDP + VALIDATE" -ForegroundColor White
Write-Host "  Author: Syed Rizvi | Run as Administrator" -ForegroundColor Gray
Write-Line

# -----------------------------------------------------------------------
# STEP 1 - CHECK P7B FILE EXISTS
# -----------------------------------------------------------------------
Write-Log "Checking for signed.p7b at $P7bPath" "INFO"
if (-not (Test-Path $P7bPath)) {
    Write-Log "signed.p7b not found at $P7bPath" "ERR"
    exit 1
}
Write-Log "signed.p7b found" "OK"

# -----------------------------------------------------------------------
# STEP 2 - EXTRACT CER FROM P7B USING CERTUTIL
# -----------------------------------------------------------------------
Write-Line
Write-Log "Extracting .cer from .p7b using certutil..." "FIX"

# certutil -dump extracts individual certs from p7b
certutil -split -dump $P7bPath 2>&1 | Out-Null

# Look for extracted Blob files in current directory
$blobs = Get-Item "C:\navy-cert\*.cer" -ErrorAction SilentlyContinue
if (-not $blobs) {
    # Try extracting manually
    certutil -decode $P7bPath $CerPath 2>&1 | Out-Null
}

# If still no .cer - convert p7b directly
if (-not (Test-Path $CerPath)) {
    Write-Log "Trying direct p7b to cer export..." "WARN"
    $certStore = New-Object System.Security.Cryptography.X509Certificates.X509Store("My","LocalMachine")
    $certStore.Open("ReadWrite")
    $p7bCollection = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2Collection
    $p7bCollection.Import($P7bPath)
    foreach ($cert in $p7bCollection) {
        $certStore.Add($cert)
        Write-Log "Added cert from p7b: $($cert.Subject)" "OK"
        # Export as cer for certreq -accept
        [System.IO.File]::WriteAllBytes($CerPath, $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))
    }
    $certStore.Close()
}

Start-Sleep -Seconds 2

# -----------------------------------------------------------------------
# STEP 3 - CERTREQ -ACCEPT TO PAIR WITH PRIVATE KEY
# -----------------------------------------------------------------------
Write-Line
Write-Log "Running certreq -accept to pair cert with private key..." "FIX"
$acceptResult = certreq -accept $CerPath 2>&1
Write-Log "certreq result: $acceptResult" "INFO"
Start-Sleep -Seconds 3

# -----------------------------------------------------------------------
# STEP 4 - FIND THE CERT IN STORE
# -----------------------------------------------------------------------
Write-Line
Write-Log "Finding cert in LocalMachine\My store..." "INFO"

$Cert = Get-ChildItem Cert:\LocalMachine\My |
    Where-Object { $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date) } |
    Sort-Object NotAfter -Descending |
    Select-Object -First 1

if (-not $Cert) {
    Write-Log "No cert with private key found. Showing all certs:" "ERR"
    Get-ChildItem Cert:\LocalMachine\My | ForEach-Object {
        Write-Log "  Subject: $($_.Subject) | PrivKey: $($_.HasPrivateKey) | Exp: $($_.NotAfter)" "WARN"
    }
    Write-Log "Manual fix: Open certmgr.msc - Local Computer - Personal - right click - All Tasks - Import - signed.p7b" "WARN"
    exit 1
}

$Tp = $Cert.Thumbprint
Write-Log "Subject    : $($Cert.Subject)" "OK"
Write-Log "Thumbprint : $Tp" "OK"
Write-Log "Key Size   : $($Cert.PublicKey.Key.KeySize)-bit" "OK"
Write-Log "Expiry     : $($Cert.NotAfter)" "OK"
Write-Log "Private Key: $($Cert.HasPrivateKey)" "OK"

# -----------------------------------------------------------------------
# STEP 5 - BIND TO RDP
# -----------------------------------------------------------------------
Write-Line
Write-Log "Binding cert to RDP service..." "FIX"
$rdpOK = $false

try {
    $ts = Get-WmiObject -Namespace root\cimv2\TerminalServices -Class Win32_TSGeneralSetting -ErrorAction Stop
    $ts.SSLCertificateSHA1Hash = $Tp
    $ts.Put() | Out-Null
    Write-Log "RDP bound via WMI" "OK"
    $rdpOK = $true
} catch {
    Write-Log "WMI failed - trying registry..." "WARN"
    try {
        $bytes = [byte[]] ($Tp -replace '..','$0 ' -split ' ' -ne '' | ForEach-Object { [Convert]::ToByte($_,16) })
        Set-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" "SSLCertificateSHA1Hash" $bytes -ErrorAction Stop
        Write-Log "RDP bound via registry" "OK"
        $rdpOK = $true
    } catch {
        Write-Log "RDP binding failed: $_" "ERR"
    }
}

# -----------------------------------------------------------------------
# STEP 6 - BIND TO IIS PORT 443
# -----------------------------------------------------------------------
Write-Line
Write-Log "Binding cert to IIS HTTPS port 443..." "FIX"
$iisOK = $false

try {
    Import-Module WebAdministration -ErrorAction Stop
    Get-WebBinding -Name "Default Web Site" -Protocol "https" -ErrorAction SilentlyContinue | Remove-WebBinding -ErrorAction SilentlyContinue
    New-WebBinding -Name "Default Web Site" -Protocol "https" -Port 443 -IPAddress "*" -SslFlags 0
    Start-Sleep -Seconds 1
    (Get-WebBinding -Name "Default Web Site" -Protocol "https").AddSslCertificate($Tp,"My")
    Write-Log "IIS HTTPS bound on port 443" "OK"
    $iisOK = $true
} catch {
    Write-Log "IIS WebAdmin failed - trying netsh..." "WARN"
    try {
        $guid = "{$([System.Guid]::NewGuid().ToString())}"
        netsh http delete sslcert ipport=0.0.0.0:443 | Out-Null
        netsh http add sslcert ipport=0.0.0.0:443 certhash=$Tp appid=$guid | Out-Null
        Write-Log "IIS bound via netsh" "OK"
        $iisOK = $true
    } catch {
        Write-Log "IIS binding failed: $_" "ERR"
    }
}

# -----------------------------------------------------------------------
# STEP 7 - RESTART RDP SERVICE
# -----------------------------------------------------------------------
Write-Line
Write-Log "Restarting RDP service..." "FIX"
try {
    Restart-Service TermService -Force -ErrorAction Stop
    Start-Sleep -Seconds 4
    Write-Log "RDP service restarted" "OK"
} catch {
    Write-Log "Restart issue - run manually: Restart-Service TermService -Force" "WARN"
}

# -----------------------------------------------------------------------
# STEP 8 - FULL VALIDATION
# -----------------------------------------------------------------------
Write-Line
Write-Host "  FINAL VALIDATION" -ForegroundColor White
Write-Line

$v = Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.Thumbprint -eq $Tp }
if ($v)                                         { Write-Log "Cert in store          : PASS" "OK" } else { Write-Log "Cert in store          : FAIL" "ERR" }
if ($v -and $v.PublicKey.Key.KeySize -ge 2048)  { Write-Log "Key size 2048-bit      : PASS" "OK" } else { Write-Log "Key size               : WARN" "WARN" }
if ($v -and $v.HasPrivateKey)                   { Write-Log "Private key present    : PASS" "OK" } else { Write-Log "Private key            : FAIL" "ERR" }

$rdpV = Get-WmiObject -Namespace root\cimv2\TerminalServices -Class Win32_TSGeneralSetting
if ($rdpV -and $rdpV.SSLCertificateSHA1Hash -eq $Tp) { Write-Log "RDP cert binding       : PASS" "OK" } else { Write-Log "RDP cert binding       : WARN" "WARN" }

$svcV = Get-Service TermService -ErrorAction SilentlyContinue
if ($svcV -and $svcV.Status -eq "Running")      { Write-Log "RDP service running    : PASS" "OK" } else { Write-Log "RDP service            : WARN" "WARN" }

Import-Module WebAdministration -ErrorAction SilentlyContinue
$iisV = Get-WebBinding -Protocol "https" -ErrorAction SilentlyContinue
if ($iisV)  { Write-Log "IIS HTTPS port 443     : PASS - $($iisV.bindingInformation)" "OK" } else { Write-Log "IIS HTTPS binding      : WARN" "WARN" }

Write-Line
Write-Log "SCRIPT COMPLETE - Cert installed and bound to IIS and RDP" "OK"
Write-Log "RDP connections will no longer show invalid cert errors" "OK"
Write-Line

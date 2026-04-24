$ErrorActionPreference = 'Continue'

$CertPath      = 'C:\temp\script\OTP.cer'
$ProdKeystore  = '\\10.168.0.32\e$\CERTS\s4pceotcac.p12'
$LocalKeystore = 'C:\temp\script\renewed.p12'
$Alias         = 's4hpce'
$Password      = 's4hpce'
$TargetServer  = '10.168.0.32'

function Say ($m) { Write-Host ""; Write-Host "[*]  $m" -ForegroundColor Cyan }
function Ok  ($m) { Write-Host "[ok] $m" -ForegroundColor Green }
function Warn($m) { Write-Host "[!!] $m" -ForegroundColor Yellow }
function Die ($m) { Write-Host "[xx] $m" -ForegroundColor Red; exit 1 }

Say "Preflight"
if (-not (Test-Path -LiteralPath $CertPath))     { Die "Cert not found: $CertPath" }
if (-not (Test-Path -LiteralPath $ProdKeystore)) { Die "Prod keystore not reachable: $ProdKeystore" }
$certItem  = Get-Item -LiteralPath $CertPath
$ksItem    = Get-Item -LiteralPath $ProdKeystore
Ok "Cert        : $CertPath ($($certItem.Length) bytes, $($certItem.LastWriteTime))"
Ok "Prod keystore: $ProdKeystore ($($ksItem.Length) bytes, $($ksItem.LastWriteTime))"

Say "Examine cert"
try {
    $certObj = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $CertPath
    Write-Host "  Subject   : $($certObj.Subject)"
    Write-Host "  Issuer    : $($certObj.Issuer)"
    Write-Host "  NotBefore : $($certObj.NotBefore)"
    Write-Host "  NotAfter  : $($certObj.NotAfter)"
    Write-Host "  Thumbprint: $($certObj.Thumbprint)"
    $now = Get-Date
    if ($now -gt $certObj.NotAfter)  { Die "Cert already expired on $($certObj.NotAfter)" }
    if ($now -lt $certObj.NotBefore) { Die "Cert not yet valid (starts $($certObj.NotBefore))" }
    Ok "Cert valid for $([int](($certObj.NotAfter - $now).TotalDays)) days"
} catch { Warn "Cert parse failed: $_" }

Say "Find keytool"
$ktCandidates = @(
    'C:\Program Files\Java\*\bin\keytool.exe',
    'C:\Program Files (x86)\Java\*\bin\keytool.exe',
    'C:\Program Files\OpenJDK\*\bin\keytool.exe',
    'C:\Program Files\Eclipse Adoptium\*\bin\keytool.exe',
    'C:\Program Files\Amazon Corretto\*\bin\keytool.exe',
    '\\10.168.0.32\e$\OTC\OpenText\Core Archive Connector\java\bin\keytool.exe'
)
$kt = $null
foreach ($p in $ktCandidates) {
    $f = Get-ChildItem -Path $p -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($f) { $kt = $f.FullName; break }
}
if (-not $kt) { $envKt = Get-Command keytool.exe -ErrorAction SilentlyContinue; if ($envKt) { $kt = $envKt.Path } }
if (-not $kt) { Die "keytool.exe not found" }
Ok "keytool: $kt"

Say "Pull prod keystore to local for safe import test"
if (Test-Path -LiteralPath $LocalKeystore) { Remove-Item -LiteralPath $LocalKeystore -Force }
Copy-Item -LiteralPath $ProdKeystore -Destination $LocalKeystore -Force
Ok "Local copy: $LocalKeystore"

Say "Current prod keystore aliases"
$listBefore = & $kt -list -keystore $LocalKeystore -storetype PKCS12 -storepass $Password 2>&1
$listBefore | ForEach-Object { Write-Host "  $_" }

Say "Import new cert into local copy (alias: $Alias)"
$impArgs = @('-importcert','-trustcacerts','-noprompt','-alias',$Alias,'-file',$CertPath,'-keystore',$LocalKeystore,'-storetype','PKCS12','-storepass',$Password)
$impOut = & $kt @impArgs 2>&1
$impExit = $LASTEXITCODE

if ($impExit -ne 0) {
    Write-Host ""
    Write-Host "==== keytool output ====" -ForegroundColor Red
    $impOut | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    Write-Host "========================" -ForegroundColor Red
    Write-Host ""
    $txt = ($impOut | Out-String)
    if ($txt -match 'Failed to establish chain') {
        Warn "CHAIN ERROR - the intermediate CA is not in JDK cacerts."
        if ($certObj) { Warn "Cert issuer: $($certObj.Issuer)" }
        Warn "Ask client for the chain: either a .p7b file or the intermediate CA cert."
    } elseif ($txt -match 'Public keys in reply and keystore don''t match') {
        Warn "PUBLIC KEY MISMATCH - this cert was signed against a DIFFERENT CSR/keystore."
        Warn "The CSR you sent to the client came from a different keystore than E:\CERTS\s4pceotcac.p12."
        Warn "Look for a local keystore (s4pceotcac_new.p12 or similar) from Phase 1."
    }
    Die "Import failed. Prod is UNTOUCHED (only local copy was modified)."
}
Ok "Import succeeded on local copy"

Say "Verify imported cert"
& $kt -list -v -alias $Alias -keystore $LocalKeystore -storetype PKCS12 -storepass $Password 2>&1 | Select-Object -First 30 | ForEach-Object { Write-Host "  $_" }

Say "Backup prod keystore"
$stamp  = Get-Date -Format 'yyyyMMddHHmmss'
$newBak = "$ProdKeystore.bak-$stamp"
Copy-Item -LiteralPath $ProdKeystore -Destination $newBak -Force
Ok "Backup: $newBak"

Say "Push renewed keystore to prod"
Copy-Item -LiteralPath $LocalKeystore -Destination $ProdKeystore -Force
Ok "Pushed: $ProdKeystore"

Say "Also push to OpenText path (in case service reads from there)"
$otcPath = '\\10.168.0.32\e$\OTC\OpenText\Core Archive Connector\java\bin\s4pceotcac.p12'
if (Test-Path -LiteralPath $otcPath) {
    $otcBak = "$otcPath.bak-$stamp"
    Copy-Item -LiteralPath $otcPath -Destination $otcBak -Force
    Copy-Item -LiteralPath $LocalKeystore -Destination $otcPath -Force
    Ok "Also updated: $otcPath (backup: $otcBak)"
} else {
    Warn "OpenText path not accessible - only CERTS path updated"
}

Say "Discover service on $TargetServer"
$svc = $null
try {
    $all = Get-Service -ComputerName $TargetServer -ErrorAction Stop
    $matches = @($all | Where-Object { $_.DisplayName -like '*OpenText*' -or $_.DisplayName -like '*Archive*' -or $_.Name -like '*otcac*' -or $_.Name -like '*opentext*' -or $_.Name -like '*archive*' })
    if ($matches.Count -gt 0) {
        Write-Host "Matched services:" -ForegroundColor Cyan
        $matches | ForEach-Object { Write-Host "  Name='$($_.Name)'  Display='$($_.DisplayName)'  Status=$($_.Status)" }
        $svc = $matches | Select-Object -First 1
    } else {
        Warn "No OpenText/Archive service found. Listing ALL services:"
        $all | Select-Object Name, DisplayName, Status | Sort-Object DisplayName | Format-Table -AutoSize | Out-String | Write-Host
    }
} catch { Warn "Get-Service failed: $_" }

if (-not $svc) {
    Warn "Keystore is on prod. Ask SAP admin to restart the OpenText cert-using service on $TargetServer."
    Write-Host ""
    Ok "DONE (manual service restart pending)"
    Ok "  CERTS path  : $ProdKeystore"
    Ok "  Backup      : $newBak"
    exit 0
}

Say "Restart service '$($svc.Name)'"
$restarted = $false
try {
    Restart-Service -InputObject $svc -Force -ErrorAction Stop
    Start-Sleep -Seconds 3
    $check = Get-Service -ComputerName $TargetServer -Name $svc.Name
    Ok "RPC restart OK. Status: $($check.Status)"
    $restarted = $true
} catch { Warn "RPC failed: $_" }

if (-not $restarted) {
    try {
        & sc.exe "\\$TargetServer" stop $svc.Name 2>&1 | Out-Host
        Start-Sleep -Seconds 5
        & sc.exe "\\$TargetServer" start $svc.Name 2>&1 | Out-Host
        Ok "sc.exe restart completed"
        $restarted = $true
    } catch { Warn "sc.exe failed: $_" }
}

if (-not $restarted) {
    Warn "Restart methods failed. Keystore is on prod. Ask SAP admin to restart '$($svc.Name)'."
}

Write-Host ""
Ok "DONE"
Ok "  CERTS path  : $ProdKeystore"
Ok "  Backup      : $newBak"
if ($svc) { Ok "  Service     : $($svc.Name)" }

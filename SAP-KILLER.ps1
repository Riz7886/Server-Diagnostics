[CmdletBinding()]
param(
    [string]$CertPath   = 'C:\temp\script\OTP.cer',
    [string[]]$Passwords = @('s4hpce','changeit','password','parker','sap','s4pceotcac'),
    [string]$SapServer  = '10.168.0.32',
    [switch]$AutoPush,
    [switch]$RestartService
)

$ErrorActionPreference = 'Continue'

function Say  ($m) { Write-Host ""; Write-Host "[*]  $m" -ForegroundColor Cyan }
function Ok   ($m) { Write-Host "[ok] $m"                -ForegroundColor Green }
function Warn ($m) { Write-Host "[!!] $m"                -ForegroundColor Yellow }
function Die  ($m) { Write-Host "[xx] $m"                -ForegroundColor Red; exit 1 }
function Box  ($m) {
    Write-Host ""
    Write-Host "====================================================================" -ForegroundColor Cyan
    Write-Host "  $m"                                                                 -ForegroundColor Cyan
    Write-Host "====================================================================" -ForegroundColor Cyan
}

Box "SAP CERT RENEWAL - Hunt + Validate + Push"

Say "Preflight"
if (-not (Test-Path -LiteralPath $CertPath)) { Die "Cert not found: $CertPath" }
$certFile = Get-Item -LiteralPath $CertPath
Ok "Cert: $CertPath ($($certFile.Length) bytes, $($certFile.LastWriteTime))"

Say "Parse cert"
$newCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $CertPath
Write-Host "  Subject    : $($newCert.Subject)"
Write-Host "  Issuer     : $($newCert.Issuer)"
Write-Host "  NotBefore  : $($newCert.NotBefore)"
Write-Host "  NotAfter   : $($newCert.NotAfter)"
Write-Host "  Thumbprint : $($newCert.Thumbprint)"
$now = Get-Date
if ($now -gt $newCert.NotAfter)  { Die "Cert expired on $($newCert.NotAfter)" }
if ($now -lt $newCert.NotBefore) { Die "Cert not valid until $($newCert.NotBefore)" }
Ok "Cert valid for $([int]($newCert.NotAfter - $now).TotalDays) days"

$sha       = [System.Security.Cryptography.SHA256]::Create()
$newPubKey = $newCert.GetPublicKey()
$newPubHash = [BitConverter]::ToString($sha.ComputeHash($newPubKey))
Write-Host "  PubKey SHA : $($newPubHash.Substring(0,47))..."

Say "Find keytool"
$kt = $null
$ktCandidates = @(
    'C:\Program Files\Java\*\bin\keytool.exe',
    'C:\Program Files (x86)\Java\*\bin\keytool.exe',
    'C:\Program Files\OpenJDK\*\bin\keytool.exe',
    'C:\Program Files\Eclipse Adoptium\*\bin\keytool.exe',
    'C:\Program Files\Amazon Corretto\*\bin\keytool.exe',
    "\\$SapServer\e$\OTC\OpenText\Core Archive Connector\java\bin\keytool.exe"
)
foreach ($p in $ktCandidates) {
    $f = Get-ChildItem -Path $p -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($f) { $kt = $f.FullName; break }
}
if (-not $kt) { $envKt = Get-Command keytool.exe -ErrorAction SilentlyContinue; if ($envKt) { $kt = $envKt.Path } }
if (-not $kt) { Die "keytool.exe not found" }
Ok "keytool: $kt"

Say "Hunting keystores and CSRs"
$SearchRoots = @(
    "\\$SapServer\e$\OTC\OpenText\Core Archive Connector\java\bin",
    "\\$SapServer\e$\OTC",
    "\\$SapServer\e$\CERTS",
    "\\$SapServer\e$\temp",
    "\\$SapServer\e$\usr",
    "\\$SapServer\e$\SAP",
    "\\$SapServer\e$\sapdata",
    "\\$SapServer\c$\temp",
    "\\$SapServer\c$\Program Files\SAP",
    "\\$SapServer\c$\Program Files\sapdb",
    "\\$SapServer\c$\Users",
    'C:\Program Files\SAP',
    'C:\Program Files\sapdb',
    'C:\Program Files (x86)\SAP',
    'C:\temp',
    'C:\Users',
    'C:\sap',
    'C:\usr\sap'
)

$stores = New-Object System.Collections.Generic.List[object]
$csrs   = New-Object System.Collections.Generic.List[object]
$seen   = @{}

foreach ($root in $SearchRoots) {
    if (-not (Test-Path -LiteralPath $root)) {
        Write-Host "  [skip] $root" -ForegroundColor DarkGray
        continue
    }
    Write-Host "  [scan] $root" -ForegroundColor Gray
    try {
        Get-ChildItem -Path $root -Recurse -Include *.p12,*.pfx,*.jks,*.keystore,*.pse -ErrorAction SilentlyContinue -Force | ForEach-Object {
            if (-not $seen.ContainsKey($_.FullName)) { $stores.Add($_); $seen[$_.FullName] = $true }
        }
        Get-ChildItem -Path $root -Recurse -Include *.csr -ErrorAction SilentlyContinue -Force | ForEach-Object { $csrs.Add($_) }
    } catch {}
}
Ok "Found $($stores.Count) keystores, $($csrs.Count) CSR files"

if ($csrs.Count -gt 0) {
    Say "CSR files (may reveal Phase 1 keystore location - look in same folder)"
    $csrs | Sort-Object LastWriteTime -Descending | ForEach-Object {
        Write-Host "  $($_.FullName)  ($($_.LastWriteTime))"
    }
}

Say "Keystore candidates (newest first)"
$stores | Sort-Object LastWriteTime -Descending | ForEach-Object {
    $tag = if ($_.Extension -eq '.pse') { "PSE" } else { "   " }
    Write-Host ("  [{0}] {1,-70} {2,8} bytes   {3}" -f $tag, $_.FullName, $_.Length, $_.LastWriteTime)
}

Say "Testing each keystore for private-key match against OTP.cer"

$results = New-Object System.Collections.Generic.List[object]

foreach ($s in $stores) {
    Write-Host ""
    Write-Host "--- $($s.FullName)" -ForegroundColor Cyan

    if ($s.Extension -eq '.pse') {
        Write-Host "  [pse]   SAP native format - skipped (use sapgenpse if active)" -ForegroundColor Yellow
        $results.Add([pscustomobject]@{ Source=$s.FullName; Local=$null; Status='PSE'; Alias=$null; Password=$null })
        continue
    }

    $tmp = "C:\temp\script\hunt-$([guid]::NewGuid().ToString('N').Substring(0,8))$($s.Extension)"
    try { Copy-Item -LiteralPath $s.FullName -Destination $tmp -Force -ErrorAction Stop } catch {
        Warn "  copy failed: $_"
        continue
    }

    $matchedAlias = $null
    $matchedPwd   = $null
    foreach ($pwd in $Passwords) {
        try {
            $col = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2Collection
            $col.Import($tmp, $pwd, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
            foreach ($existing in $col) {
                if ($existing.HasPrivateKey) {
                    $existingHash = [BitConverter]::ToString($sha.ComputeHash($existing.GetPublicKey()))
                    if ($existingHash -eq $newPubHash) {
                        $matchedAlias = if ($existing.FriendlyName) { $existing.FriendlyName } else { 's4hpce' }
                        $matchedPwd   = $pwd
                        break
                    }
                }
            }
            if ($matchedPwd) { break }
        } catch {}
    }

    if (-not $matchedPwd) {
        Write-Host "  [skip]  no matching private key in this keystore" -ForegroundColor DarkYellow
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        continue
    }

    Write-Host "  [MATCH] private key MATCHES new cert (alias='$matchedAlias' password='$matchedPwd')" -ForegroundColor Green

    $aliasesToTry = @($matchedAlias)
    if ($matchedAlias -ne 's4hpce') { $aliasesToTry += 's4hpce' }

    $finalStatus = 'ERROR'
    $finalAlias  = $matchedAlias
    $finalErr    = ''
    foreach ($al in $aliasesToTry) {
        $copy2 = "$tmp.try"
        Copy-Item -LiteralPath $tmp -Destination $copy2 -Force
        $impOut = & $kt -importcert -trustcacerts -noprompt -alias $al -file $CertPath -keystore $copy2 -storetype PKCS12 -storepass $matchedPwd 2>&1
        $impExit = $LASTEXITCODE
        $impTxt  = $impOut | Out-String
        if ($impExit -eq 0) {
            $finalStatus = 'CLEAN'; $finalAlias = $al
            Remove-Item $tmp -Force -ErrorAction SilentlyContinue
            Move-Item -LiteralPath $copy2 -Destination $tmp -Force
            break
        } elseif ($impTxt -match 'Failed to establish chain') {
            $finalStatus = 'CHAIN'; $finalAlias = $al; $finalErr = 'Missing intermediate CA'
            Remove-Item $copy2 -Force -ErrorAction SilentlyContinue
            break
        } else {
            $finalErr = ($impTxt -split "`n")[0]
            Remove-Item $copy2 -Force -ErrorAction SilentlyContinue
        }
    }

    switch ($finalStatus) {
        'CLEAN' { Write-Host "  [CLEAN] keytool import succeeded (alias='$finalAlias') - ready to push" -ForegroundColor Green }
        'CHAIN' { Write-Host "  [CHAIN] priv key OK, import needs intermediate CA - ask client for .p7b" -ForegroundColor Yellow }
        'ERROR' { Write-Host "  [ERR]   $finalErr" -ForegroundColor DarkYellow }
    }

    $results.Add([pscustomobject]@{
        Source   = $s.FullName
        Local    = if ($finalStatus -eq 'CLEAN') { $tmp } else { $null }
        Status   = $finalStatus
        Alias    = $finalAlias
        Password = $matchedPwd
        Error    = $finalErr
    })

    if ($finalStatus -ne 'CLEAN') {
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    }
}

Box "RESULTS"

$clean = @($results | Where-Object Status -eq 'CLEAN')
$chain = @($results | Where-Object Status -eq 'CHAIN')
$pse   = @($results | Where-Object Status -eq 'PSE')
$err   = @($results | Where-Object Status -eq 'ERROR')

if ($clean.Count + $chain.Count + $pse.Count + $err.Count -eq 0) {
    Warn "No matching keystore found ANYWHERE."
    Warn "Options:"
    Warn "  (a) try more passwords: -Passwords @('s4hpce','other_pwd1','other_pwd2')"
    Warn "  (b) maybe keystore lives on SAP C: drive SMB share blocked - RDP check"
    Warn "  (c) ask client to re-sign against a CSR from E:\CERTS\s4pceotcac.p12 directly"
    exit 1
}

if ($clean.Count -gt 0) {
    Ok "CLEAN matches ($($clean.Count)):"
    foreach ($m in $clean) {
        Write-Host ""
        Write-Host "  PROD     : $($m.Source)"   -ForegroundColor Green
        Write-Host "  LOCAL    : $($m.Local)"    -ForegroundColor Green
        Write-Host "  ALIAS    : $($m.Alias)"    -ForegroundColor Green
        Write-Host "  PASSWORD : $($m.Password)" -ForegroundColor Green
    }
}
if ($chain.Count -gt 0) {
    Write-Host ""
    Warn "CHAIN-error matches (priv key OK, need intermediate CA): $($chain.Count)"
    foreach ($m in $chain) {
        Write-Host "  PROD  : $($m.Source)" -ForegroundColor Yellow
        Write-Host "  ALIAS : $($m.Alias)"  -ForegroundColor Yellow
    }
    Warn "Fix: ask client for the intermediate CA cert as .p7b (or PEM chain)"
}
if ($pse.Count -gt 0) {
    Write-Host ""
    Warn "PSE files (SAP native, not PKCS12):"
    foreach ($m in $pse) { Write-Host "  $($m.Source)" -ForegroundColor Yellow }
    Warn "If live SAP keystore is .pse, use: sapgenpse maintain_pk -a OTP.cer -p <pse> -x <pin>"
}

if ($clean.Count -eq 0) {
    Write-Host ""
    Warn "No CLEAN matches - cannot auto-push. Fix chain or PSE issue above."
    exit 0
}

if (-not $AutoPush) {
    Box "TO PUSH MANUALLY (default - safe)"
    $stamp = Get-Date -Format 'yyyyMMddHHmmss'
    foreach ($m in $clean) {
        Write-Host ""
        Write-Host "  Copy-Item -LiteralPath '$($m.Source)' -Destination '$($m.Source).bak-$stamp' -Force"
        Write-Host "  Copy-Item -LiteralPath '$($m.Local)'  -Destination '$($m.Source)' -Force"
    }
    Write-Host ""
    Write-Host "  Or re-run with: .\SAP-KILLER.ps1 -AutoPush -RestartService" -ForegroundColor Yellow
    exit 0
}

if ($clean.Count -gt 1) {
    Warn "Multiple CLEAN matches - refusing auto-push. Copy one manually."
    exit 1
}

Box "AUTO-PUSH"
$m     = $clean[0]
$stamp = Get-Date -Format 'yyyyMMddHHmmss'
$bak   = "$($m.Source).bak-$stamp"

Say "Backup prod: $bak"
Copy-Item -LiteralPath $m.Source -Destination $bak -Force
Ok "Backup: $((Get-Item $bak).Length) bytes"

Say "Push renewed keystore: $($m.Source)"
Copy-Item -LiteralPath $m.Local -Destination $m.Source -Force
Ok "Pushed"

Say "Also push to OpenText path if present"
$otcPath = "\\$SapServer\e$\OTC\OpenText\Core Archive Connector\java\bin\s4pceotcac.p12"
if ((Test-Path -LiteralPath $otcPath) -and ($otcPath -ne $m.Source)) {
    $otcBak = "$otcPath.bak-$stamp"
    Copy-Item -LiteralPath $otcPath -Destination $otcBak -Force
    Copy-Item -LiteralPath $m.Local -Destination $otcPath -Force
    Ok "OpenText path updated: $otcPath (backup: $otcBak)"
} else {
    Write-Host "  (OpenText path not found or same as prod path - skipped)" -ForegroundColor Gray
}

Say "Verify renewed cert is in prod"
& $kt -list -v -alias $m.Alias -keystore $m.Source -storetype PKCS12 -storepass $m.Password 2>&1 |
    Select-String -Pattern 'Valid|Owner|Issuer|Serial' | Select-Object -First 8 |
    ForEach-Object { Write-Host "  $_" }

if (-not $RestartService) {
    Write-Host ""
    Ok "DONE - keystore replaced. Service NOT restarted."
    Ok "  PROD : $($m.Source)"
    Ok "  BAK  : $bak"
    Write-Host ""
    Warn "Ask SAP admin to restart OpenText Core Archive Connector, or re-run with -RestartService"
    exit 0
}

Box "SERVICE RESTART on $SapServer"
$svc = $null
try {
    $all = Get-Service -ComputerName $SapServer -ErrorAction Stop
    $matched = @($all | Where-Object { $_.DisplayName -like '*OpenText*' -or $_.DisplayName -like '*Archive*' -or $_.Name -like '*otcac*' -or $_.Name -like '*opentext*' -or $_.Name -like '*archive*' })
    if ($matched.Count -gt 0) {
        Write-Host "Matched services:" -ForegroundColor Cyan
        $matched | ForEach-Object { Write-Host "  $($_.Name)  |  $($_.DisplayName)  |  $($_.Status)" }
        $svc = $matched | Select-Object -First 1
    } else {
        Warn "No OpenText/Archive service found on $SapServer"
    }
} catch { Warn "Get-Service failed: $_" }

if (-not $svc) {
    Warn "Ask SAP admin to restart the cert-using service on $SapServer"
    exit 0
}

Say "Restart '$($svc.Name)'"
$restarted = $false
try {
    Restart-Service -InputObject $svc -Force -ErrorAction Stop
    Start-Sleep -Seconds 3
    $check = Get-Service -ComputerName $SapServer -Name $svc.Name
    Ok "RPC restart OK - status: $($check.Status)"
    $restarted = $true
} catch { Warn "RPC failed: $_" }

if (-not $restarted) {
    try {
        & sc.exe "\\$SapServer" stop $svc.Name 2>&1 | Out-Host
        Start-Sleep -Seconds 5
        & sc.exe "\\$SapServer" start $svc.Name 2>&1 | Out-Host
        Ok "sc.exe restart completed"
        $restarted = $true
    } catch { Warn "sc.exe failed: $_" }
}

Box "ALL DONE"
Ok "  PROD    : $($m.Source)"
Ok "  BAK     : $bak"
Ok "  ALIAS   : $($m.Alias)"
if ($svc)       { Ok "  SERVICE : $($svc.Name)" }
if ($restarted) { Ok "  STATUS  : restarted" } else { Warn "  STATUS  : restart failed - manual action needed" }

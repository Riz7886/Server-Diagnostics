$ErrorActionPreference = 'Continue'

$CertPath   = 'C:\temp\script\OTP.cer'
$Alias      = 's4hpce'
$Password   = 's4hpce'

$SearchRoots = @(
    'C:\Program Files\SAP',
    'C:\Program Files\sapdb',
    'C:\Program Files (x86)\SAP',
    'C:\temp',
    'C:\Users',
    'C:\sap',
    'C:\usr\sap',
    '\\10.168.0.32\e$\OTC\OpenText\Core Archive Connector\java\bin',
    '\\10.168.0.32\e$\OTC',
    '\\10.168.0.32\e$\CERTS',
    '\\10.168.0.32\e$\temp',
    '\\10.168.0.32\e$\SAP',
    '\\10.168.0.32\e$\usr',
    '\\10.168.0.32\c$\Program Files\SAP',
    '\\10.168.0.32\c$\Program Files\sapdb',
    '\\10.168.0.32\c$\temp',
    '\\10.168.0.32\c$\Users'
)

function Say ($m) { Write-Host ""; Write-Host "[*]  $m" -ForegroundColor Cyan }
function Ok  ($m) { Write-Host "[ok] $m" -ForegroundColor Green }
function Warn($m) { Write-Host "[!!] $m" -ForegroundColor Yellow }
function Die ($m) { Write-Host "[xx] $m" -ForegroundColor Red; exit 1 }

Say "Preflight"
if (-not (Test-Path $CertPath)) { Die "Cert missing: $CertPath" }
Ok "Cert: $CertPath"

Say "Find keytool"
$kt = (Get-ChildItem 'C:\Program Files\Java\*\bin\keytool.exe' -ErrorAction SilentlyContinue | Select-Object -First 1).FullName
if (-not $kt) { $kt = (Get-ChildItem 'C:\Program Files (x86)\Java\*\bin\keytool.exe' -ErrorAction SilentlyContinue | Select-Object -First 1).FullName }
if (-not $kt) { $kt = (Get-Command keytool.exe -ErrorAction SilentlyContinue).Source }
if (-not $kt) { Die "keytool not found" }
Ok "keytool: $kt"

Say "Scanning all roots for keystores"
$stores = New-Object System.Collections.Generic.List[object]
$seen = @{}

foreach ($root in $SearchRoots) {
    if (-not (Test-Path -LiteralPath $root)) {
        Write-Host "  [skip] $root (not accessible)" -ForegroundColor DarkGray
        continue
    }
    Write-Host "  [scan] $root" -ForegroundColor Gray
    $hits = Get-ChildItem -Path $root -Recurse -Include *.p12,*.pfx,*.jks,*.keystore,*.pse -ErrorAction SilentlyContinue -Force
    foreach ($h in $hits) {
        if (-not $seen.ContainsKey($h.FullName)) {
            $stores.Add($h)
            $seen[$h.FullName] = $true
        }
    }
}

Ok "Total keystore candidates: $($stores.Count)"
Write-Host ""
$stores | Sort-Object LastWriteTime -Descending | ForEach-Object {
    Write-Host ("  {0,-80}  {1,7} bytes  {2}" -f $_.FullName, $_.Length, $_.LastWriteTime)
}

$matches = New-Object System.Collections.Generic.List[object]

foreach ($s in $stores) {
    Write-Host ""
    Write-Host "=== $($s.FullName) ===" -ForegroundColor Cyan

    if ($s.Extension -eq '.pse') {
        Write-Host "  [.pse - SAP PSE format, not PKCS12. Skipping keytool test]" -ForegroundColor DarkGray
        Write-Host "  NOTE: if this IS the live SAP keystore, it needs sapgenpse not keytool" -ForegroundColor Yellow
        continue
    }

    $tmp = "C:\temp\script\hunt-$([guid]::NewGuid().ToString('N').Substring(0,8))$($s.Extension)"
    try { Copy-Item -LiteralPath $s.FullName -Destination $tmp -Force -ErrorAction Stop } catch { Warn "  copy failed: $_"; continue }

    $opened = $false
    $type = 'PKCS12'
    $listOut = & $kt -list -keystore $tmp -storetype PKCS12 -storepass $Password 2>&1
    if ($LASTEXITCODE -eq 0) { $opened = $true }
    else {
        $listOut = & $kt -list -keystore $tmp -storetype JKS -storepass $Password 2>&1
        if ($LASTEXITCODE -eq 0) { $opened = $true; $type = 'JKS' }
    }

    if (-not $opened) {
        Warn "  cannot open with password '$Password' (different password or corrupt)"
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        continue
    }

    Write-Host "  opened as $type" -ForegroundColor Green
    $listOut | Select-Object -First 20 | ForEach-Object { Write-Host "    $_" }

    $impOut  = & $kt -importcert -trustcacerts -noprompt -alias $Alias -file $CertPath -keystore $tmp -storetype $type -storepass $Password 2>&1
    $impExit = $LASTEXITCODE

    if ($impExit -eq 0) {
        Write-Host ""
        Write-Host "  >>> MATCH - cert imported cleanly. THIS is the right keystore. <<<" -ForegroundColor Green
        $matches.Add([pscustomobject]@{ Source=$s.FullName; Local=$tmp; Type=$type; ChainIssue=$false })
    } else {
        $txt = $impOut | Out-String
        if ($txt -match "Public keys.*don't match") {
            Write-Host "    -> wrong private key" -ForegroundColor Yellow
        } elseif ($txt -match 'Failed to establish chain') {
            Write-Host ""
            Write-Host "  >>> LIKELY MATCH (private key OK - only missing intermediate CA) <<<" -ForegroundColor Yellow
            $matches.Add([pscustomobject]@{ Source=$s.FullName; Local=$tmp; Type=$type; ChainIssue=$true })
        } else {
            $impOut | Select-Object -First 3 | ForEach-Object { Write-Host "    $_" -ForegroundColor Yellow }
        }
        if (-not ($matches | Where-Object { $_.Local -eq $tmp })) {
            Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        }
    }
}

Write-Host ""
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host "  RESULT" -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan

if ($matches.Count -eq 0) {
    Warn "No matching keystore found."
    Warn "Try: (a) different password, (b) ask client for intermediate CA chain, (c) check if SAP uses .pse (sapgenpse workflow)"
    exit 1
}

Ok "Matches: $($matches.Count)"
foreach ($m in $matches) {
    Write-Host ""
    Write-Host "  PROD FILE : $($m.Source)"           -ForegroundColor Green
    Write-Host "  RENEWED   : $($m.Local)"            -ForegroundColor Green
    Write-Host "  TYPE      : $($m.Type)"             -ForegroundColor Green
    if ($m.ChainIssue) {
        Write-Host "  NOTE      : CHAIN error - need intermediate CA from client (.p7b file)" -ForegroundColor Yellow
    } else {
        Write-Host "  NOTE      : clean import, ready to push" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "=====================================================" -ForegroundColor Yellow
Write-Host "  To push the clean match to prod:" -ForegroundColor Yellow
Write-Host "=====================================================" -ForegroundColor Yellow
$stamp = Get-Date -Format 'yyyyMMddHHmmss'
foreach ($m in $matches | Where-Object { -not $_.ChainIssue }) {
    Write-Host ""
    Write-Host "  Copy-Item -LiteralPath '$($m.Source)' -Destination '$($m.Source).bak-$stamp' -Force"
    Write-Host "  Copy-Item -LiteralPath '$($m.Local)'  -Destination '$($m.Source)' -Force"
}
Write-Host ""
Write-Host "Then restart the OpenText / SAP service on 10.168.0.32." -ForegroundColor Yellow

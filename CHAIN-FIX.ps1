$ErrorActionPreference = 'Continue'

$CertPath   = 'C:\temp\script\OTP.cer'
$Password   = 's4hpce'
$Alias      = 's4hpce'
$LocalKs    = 'C:\temp\script\otc-fix.p12'
$ProdOtc    = '\\10.168.0.32\e$\OTC\OpenText\Core Archive Connector\java\bin\s4pceotcac.p12'
$ProdCerts  = '\\10.168.0.32\e$\CERTS\s4pceotcac.p12'

function Say ($m) { Write-Host ""; Write-Host "[*]  $m" -ForegroundColor Cyan }
function Ok  ($m) { Write-Host "[ok] $m" -ForegroundColor Green }
function Warn($m) { Write-Host "[!!] $m" -ForegroundColor Yellow }
function Die ($m) { Write-Host "[xx] $m" -ForegroundColor Red; exit 1 }

Say "Preflight"
if (-not (Test-Path $CertPath)) { Die "Cert missing: $CertPath" }
if (-not (Test-Path $ProdOtc))  { Die "Prod OTC keystore unreachable: $ProdOtc" }

$kt = (Get-ChildItem 'C:\Program Files\Java\*\bin\keytool.exe' -ErrorAction SilentlyContinue | Select-Object -First 1).FullName
if (-not $kt) { $kt = (Get-ChildItem 'C:\Program Files (x86)\Java\*\bin\keytool.exe' -ErrorAction SilentlyContinue | Select-Object -First 1).FullName }
if (-not $kt) { Die "keytool.exe not found" }
Ok "keytool: $kt"

Say "Fresh copy of prod keystore"
if (Test-Path $LocalKs) { Remove-Item $LocalKs -Force }
Copy-Item $ProdOtc $LocalKs -Force
Ok "Local: $LocalKs"

Say "Parse signed cert"
$leaf = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $CertPath
Write-Host "  Subject : $($leaf.Subject)"
Write-Host "  Issuer  : $($leaf.Issuer)"
Write-Host "  Expires : $($leaf.NotAfter)"

Say "Build chain from Windows cert store"
$chain = New-Object System.Security.Cryptography.X509Certificates.X509Chain
$chain.ChainPolicy.RevocationMode = 'NoCheck'
$built = $chain.Build($leaf)
Write-Host "  Chain elements: $($chain.ChainElements.Count)  BuildOK: $built"

if ($chain.ChainElements.Count -lt 2) {
    Warn "Windows could not find intermediate CA."
    Warn "Trying to download intermediate from AIA extension..."
    $aia = $leaf.Extensions | Where-Object { $_.Oid.Value -eq '1.3.6.1.5.5.7.1.1' }
    if ($aia) {
        $aiaStr = (New-Object System.Security.Cryptography.AsnEncodedData $aia.Oid, $aia.RawData).Format($false)
        Write-Host "  AIA: $aiaStr"
        if ($aiaStr -match 'https?://\S+?\.(crt|cer|p7b|p7c)') {
            $url = $matches[0] -replace '[,)\s]+$',''
            Write-Host "  Download: $url" -ForegroundColor Yellow
            $intFile = 'C:\temp\script\intermediate-dl.cer'
            try {
                Invoke-WebRequest -Uri $url -OutFile $intFile -UseBasicParsing -ErrorAction Stop
                Ok "Downloaded: $intFile"
                Say "Import downloaded intermediate"
                & $kt -importcert -trustcacerts -noprompt -alias "intermediate-ca" -file $intFile -keystore $LocalKs -storetype PKCS12 -storepass $Password 2>&1 | ForEach-Object { Write-Host "    $_" }
            } catch {
                Warn "Download failed: $_"
                Die "Need the intermediate CA. Ask client for the .p7b or intermediate .cer file."
            }
        } else {
            Die "No downloadable intermediate URL in cert AIA. Ask client for .p7b"
        }
    } else {
        Die "Cert has no AIA extension. Ask client for intermediate .p7b"
    }
} else {
    Say "Export and import intermediates from Windows chain"
    $i = 0
    foreach ($el in $chain.ChainElements) {
        if ($i -gt 0) {
            $f = "C:\temp\script\chain-ca-$i.cer"
            [System.IO.File]::WriteAllBytes($f, $el.Certificate.Export('Cert'))
            Write-Host "  Importing: $($el.Certificate.Subject)" -ForegroundColor Yellow
            & $kt -importcert -trustcacerts -noprompt -alias "chainca$i" -file $f -keystore $LocalKs -storetype PKCS12 -storepass $Password 2>&1 | ForEach-Object { Write-Host "    $_" }
        }
        $i++
    }
}

Say "Import leaf (signed cert) into keystore"
$impOut  = & $kt -importcert -trustcacerts -noprompt -alias $Alias -file $CertPath -keystore $LocalKs -storetype PKCS12 -storepass $Password 2>&1
$impExit = $LASTEXITCODE
$impOut | ForEach-Object { Write-Host "    $_" }

if ($impExit -ne 0) {
    $txt = ($impOut | Out-String)
    if ($txt -match "Failed to establish chain") {
        Die "Chain STILL broken. The intermediate we imported is not the right one. Ask client for their intermediate .p7b file."
    } else {
        Die "Import failed - see keytool output above. Prod is UNTOUCHED."
    }
}

Ok "Leaf cert imported into local copy successfully"

Say "Verify the renewed cert"
& $kt -list -v -alias $Alias -keystore $LocalKs -storetype PKCS12 -storepass $Password 2>&1 | Select-Object -First 15 | ForEach-Object { Write-Host "    $_" }

Write-Host ""
Write-Host "=====================================================" -ForegroundColor Green
Write-Host "  SUCCESS - local keystore ready at: $LocalKs" -ForegroundColor Green
Write-Host "=====================================================" -ForegroundColor Green
Write-Host ""
Write-Host "To push to prod NOW, run:" -ForegroundColor Yellow
$stamp = Get-Date -Format 'yyyyMMddHHmmss'
Write-Host ""
Write-Host "  Copy-Item '$ProdOtc'   '$ProdOtc.bak-$stamp' -Force"
Write-Host "  Copy-Item '$ProdCerts' '$ProdCerts.bak-$stamp' -Force"
Write-Host "  Copy-Item '$LocalKs'   '$ProdOtc'   -Force"
Write-Host "  Copy-Item '$LocalKs'   '$ProdCerts' -Force"
Write-Host ""
Write-Host "Then restart OpenText service on 10.168.0.32" -ForegroundColor Yellow

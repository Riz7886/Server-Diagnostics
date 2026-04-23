#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Position=0)]
    [ValidateSet('Csr','Import','Push','Verify')]
    [string]$Phase = 'Csr',

    [string]$RemoteShare = '\\10.168.0.32\e$\CERTS',

    [string]$Alias   = 's4hpce',
    [string]$Subject = 'CN=ns2otpapp.sap.parker.corp, OU=COR, O=Parker Hannifin, L=Cleaveland, C=US',
    [string]$SanDns  = 'ns2otpapp.sap.parker.corp',
    [int]$KeySize    = 2048,

    [string]$WorkDir = (Join-Path $PSScriptRoot 'work'),

    [string]$SignedCert,
    [string]$KeytoolPath
)

$ErrorActionPreference = 'Stop'

function Say {
    param([string]$Msg, [string]$Color = 'Cyan')
    Write-Host "[*] $Msg" -ForegroundColor $Color
}

function Find-Keytool {
    param([string]$Hint)
    if ($Hint -and (Test-Path $Hint)) { return $Hint }
    $cmd = Get-Command keytool.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Path }
    $paths = @(
        'C:\Program Files\Java\jre*\bin\keytool.exe',
        'C:\Program Files\Java\jdk*\bin\keytool.exe',
        'C:\Program Files\Eclipse Adoptium\*\bin\keytool.exe',
        'C:\Program Files\OpenJDK\*\bin\keytool.exe',
        'C:\Program Files\Amazon Corretto\*\bin\keytool.exe',
        'C:\Program Files (x86)\Java\*\bin\keytool.exe',
        'C:\Program Files (x86)\SAP\FrontEnd\SAPGUI\*\bin\keytool.exe',
        'C:\Program Files\SAP\*\bin\keytool.exe',
        'C:\Java\*\bin\keytool.exe',
        'C:\OpenJDK*\bin\keytool.exe',
        'D:\Java\*\bin\keytool.exe',
        'E:\Java\*\bin\keytool.exe'
    )
    foreach ($p in $paths) {
        $f = Get-ChildItem -Path $p -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($f) { return $f.FullName }
    }
    return $null
}

if (-not (Test-Path $WorkDir)) { New-Item -Path $WorkDir -ItemType Directory -Force | Out-Null }
$keystoreLocal = Join-Path $WorkDir 's4pceotcac_new.p12'
$csrLocal      = Join-Path $WorkDir 'OTP.csr'
$pwStore       = Join-Path $WorkDir 'pw.xml'

$kt = Find-Keytool -Hint $KeytoolPath
if (-not $kt) {
    Say "keytool.exe not found on jump server." 'Red'
    Say "Run: Get-ChildItem 'C:\Program Files' -Recurse -Filter keytool.exe -ErrorAction SilentlyContinue" 'Yellow'
    Say "Then re-run with:  .\RENEW.ps1 $Phase -KeytoolPath 'C:\path\to\keytool.exe'" 'Yellow'
    exit 1
}
Say "keytool: $kt" 'Green'

switch ($Phase) {

    'Csr' {
        Say "Phase 1: generate new keystore + CSR (locally on jump server)" 'Cyan'
        $pw  = Read-Host -AsSecureString 'NEW keystore password'
        $pw2 = Read-Host -AsSecureString 'Confirm password'
        $p1 = [System.Net.NetworkCredential]::new('', $pw).Password
        $p2 = [System.Net.NetworkCredential]::new('', $pw2).Password
        if ($p1 -ne $p2) { Say 'Passwords do not match' 'Red'; exit 1 }
        if ($p1.Length -lt 6) { Say 'Password too short (6+ chars)' 'Red'; exit 1 }
        $pw | Export-Clixml -Path $pwStore
        Say "Password saved (DPAPI): $pwStore" 'Green'

        if (Test-Path $keystoreLocal) { Remove-Item $keystoreLocal -Force }
        if (Test-Path $csrLocal)      { Remove-Item $csrLocal -Force }

        $genArgs = @(
            '-genkeypair',
            '-alias', $Alias,
            '-keyalg', 'RSA',
            '-keysize', "$KeySize",
            '-dname', $Subject,
            '-ext', "SAN=dns:$SanDns",
            '-keystore', $keystoreLocal,
            '-storetype', 'PKCS12',
            '-storepass', $p1,
            '-keypass', $p1,
            '-validity', '730'
        )
        & $kt @genArgs
        if ($LASTEXITCODE -ne 0) { Say "keytool -genkeypair failed: $LASTEXITCODE" 'Red'; exit 1 }
        Say "Keystore created: $keystoreLocal" 'Green'

        $reqArgs = @(
            '-certreq',
            '-alias', $Alias,
            '-keystore', $keystoreLocal,
            '-storetype', 'PKCS12',
            '-storepass', $p1,
            '-file', $csrLocal,
            '-ext', "SAN=dns:$SanDns"
        )
        & $kt @reqArgs
        if ($LASTEXITCODE -ne 0) { Say "keytool -certreq failed: $LASTEXITCODE" 'Red'; exit 1 }
        Say "CSR created: $csrLocal" 'Green'

        try {
            if (-not (Test-Path $RemoteShare)) {
                Say "Remote share not visible: $RemoteShare (copy keystore manually later)" 'Yellow'
            } else {
                Copy-Item $keystoreLocal (Join-Path $RemoteShare 's4pceotcac_new.p12') -Force
                Say "Keystore pushed to $RemoteShare\s4pceotcac_new.p12" 'Green'
            }
        } catch {
            Say "SMB push failed: $($_.Exception.Message)" 'Yellow'
            Say "Copy manually via File Explorer: $keystoreLocal -> $RemoteShare" 'Yellow'
        }

        Write-Host ''
        Write-Host '=========== CSR BELOW (paste into email to Rhonda) ===========' -ForegroundColor Yellow
        Get-Content $csrLocal
        Write-Host '==============================================================' -ForegroundColor Yellow
        Write-Host ''
        Say "Phase 1 done. Email CSR above." 'Green'
        Say "Tomorrow, after signed cert returns:  .\RENEW.ps1 Import -SignedCert C:\path\OTP.cer" 'Cyan'
    }

    'Import' {
        if (-not $SignedCert) { Say 'Missing -SignedCert C:\path\OTP.cer' 'Red'; exit 1 }
        if (-not (Test-Path $SignedCert)) { Say "Not found: $SignedCert" 'Red'; exit 1 }
        if (-not (Test-Path $pwStore)) { Say 'Password not saved - run Csr phase first' 'Red'; exit 1 }
        if (-not (Test-Path $keystoreLocal)) { Say "Missing keystore: $keystoreLocal - run Csr phase first" 'Red'; exit 1 }

        $pw = Import-Clixml -Path $pwStore
        $p1 = [System.Net.NetworkCredential]::new('', $pw).Password

        $impArgs = @(
            '-importcert',
            '-trustcacerts',
            '-noprompt',
            '-alias', $Alias,
            '-file', $SignedCert,
            '-keystore', $keystoreLocal,
            '-storetype', 'PKCS12',
            '-storepass', $p1
        )
        & $kt @impArgs
        if ($LASTEXITCODE -ne 0) { Say "keytool -importcert failed: $LASTEXITCODE" 'Red'; exit 1 }
        Say 'Signed cert imported into local keystore' 'Green'

        & $kt -list -v -alias $Alias -keystore $keystoreLocal -storetype PKCS12 -storepass $p1
        if ($LASTEXITCODE -ne 0) { Say 'keytool -list failed after import' 'Red'; exit 1 }

        try {
            $stamp  = Get-Date -Format 'yyyyMMddHHmmss'
            $prod   = Join-Path $RemoteShare 's4pceotcac.p12'
            $backup = Join-Path $RemoteShare "s4pceotcac.p12.bak-$stamp"
            if (Test-Path $prod) {
                Copy-Item $prod $backup -Force
                Say "Production backup saved: $backup" 'Green'
            }
            Copy-Item $keystoreLocal $prod -Force
            Say "Production keystore updated: $prod" 'Green'
        } catch {
            Say "SMB push failed: $($_.Exception.Message)" 'Red'
            Say "Manually copy: $keystoreLocal -> $RemoteShare\s4pceotcac.p12" 'Yellow'
            exit 1
        }

        Write-Host ''
        Say "DONE. Restart target service: OpenText Core Archive Connector" 'Yellow'
    }

    'Push' {
        if (-not (Test-Path $keystoreLocal)) { Say "No local keystore: $keystoreLocal" 'Red'; exit 1 }
        Copy-Item $keystoreLocal (Join-Path $RemoteShare 's4pceotcac_new.p12') -Force
        Say 'Pushed' 'Green'
    }

    'Verify' {
        if (-not (Test-Path $pwStore)) { Say 'No password file' 'Red'; exit 1 }
        if (-not (Test-Path $keystoreLocal)) { Say "No keystore: $keystoreLocal" 'Red'; exit 1 }
        $pw = Import-Clixml -Path $pwStore
        $p1 = [System.Net.NetworkCredential]::new('', $pw).Password
        & $kt -list -v -alias $Alias -keystore $keystoreLocal -storetype PKCS12 -storepass $p1
    }
}

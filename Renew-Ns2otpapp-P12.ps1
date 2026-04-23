[CmdletBinding()]
param(
    [Parameter(Mandatory=$true,Position=0)][ValidateSet('GenerateCsr','ImportSignedCert','Verify','Rollback')][string]$Phase,
    [string]$Target = '10.168.0.32',
    [string[]]$FallbackTargets = @(),
    [pscredential]$Credential,
    [string]$KeystorePath = 'E:\CERTS\s4pceotcac.p12',
    [string]$NewKeystorePath = 'E:\CERTS\s4pceotcac_new.p12',
    [string]$CsrPath = 'E:\CERTS\OTP.csr',
    [string]$CerPath = 'E:\CERTS\OTP.cer',
    [string]$Alias = 's4hpce',
    [string]$Subject = 'CN=ns2otpapp.sap.parker.corp, OU=COR, O=Parker Hannifin, L=Cleaveland, C=US',
    [string]$SanDns = 'ns2otpapp.sap.parker.corp',
    [int]$KeySize = 2048,
    [int]$ValidityDays = 825,
    [string]$ServiceName = 'OpenText Core Archive Connector',
    [string]$KeytoolPath,
    [string]$SignedCertFile,
    [string]$LocalOutDir = $(Join-Path $PSScriptRoot 'out'),
    [string]$LogDir = $(Join-Path $PSScriptRoot 'logs'),
    [string]$SecretFile = $(Join-Path $PSScriptRoot 'secrets\ns2otpapp-p12.xml'),
    [switch]$AutoSwap,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$runId = Get-Date -Format 'yyyyMMdd-HHmmss'

foreach ($d in @($LocalOutDir,$LogDir,(Split-Path -Parent $SecretFile))) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}
$LogFile = Join-Path $LogDir ("renew-ns2otpapp-{0}-{1}.log" -f $Phase,$runId)

function Write-L {
    param([string]$Level,[string]$Msg)
    $line = "[{0}] {1} {2}" -f (Get-Date -Format 'HH:mm:ss'),$Level.PadRight(5),$Msg
    $color = switch ($Level) { 'ERROR'{'Red'} 'WARN'{'Yellow'} 'OK'{'Green'} 'HEAD'{'Cyan'} default {'Gray'} }
    Write-Host $line -ForegroundColor $color
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

function ConvertTo-Plain {
    param([securestring]$S)
    return [System.Net.NetworkCredential]::new('',$S).Password
}

function Save-Secret {
    param([string]$Path,[securestring]$Value,[hashtable]$Meta = @{})
    [pscustomobject]@{ Secret=$Value; Meta=$Meta; Saved=(Get-Date) } | Export-Clixml -Path $Path -Force
}

function Get-Secret {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $null }
    return Import-Clixml -Path $Path
}

function Get-LiveTarget {
    param([string[]]$Candidates,[pscredential]$Cred)
    foreach ($t in $Candidates) {
        try {
            $null = Test-WSMan -ComputerName $t -Credential $Cred -Authentication Default -ErrorAction Stop
            Write-L -Level OK -Msg ("WinRM reachable: {0}" -f $t)
            return $t
        } catch { Write-L -Level WARN -Msg ("WinRM failed for {0}: {1}" -f $t,$_.Exception.Message) }
    }
    return $null
}

function Invoke-OnTarget {
    param([System.Management.Automation.Runspaces.PSSession]$Session,[scriptblock]$Script,[object[]]$Args = @())
    return Invoke-Command -Session $Session -ScriptBlock $Script -ArgumentList $Args
}

function Resolve-Keytool {
    param([System.Management.Automation.Runspaces.PSSession]$Session,[string]$Explicit)
    if ($Explicit) {
        $ok = Invoke-OnTarget -Session $Session -Script { param($p) Test-Path $p } -Args @($Explicit)
        if ($ok) { return $Explicit }
        throw "keytool path provided but not found on target: $Explicit"
    }
    $paths = Invoke-OnTarget -Session $Session -Script {
        $found = @()
        $roots = @(
            'C:\Program Files\Java','C:\Program Files (x86)\Java',
            'C:\Program Files\OpenJDK','C:\Program Files\Eclipse Adoptium',
            'C:\Program Files\Amazon Corretto','C:\Program Files\Zulu',
            'E:\OTC','E:\Java','D:\Java'
        )
        foreach ($r in $roots) {
            if (Test-Path $r) {
                $hits = Get-ChildItem $r -Filter keytool.exe -Recurse -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
                if ($hits) { $found += $hits }
            }
        }
        try { $p = (Get-Command keytool.exe -ErrorAction Stop).Source; if ($p) { $found += $p } } catch {}
        $found | Sort-Object -Unique
    }
    if (-not $paths) { throw "keytool.exe not found on target. Pass -KeytoolPath explicitly." }
    $chosen = $paths | Select-Object -First 1
    Write-L -Level INFO -Msg ("keytool: {0}" -f $chosen)
    if (($paths | Measure-Object).Count -gt 1) {
        Write-L -Level INFO -Msg ("Other keytool copies on target: {0}" -f (($paths | Select-Object -Skip 1) -join ' | '))
    }
    return $chosen
}

function Read-NewPassword {
    param([string]$Label = 'keystore')
    while ($true) {
        $p1 = Read-Host -AsSecureString -Prompt ("Enter NEW {0} password (min 12, letters+digits, no quotes/spaces)" -f $Label)
        $p2 = Read-Host -AsSecureString -Prompt 'Confirm'
        $u1 = ConvertTo-Plain $p1
        $u2 = ConvertTo-Plain $p2
        if ($u1 -ne $u2) { Write-Host 'Passwords do not match, try again.' -ForegroundColor Yellow; continue }
        if ($u1.Length -lt 12) { Write-Host 'Too short (min 12).' -ForegroundColor Yellow; continue }
        if ($u1 -match '\s' -or $u1 -match '["'']') { Write-Host 'No whitespace or quotes.' -ForegroundColor Yellow; continue }
        if ($u1 -notmatch '[A-Z]' -or $u1 -notmatch '[a-z]' -or $u1 -notmatch '\d') { Write-Host 'Need upper + lower + digit.' -ForegroundColor Yellow; continue }
        return $p1
    }
}

Write-L -Level HEAD -Msg ("Phase={0}  RunId={1}" -f $Phase,$runId)
Write-L -Level INFO -Msg ("Log: {0}" -f $LogFile)

if (-not $Credential) {
    $Credential = Get-Credential -UserName 'c5406751' -Message 'Enter password for c5406751 (add DOMAIN\ prefix if needed, e.g. PARKER\c5406751)'
}

$allTargets = @($Target) + $FallbackTargets | Select-Object -Unique
$live = Get-LiveTarget -Candidates $allTargets -Cred $Credential
if (-not $live) { throw "No target reachable via WinRM. Tried: $($allTargets -join ', '). Check VPN / firewall / credentials." }

$session = New-PSSession -ComputerName $live -Credential $Credential -Authentication Default -ErrorAction Stop
Write-L -Level OK -Msg ("PSSession open to {0}" -f $live)

try {
    $keytool = Resolve-Keytool -Session $session -Explicit $KeytoolPath

    switch ($Phase) {

        'GenerateCsr' {
            Write-L -Level HEAD -Msg 'PHASE 1 - generate fresh keystore + CSR'

            if (Test-Path $SecretFile) {
                if (-not $Force) {
                    Write-L -Level WARN -Msg 'Secret file already exists from a previous phase-1 run.'
                    $ans = Read-Host 'Overwrite and generate fresh keystore? (yes/no)'
                    if ($ans -ne 'yes') { throw 'Aborted. Existing secret kept.' }
                }
            }
            $pwSecure = Read-NewPassword -Label 'NEW keystore'
            $newPw = ConvertTo-Plain $pwSecure
            Save-Secret -Path $SecretFile -Value $pwSecure -Meta @{ Keystore=$NewKeystorePath; Alias=$Alias; Subject=$Subject; RunId=$runId; Target=$live }
            Write-L -Level OK -Msg ("New password saved (DPAPI, this user only): {0}" -f $SecretFile)

            Invoke-OnTarget -Session $session -Script {
                param($Dir) if (-not (Test-Path $Dir)) { New-Item -ItemType Directory -Path $Dir -Force | Out-Null }
            } -Args @((Split-Path -Parent $KeystorePath))

            $oldExists = Invoke-OnTarget -Session $session -Script { param($p) Test-Path $p } -Args @($KeystorePath)
            if ($oldExists) {
                $bk = "$KeystorePath.$runId.bak"
                Invoke-OnTarget -Session $session -Script { param($s,$d) Copy-Item $s $d -Force } -Args @($KeystorePath,$bk)
                Write-L -Level OK -Msg ("Existing keystore backed up on target: {0}" -f $bk)
            } else {
                Write-L -Level WARN -Msg ("No existing keystore at {0} - this is a fresh install path." -f $KeystorePath)
            }

            $newExists = Invoke-OnTarget -Session $session -Script { param($p) Test-Path $p } -Args @($NewKeystorePath)
            if ($newExists) {
                $bk2 = "$NewKeystorePath.$runId.bak"
                Invoke-OnTarget -Session $session -Script { param($s,$d) Move-Item $s $d -Force } -Args @($NewKeystorePath,$bk2)
                Write-L -Level WARN -Msg ("Previous new-keystore moved aside: {0}" -f $bk2)
            }

            Write-L -Level INFO -Msg 'keytool -genkeypair (fresh keypair in new keystore)'
            $gen = Invoke-OnTarget -Session $session -Script {
                param($kt,$ks,$alias,$pw,$subj,$san,$size,$validity)
                $a = @('-genkeypair','-keystore',$ks,'-storetype','PKCS12','-storepass',$pw,'-keypass',$pw,
                       '-alias',$alias,'-keyalg','RSA','-keysize',$size,'-dname',$subj,
                       '-ext',("SAN=dns:{0}" -f $san),'-validity',$validity)
                $out = & $kt @a 2>&1
                [pscustomobject]@{ Code=$LASTEXITCODE; Out=($out | Out-String) }
            } -Args @($keytool,$NewKeystorePath,$Alias,$newPw,$Subject,$SanDns,$KeySize,$ValidityDays)
            Add-Content $LogFile -Value $gen.Out
            if ($gen.Code -ne 0) { throw ("keytool genkeypair failed (exit {0}). See log." -f $gen.Code) }
            Write-L -Level OK -Msg 'Keypair generated.'

            Write-L -Level INFO -Msg 'keytool -certreq (export CSR)'
            $req = Invoke-OnTarget -Session $session -Script {
                param($kt,$ks,$alias,$pw,$csr,$san)
                $a = @('-certreq','-keystore',$ks,'-storetype','PKCS12','-storepass',$pw,
                       '-alias',$alias,'-file',$csr,'-ext',("SAN=dns:{0}" -f $san))
                $out = & $kt @a 2>&1
                [pscustomobject]@{ Code=$LASTEXITCODE; Out=($out | Out-String) }
            } -Args @($keytool,$NewKeystorePath,$Alias,$newPw,$CsrPath,$SanDns)
            Add-Content $LogFile -Value $req.Out
            if ($req.Code -ne 0) { throw ("keytool certreq failed (exit {0}). See log." -f $req.Code) }

            $csrRemoteExists = Invoke-OnTarget -Session $session -Script { param($p) Test-Path $p } -Args @($CsrPath)
            if (-not $csrRemoteExists) { throw "CSR not found on target at $CsrPath" }

            $localCsr = Join-Path $LocalOutDir ("OTP-{0}.csr" -f $runId)
            Copy-Item -Path $CsrPath -Destination $localCsr -FromSession $session -Force
            $localCsrAlias = Join-Path $LocalOutDir 'OTP.csr'
            Copy-Item -Path $localCsr -Destination $localCsrAlias -Force

            $eDriveCsr = $null
            if (Test-Path 'E:\') {
                try {
                    $eDriveCsr = "E:\OTP-$runId.csr"
                    Copy-Item -Path $localCsr -Destination $eDriveCsr -Force
                    Copy-Item -Path $localCsr -Destination 'E:\OTP.csr' -Force
                    Write-L -Level OK -Msg ("CSR also saved to jump server E: drive: {0}" -f $eDriveCsr)
                    Write-L -Level OK -Msg 'CSR also saved as:  E:\OTP.csr'
                } catch {
                    Write-L -Level WARN -Msg ("Could not write to E:\ on jump server: {0}" -f $_.Exception.Message)
                    $eDriveCsr = $null
                }
            } else {
                Write-L -Level INFO -Msg 'No E: drive on this jump server - CSR stays in script out\ folder.'
            }

            $csrText = Get-Content -Path $localCsr -Raw

            Write-L -Level OK -Msg ("CSR on jump server: {0}" -f $localCsr)
            Write-L -Level OK -Msg ("CSR on target:      {0}" -f $CsrPath)
            Write-L -Level HEAD -Msg 'PHASE 1 COMPLETE'

            Write-Host ''
            Write-Host '============ CSR TEXT (copy this into the email body) ============' -ForegroundColor Cyan
            Write-Host $csrText -ForegroundColor White
            Write-Host '===================================================================' -ForegroundColor Cyan
            Write-Host ''
            Write-Host '===================== NEXT STEPS =====================' -ForegroundColor Cyan
            Write-Host '  1. Email the CSR to Rhonda (Zeng, Rhonda). Two options:' -ForegroundColor White
            Write-Host '       a) Attach the file, OR' -ForegroundColor White
            Write-Host '       b) Paste the CSR text block above into the email body.' -ForegroundColor White
            Write-Host '     Files ready for attachment:' -ForegroundColor Gray
            Write-Host ("       {0}" -f $localCsrAlias) -ForegroundColor Yellow
            if ($eDriveCsr) { Write-Host ("       {0}" -f $eDriveCsr) -ForegroundColor Yellow; Write-Host '       E:\OTP.csr' -ForegroundColor Yellow }
            Write-Host ''
            Write-Host '  2. Ask Rhonda for: signed leaf cert + FULL CA chain' -ForegroundColor White
            Write-Host '       (Parker Hannifin General Issuing CA v2 + Root)' -ForegroundColor Gray
            Write-Host '       preferred: one .cer (PEM) or .p7b containing leaf + chain.' -ForegroundColor Gray
            Write-Host ''
            Write-Host '  3. When she returns the signed cert, either:' -ForegroundColor White
            Write-Host '       a) Save to jump server (any path) and run:' -ForegroundColor White
            Write-Host '          .\Renew-Ns2otpapp-P12.ps1 ImportSignedCert -SignedCertFile E:\OTP.cer' -ForegroundColor Yellow
            Write-Host '       b) If she uploads it DIRECTLY to the target at E:\CERTS\OTP.cer,' -ForegroundColor White
            Write-Host '          just run:' -ForegroundColor White
            Write-Host '          .\Renew-Ns2otpapp-P12.ps1 ImportSignedCert' -ForegroundColor Yellow
            Write-Host '          (script will auto-detect the cert on the target)' -ForegroundColor Gray
            Write-Host '======================================================' -ForegroundColor Cyan

            try { Add-Content $LogFile -Value '----- CSR TEXT -----'; Add-Content $LogFile -Value $csrText } catch {}
        }

        'ImportSignedCert' {
            Write-L -Level HEAD -Msg 'PHASE 2 - import signed cert and swap keystore'

            $sec = Get-Secret -Path $SecretFile
            if (-not $sec) { throw "Secret file missing at $SecretFile. Phase 1 did not complete on this jump server." }
            $pw = ConvertTo-Plain $sec.Secret

            if ($SignedCertFile) {
                if (-not (Test-Path $SignedCertFile)) { throw "Signed cert not found at -SignedCertFile: $SignedCertFile" }
                Write-L -Level INFO -Msg ("Pushing signed cert from jump server to target: {0} -> {1}" -f $SignedCertFile,$CerPath)
                Copy-Item -Path $SignedCertFile -Destination $CerPath -ToSession $session -Force
            } else {
                $candidates = @($CerPath,'E:\CERTS\OTP.cer','E:\OTP.cer','C:\CERTS\OTP.cer')
                $found = $null
                foreach ($c in $candidates) {
                    $exists = Invoke-OnTarget -Session $session -Script { param($p) Test-Path $p } -Args @($c)
                    if ($exists) { $found = $c; break }
                }
                if (-not $found) { throw "No -SignedCertFile passed and no cert found on target. Checked: $($candidates -join ', '). Either pass -SignedCertFile, or have Rhonda upload the signed .cer to the target (E:\CERTS\OTP.cer)." }
                Write-L -Level OK -Msg ("Auto-detected signed cert on target: {0}" -f $found)
                if ($found -ne $CerPath) {
                    Invoke-OnTarget -Session $session -Script { param($s,$d) Copy-Item $s $d -Force } -Args @($found,$CerPath)
                    Write-L -Level INFO -Msg ("Normalized path for import: {0}" -f $CerPath)
                }
            }

            $certInfo = Invoke-OnTarget -Session $session -Script {
                param($kt,$cer)
                $out = & $kt -printcert -file $cer 2>&1
                [pscustomobject]@{ Code=$LASTEXITCODE; Out=($out | Out-String) }
            } -Args @($keytool,$CerPath)
            if ($certInfo.Code -eq 0) {
                Add-Content $LogFile -Value "----- INCOMING CERT DETAILS -----`n$($certInfo.Out)"
                if ($certInfo.Out -match 'Owner:\s*([^\r\n]+)')     { Write-L -Level INFO -Msg ("Incoming Subject: {0}" -f $Matches[1].Trim()) }
                if ($certInfo.Out -match 'Issuer:\s*([^\r\n]+)')    { Write-L -Level INFO -Msg ("Incoming Issuer : {0}" -f $Matches[1].Trim()) }
                if ($certInfo.Out -match 'Valid from:[^\r\n]*until:\s*([^\r\n]+)') { Write-L -Level INFO -Msg ("Incoming Valid until: {0}" -f $Matches[1].Trim()) }
            }

            Write-L -Level INFO -Msg 'keytool -importcert -trustcacerts -noprompt (install reply)'
            $imp = Invoke-OnTarget -Session $session -Script {
                param($kt,$ks,$alias,$pw,$cer)
                $a = @('-importcert','-keystore',$ks,'-storetype','PKCS12','-storepass',$pw,
                       '-alias',$alias,'-file',$cer,'-trustcacerts','-noprompt')
                $out = & $kt @a 2>&1
                [pscustomobject]@{ Code=$LASTEXITCODE; Out=($out | Out-String) }
            } -Args @($keytool,$NewKeystorePath,$Alias,$pw,$CerPath)
            Add-Content $LogFile -Value $imp.Out
            if ($imp.Code -ne 0) {
                Write-L -Level ERROR -Msg 'keytool import failed. Most common cause: CA chain not trusted on the target JRE cacerts.'
                throw ("keytool importcert failed (exit {0}). See log. Ask Rhonda for a PKCS#7 bundle (.p7b) or concatenated PEM chain." -f $imp.Code)
            }
            Write-L -Level OK -Msg 'Signed cert imported into new keystore.'

            Write-L -Level INFO -Msg 'Verifying keystore entry...'
            $lst = Invoke-OnTarget -Session $session -Script {
                param($kt,$ks,$alias,$pw)
                $a = @('-list','-v','-keystore',$ks,'-storetype','PKCS12','-storepass',$pw,'-alias',$alias)
                $out = & $kt @a 2>&1
                [pscustomobject]@{ Code=$LASTEXITCODE; Out=($out | Out-String) }
            } -Args @($keytool,$NewKeystorePath,$Alias,$pw)
            Add-Content $LogFile -Value $lst.Out
            if ($lst.Code -ne 0) { throw ("keytool list failed. See log.") }
            if ($lst.Out -match 'Valid from:[^\r\n]*until:\s*([^\r\n]+)') {
                Write-L -Level OK -Msg ("New cert valid until: {0}" -f $Matches[1].Trim())
            }
            $chainLen = ([regex]::Matches($lst.Out,'Certificate\[\d+\]:')).Count
            Write-L -Level INFO -Msg ("Chain length in keystore: {0}" -f $chainLen)
            if ($chainLen -lt 2) {
                Write-L -Level WARN -Msg 'Only leaf cert in chain - app may fail TLS if clients need intermediates. Ask Rhonda for full chain.'
            }

            if (-not $AutoSwap -and -not $Force) {
                Write-L -Level HEAD -Msg 'Swap new keystore into production path?'
                Write-Host ("  Production : {0}" -f $KeystorePath)
                Write-Host ("  Ready      : {0}" -f $NewKeystorePath)
                Write-Host ("  Service    : {0}" -f $ServiceName)
                Write-Host ''
                Write-Host 'On swap this script will: stop service, backup old, copy new to prod path,' -ForegroundColor Gray
                Write-Host 'start service, verify it comes up. Auto-rollback if it does not start.' -ForegroundColor Gray
                $ans = Read-Host 'Proceed with swap? (yes/no)'
                if ($ans -ne 'yes') {
                    Write-L -Level WARN -Msg 'Swap skipped. New keystore sits at -NewKeystorePath. Update OpenText config manually if needed.'
                    return
                }
            }

            Write-L -Level INFO -Msg ("Stopping service: {0}" -f $ServiceName)
            $stop = Invoke-OnTarget -Session $session -Script {
                param($svc)
                try {
                    $s = Get-Service -Name $svc -ErrorAction Stop
                    if ($s.Status -ne 'Stopped') { Stop-Service -Name $svc -Force -ErrorAction Stop }
                    return @{ Ok=$true }
                } catch { return @{ Ok=$false; Msg=$_.Exception.Message } }
            } -Args @($ServiceName)
            if (-not $stop.Ok) { throw ("Service stop failed: {0}" -f $stop.Msg) }

            $bkProd = "$KeystorePath.$runId.preswap.bak"
            $swap = Invoke-OnTarget -Session $session -Script {
                param($prod,$new,$bk)
                try {
                    if (Test-Path $prod) { Copy-Item $prod $bk -Force; Remove-Item $prod -Force }
                    Copy-Item $new $prod -Force
                    return @{ Ok=$true }
                } catch { return @{ Ok=$false; Msg=$_.Exception.Message } }
            } -Args @($KeystorePath,$NewKeystorePath,$bkProd)
            if (-not $swap.Ok) { throw ("Swap failed: {0}" -f $swap.Msg) }
            Write-L -Level OK -Msg ("Swap complete. Pre-swap backup: {0}" -f $bkProd)

            Write-L -Level INFO -Msg 'Starting service and waiting 15s...'
            $start = Invoke-OnTarget -Session $session -Script {
                param($svc)
                try {
                    Start-Service -Name $svc -ErrorAction Stop
                    Start-Sleep -Seconds 15
                    $s = Get-Service -Name $svc
                    return @{ Ok=($s.Status -eq 'Running'); Status=[string]$s.Status }
                } catch { return @{ Ok=$false; Msg=$_.Exception.Message; Status='ErrorOnStart' } }
            } -Args @($ServiceName)

            if (-not $start.Ok) {
                Write-L -Level ERROR -Msg ("Service did NOT start cleanly after swap (status={0}). Rolling back..." -f $start.Status)
                $rb = Invoke-OnTarget -Session $session -Script {
                    param($prod,$bk,$svc)
                    try {
                        Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
                        if (Test-Path $bk) { Copy-Item $bk $prod -Force }
                        Start-Service -Name $svc -ErrorAction Stop
                        Start-Sleep -Seconds 10
                        $s = Get-Service -Name $svc
                        return @{ Ok=($s.Status -eq 'Running'); Status=[string]$s.Status }
                    } catch { return @{ Ok=$false; Msg=$_.Exception.Message } }
                } -Args @($KeystorePath,$bkProd,$ServiceName)
                if ($rb.Ok) {
                    throw "Auto-rollback SUCCESS. Service is back on the old cert. Most likely the OpenText service config has the OLD keystore password hardcoded - new password needs updating in OpenText config before retrying swap. Check E:\OTC\OpenText\Core Archive Connector\ config files for keystore.password / ssl.keystore settings."
                } else {
                    throw ("Auto-rollback FAILED: {0}. Service is DOWN. Run: .\Renew-Ns2otpapp-P12.ps1 Rollback -Force" -f $rb.Msg)
                }
            }

            Write-L -Level OK -Msg ("Service running: {0}" -f $ServiceName)
            Write-L -Level HEAD -Msg 'PHASE 2 COMPLETE - certificate deployed.'
            Write-Host ''
            Write-Host ("Deployed cert expires: {0}" -f ($Matches[1] | Out-String).Trim()) -ForegroundColor Green
            Write-Host 'Run: .\Renew-Ns2otpapp-P12.ps1 Verify   to re-check any time.' -ForegroundColor Cyan
        }

        'Verify' {
            Write-L -Level HEAD -Msg 'VERIFY keystore(s) + service'
            $sec = Get-Secret -Path $SecretFile
            $pw = $null
            if ($sec) { $pw = ConvertTo-Plain $sec.Secret }
            else { Write-L -Level WARN -Msg 'No secret file; keystore contents cannot be listed without password.' }

            foreach ($k in @($KeystorePath,$NewKeystorePath)) {
                $exists = Invoke-OnTarget -Session $session -Script { param($p) Test-Path $p } -Args @($k)
                if (-not $exists) { Write-L -Level INFO -Msg ("absent : {0}" -f $k); continue }
                if (-not $pw) { Write-L -Level INFO -Msg ("present (password unknown): {0}" -f $k); continue }
                $r = Invoke-OnTarget -Session $session -Script {
                    param($kt,$ks,$pw)
                    $out = & $kt -list -v -keystore $ks -storetype PKCS12 -storepass $pw 2>&1
                    [pscustomobject]@{ Code=$LASTEXITCODE; Out=($out | Out-String) }
                } -Args @($keytool,$k,$pw)
                Write-L -Level INFO -Msg ("--- {0} ---" -f $k)
                if ($r.Code -ne 0) { Write-L -Level WARN -Msg 'keytool -list failed (wrong password?)'; continue }
                if ($r.Out -match 'Valid from:[^\r\n]*until:\s*([^\r\n]+)') { Write-L -Level OK -Msg ("Valid until: {0}" -f $Matches[1].Trim()) }
                Add-Content $LogFile -Value $r.Out
            }
            $svc = Invoke-OnTarget -Session $session -Script { param($n) Get-Service -Name $n -ErrorAction SilentlyContinue | Select-Object Name,Status,StartType } -Args @($ServiceName)
            if ($svc) { Write-L -Level INFO -Msg ("Service {0}: Status={1} StartType={2}" -f $svc.Name,$svc.Status,$svc.StartType) }
            else { Write-L -Level WARN -Msg ("Service not found on target: {0}" -f $ServiceName) }
        }

        'Rollback' {
            Write-L -Level HEAD -Msg 'ROLLBACK to last pre-swap backup'
            $bks = Invoke-OnTarget -Session $session -Script {
                param($dir)
                Get-ChildItem -Path $dir -File -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -match '\.preswap\.bak$' -or $_.Name -match '\.bak$' } |
                    Sort-Object LastWriteTime -Descending |
                    Select-Object -First 8 |
                    Select-Object FullName,LastWriteTime,Length
            } -Args @((Split-Path -Parent $KeystorePath))
            if (-not $bks) { Write-L -Level WARN -Msg 'No backup files found on target.'; return }
            Write-L -Level INFO -Msg 'Available backups (newest first):'
            $i = 0
            foreach ($b in $bks) { Write-Host ("  [{0}] {1}  ({2}, {3} bytes)" -f $i,$b.FullName,$b.LastWriteTime,$b.Length); $i++ }
            $pick = Read-Host 'Enter index to restore (empty to cancel)'
            if ([string]::IsNullOrWhiteSpace($pick)) { Write-L -Level INFO -Msg 'Cancelled.'; return }
            $idx = [int]$pick
            $chosen = $bks[$idx].FullName
            Write-L -Level WARN -Msg ("Restoring {0} -> {1}, restarting service..." -f $chosen,$KeystorePath)
            $r = Invoke-OnTarget -Session $session -Script {
                param($src,$dst,$svc)
                try {
                    Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
                    Copy-Item $src $dst -Force
                    Start-Service -Name $svc -ErrorAction Stop
                    Start-Sleep -Seconds 10
                    $s = Get-Service -Name $svc
                    return @{ Ok=($s.Status -eq 'Running'); Status=[string]$s.Status }
                } catch { return @{ Ok=$false; Msg=$_.Exception.Message } }
            } -Args @($chosen,$KeystorePath,$ServiceName)
            if ($r.Ok) { Write-L -Level OK -Msg 'Rollback complete, service running.' }
            else { throw ("Rollback failed: {0}" -f $r.Msg) }
        }
    }
}
catch {
    Write-L -Level ERROR -Msg $_.Exception.Message
    throw
}
finally {
    if ($session) { Remove-PSSession $session -ErrorAction SilentlyContinue }
    Write-L -Level INFO -Msg ("Run artifacts: {0}" -f $LogFile)
}

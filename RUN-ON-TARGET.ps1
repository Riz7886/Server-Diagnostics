#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Position=0)]
    [ValidateSet('Csr','Import','Verify','Rollback')]
    [string]$Phase = 'Csr',

    [string]$TargetIp   = '10.168.0.32',
    [string]$TargetFqdn = 'ns2otpapp.sap.parker.corp',

    [string]$KeystorePath    = 'E:\CERTS\s4pceotcac.p12',
    [string]$NewKeystorePath = 'E:\CERTS\s4pceotcac_new.p12',
    [string]$CsrPath         = 'E:\CERTS\OTP.csr',
    [string]$CerPath         = 'E:\CERTS\OTP.cer',

    [string]$Alias   = 's4hpce',
    [string]$Subject = 'CN=ns2otpapp.sap.parker.corp, OU=COR, O=Parker Hannifin, L=Cleaveland, C=US',
    [string]$SanDns  = 'ns2otpapp.sap.parker.corp',
    [int]$KeySize    = 2048,
    [string]$ServiceName = 'OpenText Core Archive Connector',

    [string]$SignedCert,
    [string]$Username = 'c5406751',
    [pscredential]$Credential,

    [string]$WorkDir = (Join-Path $PSScriptRoot 'work')
)

$ErrorActionPreference = 'Stop'

function Say {
    param([string]$Msg, [string]$Color = 'Cyan')
    Write-Host "[*] $Msg" -ForegroundColor $Color
}

function Connect-Target {
    param(
        [string]$Ip,
        [string]$Fqdn,
        [pscredential]$Cred
    )
    $sslOpt = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck

    $attempts = @(
        @{ n = 'HTTPS (port 5986) to FQDN';   sb = { New-PSSession -ComputerName $Fqdn -Credential $Cred -UseSSL -SessionOption $sslOpt -Authentication Negotiate -ErrorAction Stop } },
        @{ n = 'HTTPS (port 5986) to IP';     sb = { New-PSSession -ComputerName $Ip   -Credential $Cred -UseSSL -SessionOption $sslOpt -Authentication Negotiate -ErrorAction Stop } },
        @{ n = 'Kerberos to FQDN';            sb = { New-PSSession -ComputerName $Fqdn -Credential $Cred -Authentication Kerberos -ErrorAction Stop } },
        @{ n = 'Negotiate to FQDN';           sb = { New-PSSession -ComputerName $Fqdn -Credential $Cred -Authentication Negotiate -ErrorAction Stop } },
        @{ n = 'Default (NTLM) to IP';        sb = { New-PSSession -ComputerName $Ip   -Credential $Cred -ErrorAction Stop } }
    )

    foreach ($a in $attempts) {
        try {
            Say "Trying: $($a.n)" 'Cyan'
            $s = & $a.sb
            Say "Connected via: $($a.n)" 'Green'
            return $s
        } catch {
            $first = ($_.Exception.Message -split "`n")[0].Trim()
            Say "FAILED: $($a.n) -- $first" 'Yellow'
        }
    }
    return $null
}

if (-not (Test-Path $WorkDir)) { New-Item -Path $WorkDir -ItemType Directory -Force | Out-Null }
$pwStore  = Join-Path $WorkDir 'pw.xml'
$csrLocal = Join-Path $WorkDir 'OTP.csr'

if (-not $Credential) {
    Say "Credential prompt. If plain fails try DIRECTORY\\$Username or CRE\\$Username" 'Yellow'
    $Credential = Get-Credential -UserName $Username -Message "Windows credential for $TargetIp"
}
if (-not $Credential) { Say 'No credential' 'Red'; exit 1 }

Say "Phase: $Phase" 'Cyan'
$session = Connect-Target -Ip $TargetIp -Fqdn $TargetFqdn -Cred $Credential
if (-not $session) {
    Say "All PSSession attempts failed." 'Red'
    Say "Run  Test-NetConnection $TargetIp -Port 5985  and  Test-NetConnection $TargetIp -Port 5986  to see what's open." 'Yellow'
    Say "Or fall back to the SMB/DCOM version (RENEW.ps1 or Renew-Ns2otpapp-P12-SMB.ps1)." 'Yellow'
    exit 1
}

$preflight = Invoke-Command -Session $session -ScriptBlock {
    $kt = $null
    $candidates = @(
        'E:\OTC\OpenText\Core Archive Connector\jre\bin\keytool.exe',
        'E:\OTC\OpenText\*\jre\bin\keytool.exe',
        'E:\OTC\*\jre\bin\keytool.exe',
        'E:\OpenText\Core Archive Connector\jre\bin\keytool.exe',
        'E:\OpenText\*\jre\bin\keytool.exe',
        'E:\OpenText\*\bin\keytool.exe',
        'E:\Install\*\jre\bin\keytool.exe',
        'C:\Program Files\Java\jre*\bin\keytool.exe',
        'C:\Program Files\Java\jdk*\bin\keytool.exe',
        'C:\Program Files\Eclipse Adoptium\*\bin\keytool.exe',
        'C:\Program Files\OpenJDK\*\bin\keytool.exe',
        'C:\Program Files (x86)\Java\*\bin\keytool.exe',
        'D:\Java\*\bin\keytool.exe',
        'E:\Java\*\bin\keytool.exe',
        'C:\OTP\*\bin\keytool.exe'
    )
    foreach ($p in $candidates) {
        $f = Get-ChildItem -Path $p -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($f) { $kt = $f.FullName; break }
    }
    if (-not $kt) {
        $deep = Get-ChildItem -Path 'E:\' -Recurse -Filter 'keytool.exe' -ErrorAction SilentlyContinue -Force | Select-Object -First 1
        if ($deep) { $kt = $deep.FullName }
    }
    $listing = Get-ChildItem -Path 'E:\CERTS' -ErrorAction SilentlyContinue | Select-Object Name, Length, LastWriteTime | Sort-Object Name
    $svc = Get-Service -Name 'OpenText Core Archive Connector' -ErrorAction SilentlyContinue
    [pscustomobject]@{
        Keytool = $kt
        CertsDir = $listing
        Service = if ($svc) { $svc.Status.ToString() } else { 'NotFound' }
    }
}

Say "Target keytool: $(if ($preflight.Keytool) { $preflight.Keytool } else { 'NOT FOUND' })" $(if ($preflight.Keytool) { 'Green' } else { 'Red' })
Say "Target service: $($preflight.Service)" 'Green'
Write-Host ''
Write-Host 'E:\CERTS contents on target:' -ForegroundColor Yellow
$preflight.CertsDir | Format-Table -AutoSize | Out-String | Write-Host
if (-not $preflight.Keytool) {
    Say "Cannot proceed without keytool on target. Check Java install path." 'Red'
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    exit 1
}
$script:TargetKeytool = $preflight.Keytool

try {
    switch ($Phase) {

        'Csr' {
            $pw  = Read-Host -AsSecureString 'NEW keystore password'
            $pw2 = Read-Host -AsSecureString 'Confirm password'
            $p1 = [System.Net.NetworkCredential]::new('', $pw).Password
            $p2 = [System.Net.NetworkCredential]::new('', $pw2).Password
            if ($p1 -ne $p2) { Say 'Passwords do not match' 'Red'; exit 1 }
            if ($p1.Length -lt 6) { Say 'Password too short' 'Red'; exit 1 }
            $pw | Export-Clixml -Path $pwStore
            Say "Password saved (DPAPI): $pwStore" 'Green'

            $out = Invoke-Command -Session $session -ArgumentList $script:TargetKeytool, $Alias, $Subject, $SanDns, $KeySize, $NewKeystorePath, $CsrPath, $p1 -ScriptBlock {
                param($kt, $Alias, $Subject, $SanDns, $KeySize, $NewKeystorePath, $CsrPath, $Pw)
                $ErrorActionPreference = 'Stop'

                $certsDir = Split-Path $NewKeystorePath -Parent
                if (-not (Test-Path $certsDir)) { New-Item -Path $certsDir -ItemType Directory -Force | Out-Null }
                $stamp = Get-Date -Format 'yyyyMMddHHmmss'
                if (Test-Path $NewKeystorePath) { Move-Item $NewKeystorePath "$NewKeystorePath.old-$stamp" -Force }
                if (Test-Path $CsrPath)         { Move-Item $CsrPath         "$CsrPath.old-$stamp"         -Force }

                $gen = @(
                    '-genkeypair','-alias',$Alias,'-keyalg','RSA','-keysize',"$KeySize",
                    '-dname',$Subject,'-ext',"SAN=dns:$SanDns",
                    '-keystore',$NewKeystorePath,'-storetype','PKCS12',
                    '-storepass',$Pw,'-keypass',$Pw,'-validity','730'
                )
                & $kt @gen
                if ($LASTEXITCODE -ne 0) { throw "keytool -genkeypair failed ($LASTEXITCODE)" }

                $req = @(
                    '-certreq','-alias',$Alias,
                    '-keystore',$NewKeystorePath,'-storetype','PKCS12',
                    '-storepass',$Pw,'-file',$CsrPath,'-ext',"SAN=dns:$SanDns"
                )
                & $kt @req
                if ($LASTEXITCODE -ne 0) { throw "keytool -certreq failed ($LASTEXITCODE)" }

                $csrText = Get-Content $CsrPath -Raw
                return [pscustomobject]@{
                    Keytool   = $kt
                    Keystore  = $NewKeystorePath
                    CsrPath   = $CsrPath
                    CsrText   = $csrText
                }
            }

            Say "Keytool on target: $($out.Keytool)" 'Green'
            Say "Keystore on target: $($out.Keystore)" 'Green'
            Say "CSR on target: $($out.CsrPath)" 'Green'
            Set-Content -Path $csrLocal -Value $out.CsrText -Encoding ASCII
            Say "CSR saved locally: $csrLocal" 'Green'
            Write-Host ''
            Write-Host '=========== CSR BELOW (email to Rhonda) ===========' -ForegroundColor Yellow
            Write-Host $out.CsrText
            Write-Host '===================================================' -ForegroundColor Yellow
            Write-Host ''
            Say "Phase 1 done. Tomorrow: .\RUN-ON-TARGET.ps1 Import -SignedCert C:\path\OTP.cer" 'Cyan'
        }

        'Import' {
            if (-not (Test-Path $pwStore)) { Say 'Password store missing - re-run Csr first' 'Red'; exit 1 }
            $pw = Import-Clixml -Path $pwStore
            $p1 = [System.Net.NetworkCredential]::new('', $pw).Password

            if ($SignedCert) {
                if (-not (Test-Path $SignedCert)) { Say "Not found: $SignedCert" 'Red'; exit 1 }
                $bytes = [IO.File]::ReadAllBytes($SignedCert)
                Invoke-Command -Session $session -ArgumentList $CerPath, $bytes -ScriptBlock {
                    param($CerPath, $Bytes)
                    $d = Split-Path $CerPath -Parent
                    if (-not (Test-Path $d)) { New-Item -Path $d -ItemType Directory -Force | Out-Null }
                    [IO.File]::WriteAllBytes($CerPath, $Bytes)
                }
                Say "Signed cert pushed to target: $CerPath" 'Green'
            }

            $out = Invoke-Command -Session $session -ArgumentList $script:TargetKeytool, $Alias, $KeystorePath, $NewKeystorePath, $CerPath, $ServiceName, $p1 -ScriptBlock {
                param($kt, $Alias, $KeystorePath, $NewKeystorePath, $CerPath, $ServiceName, $Pw)
                $ErrorActionPreference = 'Stop'
                if (-not (Test-Path $CerPath))         { throw "Signed cert not found: $CerPath" }
                if (-not (Test-Path $NewKeystorePath)) { throw "New keystore missing: $NewKeystorePath" }

                $imp = @(
                    '-importcert','-trustcacerts','-noprompt',
                    '-alias',$Alias,'-file',$CerPath,
                    '-keystore',$NewKeystorePath,'-storetype','PKCS12','-storepass',$Pw
                )
                & $kt @imp
                if ($LASTEXITCODE -ne 0) { throw "keytool -importcert failed ($LASTEXITCODE)" }

                $svc = Get-Service -Name $ServiceName -ErrorAction Stop
                $wasRunning = $svc.Status -eq 'Running'
                if ($wasRunning) {
                    Stop-Service -Name $ServiceName -Force
                    $svc.WaitForStatus('Stopped','00:01:30')
                }

                $stamp = Get-Date -Format 'yyyyMMddHHmmss'
                $backupPath = "$KeystorePath.bak-$stamp"
                if (Test-Path $KeystorePath) { Copy-Item $KeystorePath $backupPath -Force }

                Copy-Item $NewKeystorePath $KeystorePath -Force

                Start-Service -Name $ServiceName
                Start-Sleep -Seconds 6
                $svc = Get-Service -Name $ServiceName
                $rolledBack = $false
                if ($svc.Status -ne 'Running') {
                    if (Test-Path $backupPath) {
                        try { Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue } catch { }
                        Copy-Item $backupPath $KeystorePath -Force
                        Start-Service -Name $ServiceName -ErrorAction SilentlyContinue
                        $rolledBack = $true
                    }
                    throw "Service failed to start after swap. Rolled back from $backupPath"
                }

                return [pscustomobject]@{
                    Keystore      = $KeystorePath
                    Backup        = $backupPath
                    ServiceStatus = $svc.Status.ToString()
                    RolledBack    = $rolledBack
                }
            }

            Say "Production keystore: $($out.Keystore)" 'Green'
            Say "Backup:              $($out.Backup)" 'Green'
            Say "Service status:      $($out.ServiceStatus)" 'Green'
        }

        'Verify' {
            if (-not (Test-Path $pwStore)) { Say 'Password store missing - run Csr first' 'Red'; exit 1 }
            $pw = Import-Clixml -Path $pwStore
            $p1 = [System.Net.NetworkCredential]::new('', $pw).Password
            $out = Invoke-Command -Session $session -ArgumentList $script:TargetKeytool, $Alias, $KeystorePath, $ServiceName, $p1 -ScriptBlock {
                param($kt, $Alias, $KeystorePath, $ServiceName, $Pw)
                $ErrorActionPreference = 'Stop'
                $listing = & $kt -list -v -alias $Alias -keystore $KeystorePath -storetype PKCS12 -storepass $Pw 2>&1
                $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
                return [pscustomobject]@{
                    KeystoreExists = (Test-Path $KeystorePath)
                    Listing        = ($listing -join "`n")
                    ServiceStatus  = if ($svc) { $svc.Status.ToString() } else { 'NotInstalled' }
                }
            }
            Say "Keystore exists: $($out.KeystoreExists)" 'Green'
            Say "Service:         $($out.ServiceStatus)" 'Green'
            Write-Host ''
            Write-Host '=========== Cert Listing ===========' -ForegroundColor Yellow
            Write-Host $out.Listing
            Write-Host '====================================' -ForegroundColor Yellow
        }

        'Rollback' {
            $out = Invoke-Command -Session $session -ArgumentList $KeystorePath, $ServiceName -ScriptBlock {
                param($KeystorePath, $ServiceName)
                $ErrorActionPreference = 'Stop'
                $dir  = Split-Path $KeystorePath -Parent
                $name = Split-Path $KeystorePath -Leaf
                $latest = Get-ChildItem -Path $dir -Filter "$name.bak-*" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
                if (-not $latest) { throw 'No backup file found' }
                $svc = Get-Service -Name $ServiceName
                if ($svc.Status -eq 'Running') {
                    Stop-Service -Name $ServiceName -Force
                    $svc.WaitForStatus('Stopped','00:01:30')
                }
                Copy-Item $latest.FullName $KeystorePath -Force
                Start-Service -Name $ServiceName
                Start-Sleep -Seconds 6
                $svc = Get-Service -Name $ServiceName
                return [pscustomobject]@{
                    RestoredFrom  = $latest.FullName
                    ServiceStatus = $svc.Status.ToString()
                }
            }
            Say "Restored: $($out.RestoredFrom)" 'Green'
            Say "Service:  $($out.ServiceStatus)" 'Green'
        }
    }
}
finally {
    if ($session) { Remove-PSSession -Session $session -ErrorAction SilentlyContinue }
}

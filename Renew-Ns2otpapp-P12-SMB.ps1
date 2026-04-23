#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Position=0)]
    [ValidateSet('GenerateCsr','ImportSignedCert','Verify','Rollback')]
    [string]$Phase = 'GenerateCsr',

    [string]$Target = '10.168.0.32',
    [string]$RemoteDrive = 'e$',

    [string]$KeystorePath    = 'E:\CERTS\s4pceotcac.p12',
    [string]$NewKeystorePath = 'E:\CERTS\s4pceotcac_new.p12',
    [string]$CsrPath         = 'E:\CERTS\OTP.csr',
    [string]$CerPath         = 'E:\CERTS\OTP.cer',

    [string]$Alias   = 's4hpce',
    [string]$Subject = 'CN=ns2otpapp.sap.parker.corp, OU=COR, O=Parker Hannifin, L=Cleaveland, C=US',
    [string]$SanDns  = 'ns2otpapp.sap.parker.corp',
    [int]$KeySize    = 2048,
    [string]$ServiceName = 'OpenText Core Archive Connector',

    [string]$SignedCertFile,

    [string]$Username = 'c5406751',
    [pscredential]$Credential,

    [string]$LocalOutputDir  = $(Join-Path $PSScriptRoot 'output'),
    [string]$SecretStorePath = $(Join-Path $PSScriptRoot '.secrets'),
    [switch]$SavePasswords,
    [switch]$PromptForPasswords,

    [int]$TimeoutSeconds = 240
)

$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

function Write-Log {
    param([string]$Level, [string]$Message)
    $ts = Get-Date -Format 'HH:mm:ss'
    $color = switch ($Level) {
        'ERROR' { 'Red' }
        'WARN'  { 'Yellow' }
        'INFO'  { 'Cyan' }
        'OK'    { 'Green' }
        default { 'White' }
    }
    Write-Host ("[{0}] {1,-5} {2}" -f $ts, $Level, $Message) -ForegroundColor $color
}

function Get-RemoteCertShare {
    param([string]$Target, [string]$Drive, [string]$RemoteDirLocal)
    $dirSub = $RemoteDirLocal -replace '^[A-Za-z]:\\',''
    return (Join-Path ("\\$Target\$Drive") $dirSub)
}

function ConvertFrom-SecureStringPlain {
    param([System.Security.SecureString]$SecureString)
    return [System.Net.NetworkCredential]::new('', $SecureString).Password
}

function ConvertTo-B64 {
    param([string]$Plain)
    return [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Plain))
}

function Resolve-Credential {
    param(
        [string]$Username,
        [pscredential]$Provided,
        [string]$StorePath,
        [switch]$Save,
        [switch]$Prompt
    )
    if ($Provided) { return $Provided }
    $credFile = Join-Path $StorePath 'jumpserver-cred.xml'
    if ((Test-Path $credFile) -and -not $Prompt) {
        try {
            $cred = Import-Clixml -Path $credFile
            Write-Log INFO "Loaded saved credential: $($cred.UserName)"
            return $cred
        } catch {
            Write-Log WARN "Failed to load saved credential: $($_.Exception.Message)"
        }
    }
    Write-Log INFO "Prompting for credential. Try plain username first; if 1326, retry with DOMAIN\$Username."
    $cred = Get-Credential -UserName $Username -Message "Credential for $Target (DCOM/SMB)"
    if (-not $cred) { throw 'No credential provided' }
    if ($Save) {
        if (-not (Test-Path $StorePath)) { New-Item -Path $StorePath -ItemType Directory -Force | Out-Null }
        $cred | Export-Clixml -Path $credFile
        Write-Log OK "Credential saved to $credFile (DPAPI)"
    }
    return $cred
}

function Resolve-KeystorePassword {
    param(
        [string]$StorePath,
        [switch]$Save,
        [switch]$Prompt,
        [string]$Purpose = 'KeystorePassword'
    )
    $pwFile = Join-Path $StorePath "$Purpose.xml"
    if ((Test-Path $pwFile) -and -not $Prompt) {
        try {
            $ss = Import-Clixml -Path $pwFile
            Write-Log INFO "Loaded saved $Purpose from DPAPI"
            return $ss
        } catch { Write-Log WARN "Failed to load $Purpose" }
    }
    $ss = Read-Host -AsSecureString -Prompt "Enter NEW keystore password (also needed for Phase 2)"
    if ($ss.Length -eq 0) { throw 'Empty password not allowed' }
    if ($Save) {
        if (-not (Test-Path $StorePath)) { New-Item -Path $StorePath -ItemType Directory -Force | Out-Null }
        $ss | Export-Clixml -Path $pwFile
        Write-Log OK "$Purpose saved to $pwFile (DPAPI)"
    }
    return $ss
}

function Invoke-RemoteExec {
    param(
        [string]$Target,
        [pscredential]$Credential,
        [string]$CommandLine
    )
    Write-Log INFO "Launching remote process via DCOM Win32_Process.Create (no WinRM)"
    $opt = New-CimSessionOption -Protocol Dcom
    $cim = $null
    try {
        $cim = New-CimSession -ComputerName $Target -Credential $Credential -SessionOption $opt -ErrorAction Stop
        $result = Invoke-CimMethod -CimSession $cim -ClassName Win32_Process -MethodName Create -Arguments @{
            CommandLine = $CommandLine
        } -ErrorAction Stop
        if ($result.ReturnValue -ne 0) {
            throw "Win32_Process.Create returned $($result.ReturnValue) (non-zero = failure)"
        }
        Write-Log OK "Remote process launched. PID=$($result.ProcessId)"
        return $result.ProcessId
    }
    finally {
        if ($cim) { Remove-CimSession -CimSession $cim -ErrorAction SilentlyContinue }
    }
}

function Wait-ForStatus {
    param([string]$StatusPath, [int]$TimeoutSeconds)
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    Write-Log INFO "Waiting for remote runner (status: $StatusPath, timeout ${TimeoutSeconds}s)"
    while ((Get-Date) -lt $deadline) {
        if (Test-Path -Path $StatusPath) {
            Start-Sleep -Milliseconds 700
            try {
                $raw = Get-Content -Path $StatusPath -Raw -ErrorAction Stop
                if ($raw -and $raw.Trim().Length -gt 0) {
                    Write-Host ''
                    return ($raw | ConvertFrom-Json -ErrorAction Stop)
                }
            } catch { }
        }
        Start-Sleep -Seconds 2
        Write-Host '.' -NoNewline
    }
    Write-Host ''
    throw "Remote runner did not complete within $TimeoutSeconds seconds"
}

function Build-RunnerGenerateCsr {
    param($KeystorePath,$NewKeystorePath,$CsrPath,$Alias,$Subject,$SanDns,$KeySize,$PwB64)
@"
`$ErrorActionPreference = 'Stop'
`$statusPath = 'E:\CERTS\_otp_status.json'
`$logPath    = 'E:\CERTS\_otp_runner.log'
Start-Transcript -Path `$logPath -Force | Out-Null
try {
    `$pw = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('$PwB64'))
    `$candidates = @(
        'C:\Program Files\Java\jre*\bin\keytool.exe',
        'C:\Program Files\Java\jdk*\bin\keytool.exe',
        'C:\Program Files\Eclipse Adoptium\*\bin\keytool.exe',
        'C:\Program Files (x86)\Java\*\bin\keytool.exe',
        'D:\Java\*\bin\keytool.exe',
        'E:\Java\*\bin\keytool.exe',
        'C:\OTP\*\bin\keytool.exe',
        'E:\OTC\*\bin\keytool.exe'
    )
    `$keytool = `$null
    foreach (`$pat in `$candidates) {
        `$f = Get-ChildItem -Path `$pat -ErrorAction SilentlyContinue | Select-Object -First 1
        if (`$f) { `$keytool = `$f.FullName; break }
    }
    if (-not `$keytool) {
        `$javaDir = Get-ChildItem 'C:\Program Files\Java' -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 1
        if (`$javaDir) {
            `$c = Join-Path `$javaDir.FullName 'bin\keytool.exe'
            if (Test-Path `$c) { `$keytool = `$c }
        }
    }
    if (-not `$keytool) { throw 'keytool.exe not found on target (searched standard Java paths)' }
    Write-Host ('Using keytool: ' + `$keytool)

    `$certsDir = Split-Path '$NewKeystorePath' -Parent
    if (-not (Test-Path `$certsDir)) { New-Item -Path `$certsDir -ItemType Directory -Force | Out-Null }
    `$stamp = Get-Date -Format 'yyyyMMddHHmmss'
    if (Test-Path '$NewKeystorePath') { Move-Item '$NewKeystorePath' ('$NewKeystorePath.old-' + `$stamp) -Force }
    if (Test-Path '$CsrPath')         { Move-Item '$CsrPath'         ('$CsrPath.old-' + `$stamp)         -Force }

    & `$keytool -genkeypair -alias '$Alias' -keyalg RSA -keysize $KeySize ``
        -dname '$Subject' ``
        -ext ('SAN=dns:' + '$SanDns') ``
        -keystore '$NewKeystorePath' -storetype PKCS12 ``
        -storepass `$pw -keypass `$pw ``
        -validity 730
    if (`$LASTEXITCODE -ne 0) { throw "keytool -genkeypair failed with exit `$LASTEXITCODE" }

    & `$keytool -certreq -alias '$Alias' ``
        -keystore '$NewKeystorePath' -storetype PKCS12 ``
        -storepass `$pw ``
        -file '$CsrPath' ``
        -ext ('SAN=dns:' + '$SanDns')
    if (`$LASTEXITCODE -ne 0) { throw "keytool -certreq failed with exit `$LASTEXITCODE" }

    `$csrContent = Get-Content '$CsrPath' -Raw

    @{
        status       = 'Success'
        phase        = 'GenerateCsr'
        keystorePath = '$NewKeystorePath'
        csrPath      = '$CsrPath'
        csrContent   = `$csrContent
        keytool      = `$keytool
        timestamp    = (Get-Date).ToString('o')
    } | ConvertTo-Json -Depth 4 | Set-Content -Path `$statusPath -Encoding UTF8
}
catch {
    @{
        status     = 'Failed'
        phase      = 'GenerateCsr'
        error      = `$_.Exception.Message
        stackTrace = `$_.ScriptStackTrace
        timestamp  = (Get-Date).ToString('o')
    } | ConvertTo-Json -Depth 4 | Set-Content -Path `$statusPath -Encoding UTF8
}
finally {
    Stop-Transcript | Out-Null
    try { Remove-Item `$MyInvocation.MyCommand.Path -Force -ErrorAction SilentlyContinue } catch { }
}
"@
}

function Build-RunnerImportSignedCert {
    param($KeystorePath,$NewKeystorePath,$CerPath,$Alias,$ServiceName,$PwB64)
@"
`$ErrorActionPreference = 'Stop'
`$statusPath = 'E:\CERTS\_otp_status.json'
`$logPath    = 'E:\CERTS\_otp_import.log'
Start-Transcript -Path `$logPath -Force | Out-Null
`$stamp = Get-Date -Format 'yyyyMMddHHmmss'
`$backupPath = '$KeystorePath' + '.bak-' + `$stamp
`$rolledBack = `$false
try {
    `$pw = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('$PwB64'))
    `$candidates = @(
        'C:\Program Files\Java\jre*\bin\keytool.exe',
        'C:\Program Files\Java\jdk*\bin\keytool.exe',
        'C:\Program Files\Eclipse Adoptium\*\bin\keytool.exe',
        'C:\Program Files (x86)\Java\*\bin\keytool.exe',
        'D:\Java\*\bin\keytool.exe',
        'E:\Java\*\bin\keytool.exe',
        'C:\OTP\*\bin\keytool.exe',
        'E:\OTC\*\bin\keytool.exe'
    )
    `$keytool = `$null
    foreach (`$pat in `$candidates) {
        `$f = Get-ChildItem -Path `$pat -ErrorAction SilentlyContinue | Select-Object -First 1
        if (`$f) { `$keytool = `$f.FullName; break }
    }
    if (-not `$keytool) { throw 'keytool.exe not found on target' }

    if (-not (Test-Path '$CerPath'))         { throw ('Signed cert not found at ' + '$CerPath') }
    if (-not (Test-Path '$NewKeystorePath')) { throw ('New keystore missing: ' + '$NewKeystorePath' + ' (run GenerateCsr first)') }

    & `$keytool -importcert -trustcacerts -noprompt ``
        -alias '$Alias' ``
        -file '$CerPath' ``
        -keystore '$NewKeystorePath' -storetype PKCS12 ``
        -storepass `$pw
    if (`$LASTEXITCODE -ne 0) { throw "keytool -importcert failed with exit `$LASTEXITCODE" }

    `$listOut = & `$keytool -list -v -alias '$Alias' ``
        -keystore '$NewKeystorePath' -storetype PKCS12 ``
        -storepass `$pw 2>&1
    `$listOut | Set-Content -Path 'E:\CERTS\_otp_list.txt' -Encoding UTF8
    if (`$LASTEXITCODE -ne 0) { throw "keytool -list failed after import" }

    `$svc = Get-Service -Name '$ServiceName' -ErrorAction Stop
    `$wasRunning = `$svc.Status -eq 'Running'
    if (`$wasRunning) {
        Stop-Service -Name '$ServiceName' -Force -ErrorAction Stop
        `$svc.WaitForStatus('Stopped','00:01:30')
    }

    if (Test-Path '$KeystorePath') { Copy-Item '$KeystorePath' `$backupPath -Force }

    Copy-Item '$NewKeystorePath' '$KeystorePath' -Force

    Start-Service -Name '$ServiceName' -ErrorAction Stop
    Start-Sleep -Seconds 6
    `$svc = Get-Service -Name '$ServiceName'
    if (`$svc.Status -ne 'Running') {
        if (Test-Path `$backupPath) {
            try { Stop-Service -Name '$ServiceName' -Force -ErrorAction SilentlyContinue } catch { }
            Copy-Item `$backupPath '$KeystorePath' -Force
            Start-Service -Name '$ServiceName' -ErrorAction SilentlyContinue
            `$rolledBack = `$true
        }
        throw ('Service failed to start after cert swap — rolled back to ' + `$backupPath)
    }

    @{
        status        = 'Success'
        phase         = 'ImportSignedCert'
        keystorePath  = '$KeystorePath'
        backupPath    = `$backupPath
        serviceStatus = `$svc.Status.ToString()
        timestamp     = (Get-Date).ToString('o')
    } | ConvertTo-Json -Depth 4 | Set-Content -Path `$statusPath -Encoding UTF8
}
catch {
    @{
        status      = 'Failed'
        phase       = 'ImportSignedCert'
        error       = `$_.Exception.Message
        stackTrace  = `$_.ScriptStackTrace
        backupPath  = `$backupPath
        rolledBack  = `$rolledBack
        timestamp   = (Get-Date).ToString('o')
    } | ConvertTo-Json -Depth 4 | Set-Content -Path `$statusPath -Encoding UTF8
}
finally {
    Stop-Transcript | Out-Null
    try { Remove-Item `$MyInvocation.MyCommand.Path -Force -ErrorAction SilentlyContinue } catch { }
}
"@
}

function Build-RunnerVerify {
    param($KeystorePath,$Alias,$ServiceName,$PwB64)
@"
`$ErrorActionPreference = 'Stop'
`$statusPath = 'E:\CERTS\_otp_status.json'
try {
    `$pw = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('$PwB64'))
    `$candidates = @(
        'C:\Program Files\Java\jre*\bin\keytool.exe',
        'C:\Program Files\Java\jdk*\bin\keytool.exe',
        'C:\Program Files\Eclipse Adoptium\*\bin\keytool.exe',
        'C:\Program Files (x86)\Java\*\bin\keytool.exe',
        'D:\Java\*\bin\keytool.exe',
        'E:\Java\*\bin\keytool.exe'
    )
    `$keytool = `$null
    foreach (`$pat in `$candidates) {
        `$f = Get-ChildItem -Path `$pat -ErrorAction SilentlyContinue | Select-Object -First 1
        if (`$f) { `$keytool = `$f.FullName; break }
    }
    if (-not `$keytool) { throw 'keytool.exe not found' }

    `$listOut = & `$keytool -list -v -alias '$Alias' -keystore '$KeystorePath' -storetype PKCS12 -storepass `$pw 2>&1
    `$svc = Get-Service -Name '$ServiceName' -ErrorAction SilentlyContinue

    @{
        status         = 'Success'
        phase          = 'Verify'
        keystorePath   = '$KeystorePath'
        keystoreExists = (Test-Path '$KeystorePath')
        certInfo       = (`$listOut -join "``n")
        serviceStatus  = if (`$svc) { `$svc.Status.ToString() } else { 'NotInstalled' }
        timestamp      = (Get-Date).ToString('o')
    } | ConvertTo-Json -Depth 4 | Set-Content -Path `$statusPath -Encoding UTF8
}
catch {
    @{
        status    = 'Failed'
        phase     = 'Verify'
        error     = `$_.Exception.Message
        timestamp = (Get-Date).ToString('o')
    } | ConvertTo-Json -Depth 4 | Set-Content -Path `$statusPath -Encoding UTF8
}
finally {
    try { Remove-Item `$MyInvocation.MyCommand.Path -Force -ErrorAction SilentlyContinue } catch { }
}
"@
}

function Build-RunnerRollback {
    param($KeystorePath,$ServiceName)
@"
`$ErrorActionPreference = 'Stop'
`$statusPath = 'E:\CERTS\_otp_status.json'
try {
    `$dir = Split-Path '$KeystorePath' -Parent
    `$name = Split-Path '$KeystorePath' -Leaf
    `$latest = Get-ChildItem -Path `$dir -Filter ("`$name" + '.bak-*') -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not `$latest) { throw 'No backup file found to roll back to' }
    `$svc = Get-Service -Name '$ServiceName'
    if (`$svc.Status -eq 'Running') {
        Stop-Service -Name '$ServiceName' -Force
        `$svc.WaitForStatus('Stopped','00:01:30')
    }
    Copy-Item `$latest.FullName '$KeystorePath' -Force
    Start-Service -Name '$ServiceName'
    Start-Sleep -Seconds 6
    `$svc = Get-Service -Name '$ServiceName'
    @{
        status        = 'Success'
        phase         = 'Rollback'
        restoredFrom  = `$latest.FullName
        serviceStatus = `$svc.Status.ToString()
        timestamp     = (Get-Date).ToString('o')
    } | ConvertTo-Json -Depth 4 | Set-Content -Path `$statusPath -Encoding UTF8
}
catch {
    @{
        status    = 'Failed'
        phase     = 'Rollback'
        error     = `$_.Exception.Message
        timestamp = (Get-Date).ToString('o')
    } | ConvertTo-Json -Depth 4 | Set-Content -Path `$statusPath -Encoding UTF8
}
finally {
    try { Remove-Item `$MyInvocation.MyCommand.Path -Force -ErrorAction SilentlyContinue } catch { }
}
"@
}

function Invoke-PhaseSmb {
    param(
        [string]$Target,
        [string]$RemoteDrive,
        [string]$RemoteCertDir,
        [pscredential]$Credential,
        [string]$RunnerCode,
        [int]$TimeoutSeconds
    )
    $share = Get-RemoteCertShare -Target $Target -Drive $RemoteDrive -RemoteDirLocal $RemoteCertDir
    if (-not (Test-Path $share)) {
        try { New-Item -Path $share -ItemType Directory -Force | Out-Null }
        catch { throw "Cannot access or create $share : $($_.Exception.Message)" }
    }

    $runnerRemote = Join-Path $share '_otp_runner.ps1'
    $statusRemote = Join-Path $share '_otp_status.json'

    Remove-Item $statusRemote -Force -ErrorAction SilentlyContinue
    Remove-Item $runnerRemote -Force -ErrorAction SilentlyContinue

    Set-Content -Path $runnerRemote -Value $RunnerCode -Encoding UTF8
    Write-Log OK "Runner pushed to $runnerRemote"

    $cmdLine = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $RemoteCertDir + '\_otp_runner.ps1"'
    [void](Invoke-RemoteExec -Target $Target -Credential $Credential -CommandLine $cmdLine)

    $status = Wait-ForStatus -StatusPath $statusRemote -TimeoutSeconds $TimeoutSeconds
    return @{ Status = $status; StatusPath = $statusRemote; RunnerPath = $runnerRemote; SharePath = $share }
}

Write-Log INFO "Phase: $Phase"
Write-Log INFO "Target: $Target  (SMB + DCOM path — no WinRM)"

if (-not (Test-Path $LocalOutputDir))  { New-Item -Path $LocalOutputDir  -ItemType Directory -Force | Out-Null }
if (-not (Test-Path $SecretStorePath)) { New-Item -Path $SecretStorePath -ItemType Directory -Force | Out-Null }

$cred = Resolve-Credential -Username $Username -Provided $Credential -StorePath $SecretStorePath -Save:$SavePasswords -Prompt:$PromptForPasswords

$drvName = 'OTPTGT'
if (Get-PSDrive -Name $drvName -ErrorAction SilentlyContinue) { Remove-PSDrive -Name $drvName -Force -ErrorAction SilentlyContinue }
try {
    New-PSDrive -Name $drvName -PSProvider FileSystem -Root "\\$Target\$RemoteDrive" -Credential $cred -ErrorAction Stop | Out-Null
    Write-Log OK "SMB drive mapped: $drvName -> \\$Target\$RemoteDrive"
} catch {
    throw "Failed to map \\$Target\$RemoteDrive : $($_.Exception.Message)"
}

try {
    $remoteCertDirLocal = Split-Path $KeystorePath -Parent

    switch ($Phase) {
        'GenerateCsr' {
            $ksPw      = Resolve-KeystorePassword -StorePath $SecretStorePath -Save:$SavePasswords -Prompt:$PromptForPasswords -Purpose 'KeystorePassword'
            $ksPwPlain = ConvertFrom-SecureStringPlain $ksPw
            $pwB64     = ConvertTo-B64 $ksPwPlain

            $runner = Build-RunnerGenerateCsr -KeystorePath $KeystorePath -NewKeystorePath $NewKeystorePath `
                -CsrPath $CsrPath -Alias $Alias -Subject $Subject -SanDns $SanDns -KeySize $KeySize -PwB64 $pwB64

            $r = Invoke-PhaseSmb -Target $Target -RemoteDrive $RemoteDrive -RemoteCertDir $remoteCertDirLocal `
                -Credential $cred -RunnerCode $runner -TimeoutSeconds $TimeoutSeconds

            $s = $r.Status
            if ($s.status -ne 'Success') {
                Write-Log ERROR "GenerateCsr failed on target: $($s.error)"
                throw $s.error
            }
            Write-Log OK "CSR generated on target: $($s.csrPath)"

            $localCsr = Join-Path $LocalOutputDir 'OTP.csr'
            $remoteCsrUnc = Join-Path $r.SharePath 'OTP.csr'
            Copy-Item -Path $remoteCsrUnc -Destination $localCsr -Force
            Write-Log OK "CSR copied back to jump server: $localCsr"

            Write-Host ''
            Write-Host '======= CSR — paste into email to CA / Rhonda =======' -ForegroundColor Yellow
            Write-Host $s.csrContent
            Write-Host '=====================================================' -ForegroundColor Yellow
            Write-Host ''
            Write-Log OK "Phase 1 done. Email CSR above. Run Phase 2 once signed cert returns."
        }

        'ImportSignedCert' {
            $ksPwFile = Join-Path $SecretStorePath 'KeystorePassword.xml'
            if (-not (Test-Path $ksPwFile)) {
                Write-Log WARN "Keystore password not saved — prompting"
                $ksPw = Resolve-KeystorePassword -StorePath $SecretStorePath -Prompt -Purpose 'KeystorePassword'
            } else {
                $ksPw = Import-Clixml -Path $ksPwFile
            }
            $ksPwPlain = ConvertFrom-SecureStringPlain $ksPw
            $pwB64     = ConvertTo-B64 $ksPwPlain

            $shareDir = Get-RemoteCertShare -Target $Target -Drive $RemoteDrive -RemoteDirLocal $remoteCertDirLocal
            $remoteCerUnc = Join-Path $shareDir 'OTP.cer'

            if ($SignedCertFile) {
                if (-not (Test-Path $SignedCertFile)) { throw "SignedCertFile not found: $SignedCertFile" }
                Copy-Item -Path $SignedCertFile -Destination $remoteCerUnc -Force
                Write-Log OK "Signed cert pushed to $remoteCerUnc"
            }
            elseif (-not (Test-Path $remoteCerUnc)) {
                throw "No -SignedCertFile and no cert at $remoteCerUnc. Pass -SignedCertFile <path> or have CA upload to E:\CERTS\OTP.cer"
            }
            else {
                Write-Log OK "Using signed cert already on target: $remoteCerUnc"
            }

            $runner = Build-RunnerImportSignedCert -KeystorePath $KeystorePath -NewKeystorePath $NewKeystorePath `
                -CerPath $CerPath -Alias $Alias -ServiceName $ServiceName -PwB64 $pwB64

            $r = Invoke-PhaseSmb -Target $Target -RemoteDrive $RemoteDrive -RemoteCertDir $remoteCertDirLocal `
                -Credential $cred -RunnerCode $runner -TimeoutSeconds $TimeoutSeconds

            $s = $r.Status
            if ($s.status -ne 'Success') {
                Write-Log ERROR "ImportSignedCert failed: $($s.error)"
                if ($s.rolledBack) { Write-Log WARN "Auto-rollback executed. Restored backup: $($s.backupPath)" }
                throw $s.error
            }
            Write-Log OK "Cert imported, keystore swapped, service: $($s.serviceStatus). Backup: $($s.backupPath)"
        }

        'Verify' {
            $ksPwFile = Join-Path $SecretStorePath 'KeystorePassword.xml'
            if (-not (Test-Path $ksPwFile)) {
                $ksPw = Resolve-KeystorePassword -StorePath $SecretStorePath -Prompt -Purpose 'KeystorePassword'
            } else {
                $ksPw = Import-Clixml -Path $ksPwFile
            }
            $ksPwPlain = ConvertFrom-SecureStringPlain $ksPw
            $pwB64     = ConvertTo-B64 $ksPwPlain

            $runner = Build-RunnerVerify -KeystorePath $KeystorePath -Alias $Alias -ServiceName $ServiceName -PwB64 $pwB64
            $r = Invoke-PhaseSmb -Target $Target -RemoteDrive $RemoteDrive -RemoteCertDir $remoteCertDirLocal `
                -Credential $cred -RunnerCode $runner -TimeoutSeconds $TimeoutSeconds
            $s = $r.Status
            if ($s.status -ne 'Success') { throw $s.error }
            Write-Log OK "Keystore: $($s.keystoreExists)  Service: $($s.serviceStatus)"
            Write-Host ''
            Write-Host '======= Cert Info =======' -ForegroundColor Yellow
            Write-Host $s.certInfo
            Write-Host '=========================' -ForegroundColor Yellow
        }

        'Rollback' {
            $runner = Build-RunnerRollback -KeystorePath $KeystorePath -ServiceName $ServiceName
            $r = Invoke-PhaseSmb -Target $Target -RemoteDrive $RemoteDrive -RemoteCertDir $remoteCertDirLocal `
                -Credential $cred -RunnerCode $runner -TimeoutSeconds $TimeoutSeconds
            $s = $r.Status
            if ($s.status -ne 'Success') { throw $s.error }
            Write-Log OK "Rolled back: $($s.restoredFrom)  Service: $($s.serviceStatus)"
        }
    }
}
finally {
    if (Get-PSDrive -Name $drvName -ErrorAction SilentlyContinue) { Remove-PSDrive -Name $drvName -Force -ErrorAction SilentlyContinue }
}

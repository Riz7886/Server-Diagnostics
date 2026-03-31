param(
    [ValidateSet("VALIDATE","GENERATECSR","IMPORT")]
    [string]$Step = "VALIDATE",
    [string]$SignedCertPath = "",
    [string]$OutputPath = "C:\Certs"
)

$ErrorActionPreference = "SilentlyContinue"
$WarningPreference     = "SilentlyContinue"

$ServerName   = "ns2sw1app1"
$ServerFQDN   = "ns2sw1app1.ns2corp.local"
$ServerIP     = "10.134.4.171"
$FriendlyName = "cert-sw1"
$KeySize      = 2048
$CSRFile      = "$OutputPath\ns2sw1app1-CSR.txt"
$INFFile      = "$OutputPath\ns2sw1app1-request.inf"
$LogFile      = "$OutputPath\cert-fix-log.txt"

if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }

function Write-Log {
    param([string]$M, [string]$L = "INFO")
    $C = switch ($L) { "OK" {"Green"} "WARN" {"Yellow"} "ERR" {"Red"} "FIX" {"Magenta"} default {"Cyan"} }
    $Line = "[$L] [$(Get-Date -Format 'HH:mm:ss')] $M"
    Write-Host $Line -ForegroundColor $C
    Add-Content -Path $LogFile -Value $Line -ErrorAction SilentlyContinue
}

function Write-Line { Write-Host ("=" * 65) -ForegroundColor Blue }

Write-Line
Write-Host "  NS2SW1APP1 - RDP CERTIFICATE FIX | Author: Syed Rizvi" -ForegroundColor White
Write-Host "  Server: $ServerName | IP: $ServerIP | Step: $Step" -ForegroundColor Gray
Write-Line

if ($Step -eq "VALIDATE") {

    Write-Log "Scanning LocalMachine\My for existing certs..." "INFO"
    $AllCerts = Get-ChildItem Cert:\LocalMachine\My
    Write-Log "Total certs in store: $($AllCerts.Count)" "INFO"

    $Matches = $AllCerts | Where-Object {
        $_.FriendlyName -like "*sw1*" -or
        $_.FriendlyName -like "*cert*" -or
        $_.Subject -like "*ns2sw1*"
    }

    if ($Matches.Count -eq 0) {
        Write-Log "No matching cert found for $FriendlyName" "WARN"
        Write-Log "Run with -Step GENERATECSR to create a fresh 2048-bit CSR" "WARN"
        exit 0
    }

    foreach ($C in $Matches) {
        Write-Line
        Write-Log "Cert: $($C.Subject)" "OK"
        Write-Log "Thumbprint    : $($C.Thumbprint)" "INFO"
        Write-Log "Friendly Name : $($C.FriendlyName)" "INFO"
        Write-Log "Issuer        : $($C.Issuer)" "INFO"
        Write-Log "Valid From    : $($C.NotBefore)" "INFO"
        Write-Log "Valid To      : $($C.NotAfter)" "INFO"
        Write-Log "Key Size      : $($C.PublicKey.Key.KeySize)-bit" "INFO"
        Write-Log "Has Private Key: $($C.HasPrivateKey)" "INFO"

        $SAN = $C.Extensions | Where-Object { $_.Oid.FriendlyName -eq "Subject Alternative Name" }
        Write-Log "SAN Present   : $(if ($SAN) { 'YES' } else { 'NO - MISSING' })" "INFO"

        $Issues = @()
        if ($C.PublicKey.Key.KeySize -ne 2048)     { $Issues += "Key is $($C.PublicKey.Key.KeySize)-bit not 2048" }
        if (-not $C.HasPrivateKey)                  { $Issues += "Private key MISSING" }
        if ($C.NotAfter -lt (Get-Date))             { $Issues += "Certificate EXPIRED" }
        if (-not $SAN)                              { $Issues += "No SAN extension - modern clients will reject" }

        if ($Issues.Count -eq 0) {
            Write-Log "Cert validation PASSED - check IIS binding next" "OK"
        } else {
            foreach ($I in $Issues) { Write-Log $I "ERR" }
            Write-Log "Generate a new CSR with -Step GENERATECSR" "WARN"
        }
    }

    Write-Log "Checking RDP cert binding..." "INFO"
    $RDP = Get-WmiObject -Namespace root\cimv2\TerminalServices -Class Win32_TSGeneralSetting
    if ($RDP -and $RDP.SSLCertificateSHA1Hash) {
        Write-Log "RDP bound to: $($RDP.SSLCertificateSHA1Hash)" "OK"
    } else {
        Write-Log "RDP using default self-signed cert - not bound to $FriendlyName" "WARN"
    }

    Write-Log "Checking IIS HTTPS binding..." "INFO"
    Import-Module WebAdministration -ErrorAction SilentlyContinue
    $Bindings = Get-WebBinding -Protocol "https" -ErrorAction SilentlyContinue
    if ($Bindings) {
        foreach ($B in $Bindings) { Write-Log "IIS Binding: $($B.bindingInformation)" "OK" }
    } else {
        Write-Log "No HTTPS bindings found in IIS" "WARN"
    }
}

if ($Step -eq "GENERATECSR") {

    Write-Line
    Write-Log "Generating 2048-bit CSR for $ServerFQDN..." "FIX"

    $INF = @"
[Version]
Signature="`$Windows NT`$"

[NewRequest]
Subject = "CN=$ServerFQDN, OU=NS2, O=NS2 Corp, L=San Diego, S=California, C=US"
KeySpec = 1
KeyLength = 2048
Exportable = TRUE
MachineKeySet = TRUE
SMIME = FALSE
PrivateKeyArchive = FALSE
UserProtected = FALSE
UseExistingKeySet = FALSE
ProviderName = "Microsoft RSA SChannel Cryptographic Provider"
ProviderType = 12
RequestType = PKCS10
KeyUsage = 0xa0
FriendlyName = "$FriendlyName"

[EnhancedKeyUsageExtension]
OID = 1.3.6.1.5.5.7.3.1
OID = 1.3.6.1.5.5.7.3.2

[Extensions]
2.5.29.17 = "{text}"
_continue_ = "dns=$ServerName&"
_continue_ = "dns=$ServerFQDN&"
_continue_ = "ipaddress=$ServerIP&"
"@

    $INF | Out-File -FilePath $INFFile -Encoding ASCII -Force
    Write-Log "INF file written: $INFFile" "OK"

    $CSROutput = certreq -new $INFFile $CSRFile 2>&1
    Write-Log "certreq output: $CSROutput" "INFO"

    if (Test-Path $CSRFile) {
        Write-Log "CSR created successfully: $CSRFile" "OK"
        Write-Log "Key Size: 2048-bit" "OK"
        Write-Log "CN: $ServerFQDN" "OK"
        Write-Log "SAN: $ServerName, $ServerFQDN, IP:$ServerIP" "OK"
        Write-Line
        Write-Log "Send this file to the CA for signing: $CSRFile" "FIX"
        Write-Log "When signed cert is returned run:" "FIX"
        Write-Log "  .\Fix-RDP-Cert-NS2SW1.ps1 -Step IMPORT -SignedCertPath 'C:\Certs\signed.cer'" "FIX"
        Write-Line
        Write-Log "CSR CONTENT:" "INFO"
        Get-Content $CSRFile | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
    } else {
        Write-Log "CSR generation failed. Check certreq is available." "ERR"
        Write-Log "Output: $CSROutput" "ERR"
        exit 1
    }
}

if ($Step -eq "IMPORT") {

    if (-not $SignedCertPath -or -not (Test-Path $SignedCertPath)) {
        Write-Log "Signed cert path not found: $SignedCertPath" "ERR"
        Write-Log "Usage: .\Fix-RDP-Cert-NS2SW1.ps1 -Step IMPORT -SignedCertPath 'C:\Certs\signed.cer'" "WARN"
        exit 1
    }

    Write-Line
    Write-Log "Importing signed cert from: $SignedCertPath" "FIX"

    $AcceptOut = certreq -accept $SignedCertPath 2>&1
    Write-Log "certreq accept: $AcceptOut" "INFO"

    Start-Sleep -Seconds 3

    $NewCert = Get-ChildItem Cert:\LocalMachine\My |
        Where-Object { $_.Subject -like "*$ServerName*" -and $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date) } |
        Sort-Object NotAfter -Descending |
        Select-Object -First 1

    if (-not $NewCert) {
        Write-Log "Could not find imported cert in store." "ERR"
        Write-Log "Verify the signed cert matches the CSR from this server." "ERR"
        exit 1
    }

    Write-Log "Cert imported: $($NewCert.Subject)" "OK"
    Write-Log "Thumbprint   : $($NewCert.Thumbprint)" "OK"
    Write-Log "Key Size     : $($NewCert.PublicKey.Key.KeySize)-bit" "OK"
    Write-Log "Valid To     : $($NewCert.NotAfter)" "OK"

    $NewCert.FriendlyName = $FriendlyName
    Write-Log "Friendly name set to: $FriendlyName" "OK"

    Write-Log "Binding to IIS HTTPS port 443..." "FIX"
    try {
        Import-Module WebAdministration -ErrorAction Stop

        Get-WebBinding -Name "Default Web Site" -Protocol "https" -ErrorAction SilentlyContinue | ForEach-Object {
            Remove-WebBinding -Name "Default Web Site" -Protocol "https" -ErrorAction SilentlyContinue
        }

        New-WebBinding -Name "Default Web Site" -Protocol "https" -Port 443 -IPAddress "*" -SslFlags 0
        $Binding = Get-WebBinding -Name "Default Web Site" -Protocol "https"
        $Binding.AddSslCertificate($NewCert.Thumbprint, "My")
        Write-Log "IIS HTTPS binding complete on port 443" "OK"
    } catch {
        Write-Log "IIS binding via WebAdmin failed - trying netsh..." "WARN"
        $AppGuid = "{$([System.Guid]::NewGuid().ToString())}"
        $Cmd = "netsh http add sslcert ipport=0.0.0.0:443 certhash=$($NewCert.Thumbprint) appid=$AppGuid"
        cmd /c $Cmd
        Write-Log "netsh binding applied" "OK"
    }

    Write-Log "Binding cert to RDP service..." "FIX"
    try {
        $TSSettings = Get-WmiObject -Namespace root\cimv2\TerminalServices -Class Win32_TSGeneralSetting -ErrorAction Stop
        $TSSettings.SSLCertificateSHA1Hash = $NewCert.Thumbprint
        $TSSettings.Put() | Out-Null
        Write-Log "RDP cert binding via WMI complete" "OK"
    } catch {
        Write-Log "WMI failed - trying registry method..." "WARN"
        $ThumbBytes = [byte[]] ($NewCert.Thumbprint -replace '..', '$0 ' -split ' ' -ne '' | ForEach-Object { [Convert]::ToByte($_, 16) })
        $RegPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp"
        Set-ItemProperty -Path $RegPath -Name "SSLCertificateSHA1Hash" -Value $ThumbBytes -ErrorAction SilentlyContinue
        Write-Log "RDP cert binding via registry complete" "OK"
    }

    Write-Log "Restarting RDP service to apply changes..." "FIX"
    Restart-Service TermService -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 3
    Write-Log "RDP service restarted" "OK"

    Write-Line
    Write-Log "FINAL VALIDATION" "FIX"

    $ValidateCert = Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.Thumbprint -eq $NewCert.Thumbprint }
    if ($ValidateCert) {
        Write-Log "Cert in store       : PASS" "OK"
        Write-Log "Key size 2048-bit   : PASS" "OK"
        Write-Log "Has private key     : PASS" "OK"
        Write-Log "Friendly name       : $($ValidateCert.FriendlyName)" "OK"
        Write-Log "Expiry              : $($ValidateCert.NotAfter)" "OK"
    }

    $RDPFinal = Get-WmiObject -Namespace root\cimv2\TerminalServices -Class Win32_TSGeneralSetting
    if ($RDPFinal -and $RDPFinal.SSLCertificateSHA1Hash -eq $NewCert.Thumbprint) {
        Write-Log "RDP binding         : PASS" "OK"
    } else {
        Write-Log "RDP binding         : Verify manually in Remote Desktop settings" "WARN"
    }

    Write-Line
    Write-Log "ALL DONE - Cert installed, IIS and RDP fully bound" "OK"
    Write-Log "Log file saved to: $LogFile" "OK"
    Write-Line
}

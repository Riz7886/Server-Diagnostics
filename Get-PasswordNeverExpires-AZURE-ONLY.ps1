$ErrorActionPreference = "Stop"

Clear-Host
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  EXPORT USERS WITH PASSWORD EXPIRATION DISABLED" -ForegroundColor Cyan
Write-Host "  Azure AD / Entra ID Only" -ForegroundColor Cyan
Write-Host "  For: Tony Schlak" -ForegroundColor Cyan
Write-Host "  Created by: Syed Rizvi" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$desktop = [Environment]::GetFolderPath("Desktop")
$csvFile = Join-Path $desktop "PasswordNeverExpires_Users_$timestamp.csv"

Write-Host "Checking for Microsoft Graph PowerShell module..." -ForegroundColor Yellow

try {
    Import-Module Microsoft.Graph.Users -ErrorAction Stop
    Write-Host "  Microsoft Graph module found" -ForegroundColor Green
} catch {
    Write-Host "  Microsoft Graph module not found" -ForegroundColor Red
    Write-Host ""
    Write-Host "Installing Microsoft Graph module..." -ForegroundColor Yellow
    Write-Host "This may take a few minutes..." -ForegroundColor Gray
    
    try {
        Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
        Import-Module Microsoft.Graph.Users -ErrorAction Stop
        Write-Host "  Microsoft Graph module installed successfully" -ForegroundColor Green
    } catch {
        Write-Host ""
        Write-Host "ERROR: Cannot install Microsoft Graph module" -ForegroundColor Red
        Write-Host "Please run this command manually first:" -ForegroundColor Yellow
        Write-Host "  Install-Module Microsoft.Graph -Scope CurrentUser" -ForegroundColor Gray
        exit 1
    }
}

Write-Host ""
Write-Host "Connecting to Microsoft Entra ID (Azure AD)..." -ForegroundColor Cyan

try {
    Connect-MgGraph -Scopes "User.Read.All" -NoWelcome
    Write-Host "  Connected successfully" -ForegroundColor Green
} catch {
    Write-Host "  Connection failed" -ForegroundColor Red
    Write-Host ""
    Write-Host "ERROR: Cannot connect to Azure AD" -ForegroundColor Red
    Write-Host "Make sure you have permissions to read users" -ForegroundColor Yellow
    exit 1
}

Write-Host ""
Write-Host "Retrieving all users from Azure AD..." -ForegroundColor Cyan
Write-Host "This may take a few minutes depending on user count..." -ForegroundColor Gray
Write-Host ""

try {
    $allUsers = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,Mail,AccountEnabled,PasswordPolicies,Department,JobTitle,CreatedDateTime -ErrorAction Stop
    Write-Host "  Retrieved $($allUsers.Count) total users" -ForegroundColor Green
} catch {
    Write-Host "  Failed to retrieve users" -ForegroundColor Red
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-MgGraph
    exit 1
}

Write-Host ""
Write-Host "Filtering users with password expiration disabled..." -ForegroundColor Cyan

$usersWithNoExpiry = $allUsers | Where-Object {
    $_.PasswordPolicies -ne $null -and $_.PasswordPolicies -like "*DisablePasswordExpiration*"
}

if ($usersWithNoExpiry.Count -eq 0) {
    Write-Host ""
    Write-Host "No users found with password expiration disabled!" -ForegroundColor Yellow
    Write-Host ""
    Disconnect-MgGraph
    exit 0
}

Write-Host "  Found $($usersWithNoExpiry.Count) users with password expiration disabled" -ForegroundColor Green
Write-Host ""

Write-Host "Creating CSV export..." -ForegroundColor Cyan

$results = @()

foreach ($user in $usersWithNoExpiry) {
    $results += [PSCustomObject]@{
        'Display Name' = $user.DisplayName
        'User Principal Name' = $user.UserPrincipalName
        'Email' = $user.Mail
        'Account Enabled' = $user.AccountEnabled
        'Password Policy' = $user.PasswordPolicies
        'Department' = $user.Department
        'Job Title' = $user.JobTitle
        'Account Created' = $user.CreatedDateTime
        'User ID' = $user.Id
    }
}

try {
    $results | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8
    Write-Host "  CSV file created successfully" -ForegroundColor Green
} catch {
    Write-Host "  Failed to create CSV" -ForegroundColor Red
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-MgGraph
    exit 1
}

Disconnect-MgGraph | Out-Null

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  SUCCESS" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Total users with password expiration disabled: $($usersWithNoExpiry.Count)" -ForegroundColor White
Write-Host ""

$enabledCount = ($results | Where-Object { $_.'Account Enabled' -eq $true }).Count
$disabledCount = ($results | Where-Object { $_.'Account Enabled' -eq $false }).Count

Write-Host "Account Status:" -ForegroundColor Yellow
Write-Host "  Enabled:  $enabledCount" -ForegroundColor Green
Write-Host "  Disabled: $disabledCount" -ForegroundColor Gray
Write-Host ""
Write-Host "CSV file location:" -ForegroundColor Yellow
Write-Host "  $csvFile" -ForegroundColor White
Write-Host ""

Start-Process $csvFile

Write-Host "CSV file opened automatically" -ForegroundColor Green
Write-Host ""
Write-Host "Send this file to Tony Schlak" -ForegroundColor Cyan
Write-Host ""
Write-Host "DONE" -ForegroundColor Green
Write-Host ""

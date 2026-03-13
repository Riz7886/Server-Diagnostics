$SavePath = "C:\temp\kb5078752-x64.msu"
$RemoteHost = "10.70.20.181"

if (-not (Test-Path "C:\temp")) {
    New-Item -ItemType Directory -Path "C:\temp" -Force
}

Write-Host "Searching Microsoft Update Catalog..." -ForegroundColor Cyan
$SearchPage = Invoke-WebRequest -Uri "https://www.catalog.update.microsoft.com/Search.aspx?q=KB5078752" -UseBasicParsing
$UpdateIDs = [regex]::Matches($SearchPage.Content, "goToDetails\('([a-f0-9\-]+)'\)") | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique

$DownloadURL = $null
foreach ($ID in $UpdateIDs) {
    $DLPage = Invoke-WebRequest -Uri "https://www.catalog.update.microsoft.com/DownloadDialog.aspx" -Method Post -Body "updateIDs=[{`"size`":0,`"languages`":`"`",`"uidInfo`":`"$ID`",`"updateID`":`"$ID`"}]&updateIDsBlockedForImport=&wsusApiPresent=&contentImport=&sku=&serverName=&ssl=&portNumber=&version=" -UseBasicParsing
    $Links = [regex]::Matches($DLPage.Content, 'https://[^"]+\.msu') | ForEach-Object { $_.Value }
    $DownloadURL = $Links | Where-Object { $_ -like "*x64*" } | Select-Object -First 1
    if ($DownloadURL) { break }
}

if (-not $DownloadURL) {
    Write-Host "ERROR: Could not find download URL" -ForegroundColor Red
    exit
}

Write-Host "Downloading KB5078752..." -ForegroundColor Cyan
Start-BitsTransfer -Source $DownloadURL -Destination $SavePath
Write-Host "Download complete." -ForegroundColor Green

Write-Host "Copying to $RemoteHost..." -ForegroundColor Cyan
$RemotePath = "\\$RemoteHost\c$\temp"

if (-not (Test-Path $RemotePath)) {
    New-Item -ItemType Directory -Path $RemotePath -Force
}

Copy-Item -Path $SavePath -Destination $RemotePath -Force
Write-Host "DONE - File is at C:\temp on $RemoteHost" -ForegroundColor Green

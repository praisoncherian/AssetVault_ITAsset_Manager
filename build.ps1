$ErrorActionPreference = "Stop"

$MainDir = "D:\Personal Praison\Develop\Asset manager"
$BuildDir = "$MainDir\AssetVault_v0.1"
$ZipFile = "$MainDir\AssetVault_v0.1.zip"

Write-Host "Creating build directory $BuildDir..."
if (-not (Test-Path $BuildDir)) {
    New-Item -ItemType Directory -Force -Path $BuildDir | Out-Null
}

Write-Host "Copying source files..."
Copy-Item "$MainDir\AssetManagerPro.ps1" -Destination "$BuildDir\AssetManagerPro.ps1" -Force
Copy-Item "$MainDir\AssetVault.ico" -Destination "$BuildDir\AssetVault.ico" -Force
Copy-Item "$MainDir\AssetVault_app_icon.png" -Destination "$BuildDir\AssetVault_app_icon.png" -Force

Write-Host "Killing any running instances of AssetVault..."
Stop-Process -Name 'AssetVault' -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

Write-Host "Compiling executable..."
Import-Module ps2exe -ErrorAction SilentlyContinue
Invoke-ps2exe "$BuildDir\AssetManagerPro.ps1" "$BuildDir\AssetVault.exe" -icon "$BuildDir\AssetVault.ico" -noConsole -RequireAdmin -company "Arbor Solutions" -product "AssetVault IT Asset Manager Platform" -title "AssetVault" -description "AssetVault IT Asset Manager" -version "1.0.0.0" -copyright "Arbor Solutions"

Write-Host "Signing executable..."
$certs = @(Get-ChildItem -Path Cert:\CurrentUser\My -CodeSigningCert | Where-Object Subject -match 'Arbor Solutions')
if ($certs.Count -eq 0) {
    Write-Host "Generating new Code Signing Certificate..."
    $certs = @(New-SelfSignedCertificate -Subject 'CN=Arbor Solutions' -Type CodeSigningCert -CertStoreLocation Cert:\CurrentUser\My)
}

if ($certs.Count -gt 0) {
    Set-AuthenticodeSignature -FilePath "$BuildDir\AssetVault.exe" -Certificate $certs[0]
}

Write-Host "Creating zip archive..."
if (Test-Path $ZipFile) { Remove-Item $ZipFile -Force }
Compress-Archive -Path "$BuildDir\*" -DestinationPath $ZipFile -Force

Write-Host "Build complete! Output: $ZipFile"

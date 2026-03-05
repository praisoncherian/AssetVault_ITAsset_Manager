# Installation Guide

AssetVault IT Asset Manager is distributed as a pre-compiled Windows executable and as a raw PowerShell script for advanced developers.

## Option 1: Executable (Recommended)
1. Download the `AssetVault_v0.1.zip` (v1.0.0) release package.
2. Extract the archive to a local directory (e.g., `C:\AssetVault`).
3. Run `AssetVault.exe`. 
   > Note: A UAC prompt will appear requesting Administrator privileges. This is required to execute the WMI queries.

## Option 2: Running from Source
1. Clone this repository to your Windows machine.
2. Open PowerShell as Administrator.
3. Navigate to `src/`.
4. Run `.\AssetManagerPro.ps1`.

## Compiling Your Own Executable
If you modify the source code, you can use the `ps2exe` module to bundle it into an updated `.exe`.
```powershell
Install-Module ps2exe -Scope CurrentUser
Invoke-ps2exe ".\src\AssetManagerPro.ps1" ".\AssetVault.exe" -icon ".\AssetVault.ico" -noConsole -RequireAdmin
```

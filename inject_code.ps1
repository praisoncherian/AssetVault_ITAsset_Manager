# Diagnostic Framework Queries & HTML Template

$ReplaceBlock = @'
        # ==========================================
        # 1. ASSET INFORMATION
        # ==========================================
        $UUID = "Unknown"; Try { $UUID = (Get-CimInstance Win32_ComputerSystemProduct).UUID } Catch {}
        $Chassis = "Unknown"; Try { $Chassis = (Get-CimInstance Win32_SystemEnclosure).ChassisTypes[0] } Catch {}
        $Domain = $System.Domain
        
        # ==========================================
        # 2. HARDWARE INVENTORY
        # ==========================================
        $GPUName = "Unknown"; $GPUVRAM = "Unknown"; $GPUDriver = "Unknown"
        Try {
            $GPU = Get-CimInstance Win32_VideoController
            $GPUName = ($GPU.Name) -join ", "
            $GPUVRAM = ($GPU | ForEach-Object { "{0:N2} GB" -f ($_.AdapterRAM / 1GB) }) -join ", "
            $GPUDriver = ($GPU.DriverVersion) -join ", "
        } Catch {}
        
        $MobMfg = "Unknown"; $MobModel = "Unknown"
        Try { $Mob = Get-CimInstance Win32_BaseBoard; $MobMfg = $Mob.Manufacturer; $MobModel = $Mob.Product } Catch {}
        
        $RamSpeed = "Unknown"; Try { $RamSpeed = (Get-CimInstance Win32_PhysicalMemory | Select -First 1).Speed } Catch {}
        
        $StorageRows = ""
        Get-CimInstance Win32_DiskDrive -ErrorAction SilentlyContinue | ForEach-Object {
            $sz = "{0:N2}" -f ($_.Size / 1GB)
            $StorageRows += "<tr><td>$($_.Model)</td><td>$($_.SerialNumber)</td><td>$sz GB</td><td>$($_.InterfaceType)</td><td>$($_.MediaType)</td></tr>"
        }
        
        # ==========================================
        # 3. HARDWARE HEALTH DIAGNOSTICS
        # ==========================================
        $DiskHealthRows = ""
        Try {
            Get-PhysicalDisk -ErrorAction SilentlyContinue | ForEach-Object {
                $DiskHealthRows += "<tr><td>$($_.FriendlyName)</td><td>$($_.HealthStatus)</td><td>$($_.OperationalStatus)</td></tr>"
            }
        } Catch { $DiskHealthRows = "<tr><td colspan='3'>Disk health queries not supported or run without admin rights.</td></tr>" }
        
        # ==========================================
        # 4. SOFTWARE INVENTORY
        # ==========================================
        $SoftRows = ""
        Try {
            $Keys = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
            $SoftList = Get-ItemProperty $Keys -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -and $_.SystemComponent -ne 1 } | Sort-Object InstallDate -Descending | Select-Object DisplayName, DisplayVersion, Publisher -First 10
            foreach ($s in $SoftList) {
                $SoftRows += "<tr><td>$($s.DisplayName)</td><td>$($s.DisplayVersion)</td><td>$($s.Publisher)</td></tr>"
            }
        } Catch {}
        
        # ==========================================
        # 5. OS STATUS & 6. SECURITY AUDIT
        # ==========================================
        $Firmware = $env:firmware_type
        $TimeZone = (Get-TimeZone).Id
        
        $BitLocker = "Not Checked"
        Try { $BitLocker = (Get-BitLockerVolume -MountPoint 'C:' -ErrorAction SilentlyContinue).VolumeStatus } Catch {}
        
        $Admins = "Unknown"
        Try { $Admins = (Get-LocalGroupMember -Group "Administrators" | Where-Object PrincipalSource -eq "Local" | Select-Object -ExpandProperty Name) -join ", " } Catch {}
        
        # ==========================================
        # 7. NETWORK DIAGNOSTICS
        # ==========================================
        $NetRows = ""
        Try {
            Get-NetAdapter | Where-Object Status -eq 'Up' | ForEach-Object {
                $IpObj = Get-NetIPConfiguration -InterfaceAlias $_.Name -ErrorAction SilentlyContinue
                $IpStr = if ($IpObj.IPv4Address) { ($IpObj.IPv4Address.IPAddress -join ", ") } else { "None" }
                $NetRows += "<tr><td>$($_.Name)</td><td>$($_.MacAddress)</td><td>$IpStr</td><td>$($_.LinkSpeed)</td></tr>"
            }
        } Catch {}

        # ==========================================
        # 8. PERFORMANCE & 9. EVENT LOGS
        # ==========================================
        $BootTime = $OS.LastBootUpTime
        
        $EventRows = ""
        Try {
            Get-WinEvent -FilterHashtable @{LogName='System','Application'; Level=2} -MaxEvents 5 -ErrorAction SilentlyContinue | ForEach-Object {
                $EventRows += "<tr><td>$($_.TimeCreated)</td><td>$($_.ProviderName)</td><td>$($_.Id)</td><td style='font-size: 11px;'>$($_.Message -replace "`n", " ")</td></tr>"
            }
        } Catch { $EventRows = "<tr><td colspan='4'>No recent critical events or access denied.</td></tr>" }

        # ==========================================
        # 11. PERIPHERALS
        # ==========================================
        $Monitors = "Unknown"; Try { $Monitors = (Get-CimInstance WmiMonitorID -Namespace root\wmi | ForEach-Object { [System.Text.Encoding]::ASCII.GetString($_.UserFriendlyName) -replace "`0", "" }) -join ", " } Catch {}
        $Printers = "Unknown"; Try { $Printers = (Get-CimInstance Win32_Printer | Where-Object Default -eq $true | Select-Object -ExpandProperty Name) -join ", " } Catch {}

        # ==========================================
        # 10. BATTERY & POWER (Powercfg execution)
        # ==========================================
        $BattHtml = "$CachePath\batt_temp_$Timestamp.html"
        $EnergyHtml = "$CachePath\energy_temp_$Timestamp.html"
        
        Start-Process powercfg -ArgumentList "/batteryreport /output `"$BattHtml`"" -NoNewWindow -Wait
        Start-Process powercfg -ArgumentList "/energy /output `"$EnergyHtml`"" -NoNewWindow -Wait
        
        $BattContent = "<p>Battery report unavailable.</p>"
        $EnergyContent = "<p>Energy report unavailable.</p>"
        if (Test-Path $BattHtml) { $raw = Get-Content $BattHtml -Raw; if ($raw -match '(?si)<body[^>]*>(.*?)</body>') { $BattContent = $matches[1] }; Remove-Item $BattHtml -Force }
        if (Test-Path $EnergyHtml) { $raw = Get-Content $EnergyHtml -Raw; if ($raw -match '(?si)<body[^>]*>(.*?)</body>') { $EnergyContent = $matches[1] }; Remove-Item $EnergyHtml -Force }

        # Value Fallbacks
        $ValStaff = if ($txtTechnician.Text.Trim() -ne "") { $txtTechnician.Text } else { "_________________" }
        $ValUser = if ($txtUser.Text.Trim() -ne "") { $txtUser.Text } else { "_________________" }

        $HtmlBody = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title>AssetVault IT Hardware Report</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Verdana, sans-serif; margin: 40px; color: #333; line-height: 1.5; font-size: 13px; }
        h1 { text-align: center; color: #1E3A8A; margin-bottom: 5px; font-size: 28px; }
        .subtitle { text-align: center; color: #64748B; border-bottom: 2px solid #1E3A8A; padding-bottom: 10px; margin-bottom: 30px; font-size: 18px; }
        h2 { color: #1E3A8A; margin-top: 25px; border-bottom: 1px solid #CBD5E1; padding-bottom: 3px; font-size: 16px; text-transform: uppercase; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; margin-bottom: 15px; font-size: 12px; }
        th, td { border: 1px solid #E2E8F0; padding: 8px; text-align: left; vertical-align: top; }
        th { background-color: #F8FAFC; width: 30%; font-weight: 600; color: #334155; }
        td { background-color: #ffffff; color: #0F172A; }
        
        /* New Terms & Agreement Layout */
        .agreement-box { margin-top: 40px; padding: 20px; border: 2px solid #333; background: #fff; page-break-inside: avoid; }
        .agreement-title { font-weight: bold; font-size: 16px; margin-bottom: 10px; text-transform: uppercase; text-align: center; border-bottom: 1px solid #ccc; padding-bottom: 10px;}
        .agreement-text { font-size: 12px; color: #444; margin-bottom: 20px; }
        
        .sign-grid { display: table; width: 100%; margin-top: 30px; }
        .sign-col { display: table-cell; width: 33.3%; text-align: center; vertical-align: bottom; }
        .sign-line { border-bottom: 1px solid #000; margin: 0 auto 5px auto; width: 80%; height: 60px; }
        .sign-label { font-size: 12px; font-weight: bold; }
        .sign-name { font-size: 13px; margin-top: 5px; color: #1E3A8A; }

        .powercfg { zoom: 0.65; margin-top: 15px; border: 1px solid #eee; padding: 15px; background: #fdfdfd; width: 100%; overflow-x: hidden; }
        .powercfg body { margin: 0; padding: 0; }
        .page-break { page-break-before: always; }
    </style>
</head>
<body>

    <h1>AssetVault</h1>
    <div class="subtitle">Diagnostic &amp; Hardware Report</div>
    
    <h2>1. Asset Information</h2>
    <table>
        <tr><th>Report Generated</th><td>$DisplayDate (Local Time)</td></tr>
        <tr><th>Computer Name / Domain</th><td>$ComputerName / $Domain</td></tr>
        <tr><th>Asset Tag</th><td>$($txtAssetTag.Text)</td></tr>
        <tr><th>Manufacturer / Model</th><td>$($System.Manufacturer) / $($System.Model)</td></tr>
        <tr><th>Serial Number / UUID</th><td>$($BIOS.SerialNumber) / $UUID</td></tr>
        <tr><th>Chassis Type</th><td>$Chassis</td></tr>
        <tr><th>Assigned User</th><td>$($txtUser.Text)</td></tr>
        <tr><th>Department / Location</th><td>$($txtDepartment.Text) / $($txtLocation.Text)</td></tr>
        <tr><th>Organization</th><td>$($txtOrganization.Text)</td></tr>
        <tr><th>Staff Name</th><td>$($txtTechnician.Text)</td></tr>
    </table>

    <h2>2. Hardware Inventory</h2>
    <table>
        <tr><th>Processor (CPU)</th><td>$($CPU.Name) ($($CPU.NumberOfCores) Cores / $($CPU.NumberOfLogicalProcessors) Threads)</td></tr>
        <tr><th>Installed RAM</th><td>$RAMGB ($RamSpeed MHz)</td></tr>
        <tr><th>Motherboard</th><td>$MobMfg / $MobModel</td></tr>
        <tr><th>Graphics (GPU)</th><td>$GPUName ($GPUVRAM VRAM)</td></tr>
    </table>

    <h2>3. Storage &amp; Health Diagnostics</h2>
    <table>
        <tr><th>Drive Model</th><th>Serial Number</th><th>Capacity</th><th>Interface</th><th>Type</th></tr>
        $StorageRows
    </table>
    <table>
        <tr><th>Drive Health Target</th><th>SMART Health Status</th><th>Operational Status</th></tr>
        $DiskHealthRows
    </table>

    <h2>4. Operating System &amp; 5. Security Audit</h2>
    <table>
        <tr><th>Operating System</th><td>$($OS.Caption) (Build $($OS.BuildNumber))</td></tr>
        <tr><th>Install Date / TimeZone</th><td>$($OS.InstallDate) / $TimeZone</td></tr>
        <tr><th>Boot Mode / Arch</th><td>$Firmware / $($OS.OSArchitecture)</td></tr>
        <tr><th>Antivirus Installed</th><td>$AVName</td></tr>
        <tr><th>Firewall Profiles Enabled</th><td>$FWStatus</td></tr>
        <tr><th>TPM Present / Secure Boot</th><td>$TPM / $SecureBoot</td></tr>
        <tr><th>BitLocker System Drive</th><td>$BitLocker</td></tr>
        <tr><th>Local Admin Accounts</th><td>$Admins</td></tr>
    </table>

    <h2>6. Network Diagnostics</h2>
    <table>
        <tr><th>Adapter Name</th><th>MAC Address</th><th>IPv4 Address</th><th>Link Speed</th></tr>
        $NetRows
    </table>

    <h2>7. Event Log Analysis (Recent System &amp; App Errors)</h2>
    <table>
        <tr><th style="width:15%">Time</th><th style="width:15%">Source</th><th style="width:10%">ID</th><th style="width:60%">Message</th></tr>
        $EventRows
    </table>

    <h2>8. Software Inventory (Top 10 Recent)</h2>
    <table>
        <tr><th style="width:50%">Software Name</th><th>Version</th><th>Publisher</th></tr>
        $SoftRows
    </table>

    <h2>9. Peripherals &amp; 10. Performance</h2>
    <table>
        <tr><th>Last Boot Time</th><td>$BootTime</td></tr>
        <tr><th>Connected Monitors</th><td>$Monitors</td></tr>
        <tr><th>Default Printer</th><td>$Printers</td></tr>
    </table>
    
    <h2>11. Asset Lifecycle Data</h2>
    <table>
        <tr><th>Purchase Date</th><td>$($txtPurchaseDate.Text)</td></tr>
        <tr><th>Asset Cost</th><td>$($txtAssetCost.Text)</td></tr>
        <tr><th>Invoice Number</th><td>$($txtInvoice.Text)</td></tr>
    </table>

    <!-- Inject Battery Report -->
    <div class="page-break">
        <h2>12. Battery Diagnostics</h2>
        <div class="powercfg">
            $BattContent
        </div>
    </div>

    <!-- Inject Energy Report -->
    <div class="page-break">
        <h2>13. Energy Diagnostics</h2>
        <div class="powercfg">
            $EnergyContent
        </div>
    </div>

    <!-- Terms and Signatures Box at the very end -->
    <div class="page-break">
        <div class="agreement-box">
            <div class="agreement-title">Hardware Asset Terms and Conditions</div>
            <div class="agreement-text">
                By signing below, the Assigned User acknowledges receipt of the IT hardware listed in this document in good working condition. 
                The User agrees to adhere to all corporate IT policies regarding the appropriate use, care, and security of company assets. 
                The User accepts responsibility for the equipment and understands that any deliberate damage or loss may result in disciplinary action 
                or liability. The Staff Member verifies that the asset details and system hardware recorded in this report are accurate at the date and time of generation.
            </div>
            
            <div class="sign-grid">
                <div class="sign-col">
                    <div class="sign-line"></div>
                    <div class="sign-label">Date & Time</div>
                    <div class="sign-name">$DisplayDate</div>
                </div>
                <div class="sign-col">
                    <div class="sign-line"></div>
                    <div class="sign-label">Staff Signature</div>
                    <div class="sign-name">$ValStaff</div>
                </div>
                <div class="sign-col">
                    <div class="sign-line"></div>
                    <div class="sign-label">User Signature</div>
                    <div class="sign-name">$ValUser</div>
                </div>
            </div>
        </div>
        <div style="margin-top: 30px; text-align: center; font-size: 0.85em; color: #777; border-top: 1px solid #eee; padding-top: 15px;">
            App developed by: <strong>Arbor media solutions</strong><br/>
            Secure Report via AssetVault
        </div>
    </div>

</body>
</html>
"@
'@

$Content = Get-Content 'D:\Personal Praison\Develop\Asset manager\AssetManagerPro.ps1' -Raw

# Replace everything from "# System Queries" up to "</html>`n"@""
$Pattern = '(?ms)^        # System Queries.*?</html>\r?\n"@'

$NewContent = $Content -replace $Pattern, $ReplaceBlock
$NewContent | Out-File -FilePath 'D:\Personal Praison\Develop\Asset manager\AssetManagerPro.ps1' -Encoding UTF8

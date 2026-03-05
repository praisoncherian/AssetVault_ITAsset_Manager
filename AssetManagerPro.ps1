# ==============================================================================
# IT ASSET & HARDWARE REPORT GENERATOR PRO (Modern WPF Custom UI)
# ==============================================================================
# Self-Elevating Block
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    if ($PSCommandPath) {
        $Args = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        Start-Process powershell.exe -ArgumentList $Args -Verb RunAs
        Exit
    }
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="AssetVault" Height="680" Width="750" 
        WindowStartupLocation="CenterScreen" 
        WindowStyle="None" AllowsTransparency="True" Background="Transparent" 
        FontFamily="Segoe UI">
        
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="Foreground" Value="#333333"/>
            <Setter Property="Margin" Value="0,8,0,5"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Padding" Value="10"/>
            <Setter Property="Background" Value="#FDFDFD"/>
            <Setter Property="BorderBrush" Value="#C2C8D1"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Margin" Value="0,0,0,10"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Background="{TemplateBinding Background}" CornerRadius="6">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="0"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsFocused" Value="True">
                    <Setter Property="BorderBrush" Value="#005A9E"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    
    <Border Background="#F4F6F8" CornerRadius="12" BorderThickness="1" BorderBrush="#D1D5DB">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="50" />
                <RowDefinition Height="*" />
                <RowDefinition Height="100" />
            </Grid.RowDefinitions>
            
            <!-- Custom Title Bar -->
            <Border Grid.Row="0" Background="#FFFFFF" CornerRadius="12,12,0,0" BorderThickness="0,0,0,1" BorderBrush="#E5E7EB">
                <Grid Name="TitleBar" Background="Transparent" Margin="10,0,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="32" />
                        <ColumnDefinition Width="*" />
                        <ColumnDefinition Width="150" />
                    </Grid.ColumnDefinitions>
                    
                    <Image Name="imgLogo" Grid.Column="0" Width="24" Height="24" VerticalAlignment="Center" HorizontalAlignment="Center"/>
                    <TextBlock Grid.Column="1" Text="AssetVault IT Asset Manager" FontSize="16" FontWeight="SemiBold" Foreground="#1F2937" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="10,0,0,0" />
                    
                    <StackPanel Grid.Column="2" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Top" Margin="0,10,10,0">
                        <Button Name="btnMinimize" Width="30" Height="30" Background="Transparent" BorderThickness="0" Cursor="Hand" ToolTip="Minimize">
                            <Button.Template>
                                <ControlTemplate TargetType="Button">
                                    <Border Name="bg" Background="Transparent" CornerRadius="6">
                                        <TextBlock Text="&#x2012;" FontSize="16" FontWeight="Bold" Foreground="#6B7280" HorizontalAlignment="Center" VerticalAlignment="Center" Margin="0,-5,0,0"/>
                                    </Border>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsMouseOver" Value="True">
                                            <Setter TargetName="bg" Property="Background" Value="#E5E7EB" />
                                            <Setter Property="Foreground" Value="#111827" />
                                        </Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Button.Template>
                        </Button>
                        <Button Name="btnMaximize" Width="30" Height="30" Background="Transparent" BorderThickness="0" Cursor="Hand" ToolTip="Maximize">
                            <Button.Template>
                                <ControlTemplate TargetType="Button">
                                    <Border Name="bg" Background="Transparent" CornerRadius="6">
                                        <TextBlock Text="&#x25A1;" FontSize="16" Foreground="#6B7280" HorizontalAlignment="Center" VerticalAlignment="Center" Margin="0"/>
                                    </Border>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsMouseOver" Value="True">
                                            <Setter TargetName="bg" Property="Background" Value="#E5E7EB" />
                                            <Setter Property="Foreground" Value="#111827" />
                                        </Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Button.Template>
                        </Button>
                        <Button Name="btnClose" Width="30" Height="30" Background="Transparent" BorderThickness="0" Cursor="Hand" ToolTip="Close">
                            <Button.Template>
                                <ControlTemplate TargetType="Button">
                                    <Border Name="bg" Background="Transparent" CornerRadius="6">
                                        <TextBlock Text="&#x2715;" FontSize="16" Foreground="#6B7280" HorizontalAlignment="Center" VerticalAlignment="Center" Margin="0"/>
                                    </Border>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsMouseOver" Value="True">
                                            <Setter TargetName="bg" Property="Background" Value="#FEE2E2" />
                                            <Setter Property="Foreground" Value="#DC2626" />
                                        </Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Button.Template>
                        </Button>
                    </StackPanel>
                </Grid>
            </Border>
            
            <!-- PAGE 1: Data Entry -->
            <Grid Name="MainPage" Grid.Row="1" Visibility="Visible">
                <Grid.RowDefinitions>
                    <RowDefinition Height="*" />
                    <RowDefinition Height="100" />
                </Grid.RowDefinitions>
                
                <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto" Margin="30,20,30,0">
                <StackPanel>
                    <Border BorderThickness="0,0,0,2" BorderBrush="#005A9E" Margin="0,0,0,20" Padding="0,0,0,8">
                        <TextBlock Text="Asset &amp; Staff Information" FontSize="20" FontWeight="SemiBold" Foreground="#111827" Margin="0" />
                    </Border>
                    
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="1*"/>
                            <ColumnDefinition Width="30"/>
                            <ColumnDefinition Width="1*"/>
                        </Grid.ColumnDefinitions>
                        
                        <StackPanel Grid.Column="0">
                            <TextBlock Text="Asset Tag:" FontWeight="SemiBold"/>
                            <TextBox Name="txtAssetTag" />
                            
                            <TextBlock Text="Department:" FontWeight="SemiBold"/>
                            <TextBox Name="txtDepartment" />
                            
                            <TextBlock Text="IT Admin Name:" FontWeight="SemiBold"/>
                            <TextBox Name="txtTechnician" />
                        </StackPanel>
                        
                        <StackPanel Grid.Column="2">
                            <TextBlock Text="Assigned User:" FontWeight="SemiBold"/>
                            <TextBox Name="txtUser" />
                            
                            <TextBlock Text="Location:" FontWeight="SemiBold"/>
                            <TextBox Name="txtLocation" />
                            
                            <TextBlock Text="Organization:" FontWeight="SemiBold"/>
                            <TextBox Name="txtOrganization" />
                        </StackPanel>
                    </Grid>
                    
                    <Border BorderThickness="0,0,0,2" BorderBrush="#005A9E" Margin="0,15,0,20" Padding="0,0,0,8">
                        <TextBlock Text="Asset Lifecycle Data" FontSize="20" FontWeight="SemiBold" Foreground="#111827" Margin="0" />
                    </Border>
                    
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="1*"/>
                            <ColumnDefinition Width="30"/>
                            <ColumnDefinition Width="1*"/>
                        </Grid.ColumnDefinitions>
                        
                        <StackPanel Grid.Column="0">
                            <TextBlock Text="Purchase Date (dd/mm/yyyy):" FontWeight="SemiBold"/>
                            <TextBox Name="txtPurchaseDate" />
                        </StackPanel>
                        
                        <StackPanel Grid.Column="2">
                            <TextBlock Text="Asset Cost (Rs.):" FontWeight="SemiBold"/>
                            <TextBox Name="txtAssetCost" />
                            
                            <TextBlock Text="Invoice Number:" FontWeight="SemiBold"/>
                            <TextBox Name="txtInvoice" />
                        </StackPanel>
                    </Grid>
                    
                </StackPanel>
                </ScrollViewer>
                
                <!-- Footer & Generate Button -->
                <Border Grid.Row="1" Background="#FFFFFF" CornerRadius="0,0,12,12" BorderThickness="0,1,0,0" BorderBrush="#E5E7EB">
                    <StackPanel Orientation="Vertical" VerticalAlignment="Center">
                        <Button Name="btnGenerate" Content="Generate Report" Background="#005A9E" Foreground="White" Width="300" Height="45" FontSize="16" FontWeight="Bold" Cursor="Hand" BorderThickness="0">
                            <Button.Template>
                                <ControlTemplate TargetType="Button">
                                    <Border Background="{TemplateBinding Background}" CornerRadius="8">
                                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                    </Border>
                                </ControlTemplate>
                            </Button.Template>
                        </Button>
                        <TextBlock Name="lblStatus" Text="Ready." Margin="0,8,0,0" HorizontalAlignment="Center" FontSize="13" Foreground="#6B7280"/>
                        <TextBlock Name="lblLicense" Text="" Margin="0,15,0,0" HorizontalAlignment="Center" FontSize="11" Foreground="#9CA3AF"/>
                    </StackPanel>
                </Border>
            </Grid>

            <!-- PAGE 2: Processing Live Log -->
            <Grid Name="ProcessingPage" Grid.Row="1" Visibility="Collapsed" Background="#FFFFFF">
                <StackPanel Margin="40,30,40,30">
                    <TextBlock Text="System Diagnostic Running..." FontSize="24" FontWeight="Bold" Foreground="#1E3A8A" Margin="0,0,0,20"/>
                    <TextBlock Text="Do not close the application. Gathering WMI Data." FontSize="14" Foreground="#64748B" Margin="0,0,0,20"/>
                    
                    <Border BorderThickness="1" BorderBrush="#E2E8F0" CornerRadius="8" Background="#F8FAFC" Padding="15">
                        <StackPanel Name="LogPanel">
                            <TextBlock Name="logAsset" Text="[Pending] Asset Information Query" FontSize="13" Foreground="#475569" Margin="0,0,0,5"/>
                            <TextBlock Name="logHW" Text="[Pending] Core Hardware (CPU/RAM/GPU/Storage)" FontSize="13" Foreground="#475569" Margin="0,0,0,5"/>
                            <TextBlock Name="logHealth" Text="[Pending] System Health &amp; Disk Status" FontSize="13" Foreground="#475569" Margin="0,0,0,5"/>
                            <TextBlock Name="logSec" Text="[Pending] Security Audit &amp; Firewall Status" FontSize="13" Foreground="#475569" Margin="0,0,0,5"/>
                            <TextBlock Name="logNet" Text="[Pending] Network &amp; Peripherals" FontSize="13" Foreground="#475569" Margin="0,0,0,5"/>
                            <TextBlock Name="logPerf" Text="[Pending] Error Logs &amp; Software Inventory" FontSize="13" Foreground="#475569" Margin="0,0,0,5"/>
                            <TextBlock Name="logBatt" Text="[Pending] Exporting Powercfg Reports" FontSize="13" Foreground="#475569" Margin="0,0,0,5"/>
                            <TextBlock Name="logPdf" Text="[Pending] Compiling HTML and rendering PDF Wrapper" FontSize="13" Foreground="#475569" Margin="0,0,0,5"/>
                        </StackPanel>
                    </Border>
                    
                    <Button Name="btnFinish" Content="Back to Main Menu" Background="#10B981" Foreground="White" Width="200" Height="45" FontSize="14" FontWeight="Bold" Cursor="Hand" BorderThickness="0" Margin="0,30,0,0" Visibility="Collapsed">
                        <Button.Template>
                            <ControlTemplate TargetType="Button">
                                <Border Background="{TemplateBinding Background}" CornerRadius="8">
                                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                </Border>
                            </ControlTemplate>
                        </Button.Template>
                    </Button>
                </StackPanel>
            </Grid>
            
        </Grid>
    </Border>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
try {
    $Window = [Windows.Markup.XamlReader]::Load($reader)
}
catch {
    [System.Windows.Forms.MessageBox]::Show("XAML Error: $($_.Exception.Message)", "IT Asset Manager Pro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

# Bind controls
$TitleBar = $Window.FindName("TitleBar")
$btnClose = $Window.FindName("btnClose")
$btnMinimize = $Window.FindName("btnMinimize")
$btnMaximize = $Window.FindName("btnMaximize")
$imgLogo = $Window.FindName("imgLogo")
$MainPage = $Window.FindName("MainPage")
$ProcessingPage = $Window.FindName("ProcessingPage")
$btnFinish = $Window.FindName("btnFinish")
$logAsset = $Window.FindName("logAsset")
$logHW = $Window.FindName("logHW")
$logHealth = $Window.FindName("logHealth")
$logSec = $Window.FindName("logSec")
$logNet = $Window.FindName("logNet")
$logPerf = $Window.FindName("logPerf")
$logBatt = $Window.FindName("logBatt")
$logPdf = $Window.FindName("logPdf")

$txtAssetTag = $Window.FindName("txtAssetTag")
$txtDepartment = $Window.FindName("txtDepartment")
$txtUser = $Window.FindName("txtUser")
$txtLocation = $Window.FindName("txtLocation")
$txtOrganization = $Window.FindName("txtOrganization")
$txtPurchaseDate = $Window.FindName("txtPurchaseDate")
$txtAssetCost = $Window.FindName("txtAssetCost")
$txtInvoice = $Window.FindName("txtInvoice")
$txtTechnician = $Window.FindName("txtTechnician")
$btnGenerate = $Window.FindName("btnGenerate")
$lblStatus = $Window.FindName("lblStatus")
$lblLicense = $Window.FindName("lblLicense")

$lblLicense.Text = "AssetVault IT Asset Manager Platform developed by Arbor Solutions."

# Window Handling Hooks
$Window.ResizeMode = 'CanResizeWithGrip'

$TitleBar.Add_MouseLeftButtonDown({
        $Window.DragMove()
    })

$btnClose.Add_Click({ $Window.Close() })

$btnMinimize.Add_Click({ $Window.WindowState = 'Minimized' })

$btnMaximize.Add_Click({ 
        if ($Window.WindowState -eq 'Maximized') { $Window.WindowState = 'Normal' }
        else { $Window.WindowState = 'Maximized' }
    })

$btnFinish.Add_Click({
        $ProcessingPage.Visibility = 'Collapsed'
        $btnFinish.Visibility = 'Collapsed'
        $MainPage.Visibility = 'Visible'
        $btnGenerate.IsEnabled = $true
    })

$txtAssetCost.Add_PreviewTextInput({
        param($sender, $e)
        if ($e.Text -notmatch '^\d+$') {
            $e.Handled = $true
        }
    })

# PDF Generator Edge Helper
function Export-HtmlToPdf($HtmlPath, $PdfPath) {
    $EdgePaths = @(
        "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        "C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
    )
    $EdgeExe = $EdgePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    
    if (-not $EdgeExe) { throw "Microsoft Edge is required to generate the PDF but was not found." }
    
    $TempDir = Join-Path $env:TEMP "EdgePdf_$(Get-Random)"
    
    $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
    $ProcessInfo.FileName = $EdgeExe
    $ProcessInfo.Arguments = "--headless --disable-gpu --no-pdf-header-footer --print-to-pdf=`"$PdfPath`" --user-data-dir=`"$TempDir`" `"file:///$($HtmlPath -replace '\\','/')`""
    $ProcessInfo.CreateNoWindow = $true
    $ProcessInfo.UseShellExecute = $false
    $Process = [System.Diagnostics.Process]::Start($ProcessInfo)
    $Process.WaitForExit(30000)
    
    if (Test-Path $TempDir) { Remove-Item $TempDir -Recurse -Force -ErrorAction SilentlyContinue }

    if ($Process.ExitCode -ne 0) {
        throw "Edge failed to generate PDF. Exit code: $($Process.ExitCode)"
    }
}

$btnGenerate.Add_Click({
        $btnGenerate.IsEnabled = $false
    
        # Date Validation
        $purchaseDateRaw = $txtPurchaseDate.Text.Trim()
        $parsedDate = $null
        if ($purchaseDateRaw -ne "") {
            try {
                $parsedDate = [datetime]::ParseExact($purchaseDateRaw, 'dd/MM/yyyy', $null)
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Invalid Purchase Date format. Please use exactly dd/mm/yyyy (e.g. 15/05/2021).", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $btnGenerate.IsEnabled = $true
                return
            }
        }

        # Flip UI to Processing Page
        $MainPage.Visibility = 'Collapsed'
        $ProcessingPage.Visibility = 'Visible'
        $logAsset.Text = "[Processing...] Asset Information Query"
        $logAsset.Foreground = "#CA8A04"
        $logHW.Text = "[Pending] Core Hardware"
        $logHW.Foreground = "#475569"
        $logHealth.Text = "[Pending] System Health"
        $logHealth.Foreground = "#475569"
        $logSec.Text = "[Pending] Security"
        $logSec.Foreground = "#475569"
        $logNet.Text = "[Pending] Network"
        $logNet.Foreground = "#475569"
        $logPerf.Text = "[Pending] Perf Logs"
        $logPerf.Foreground = "#475569"
        $logBatt.Text = "[Pending] Powercfg Export"
        $logBatt.Foreground = "#475569"
        $logPdf.Text = "[Pending] PDF Compilation"
        $logPdf.Foreground = "#475569"

        $Dispatcher = [System.Windows.Threading.Dispatcher]::CurrentDispatcher
        $RefreshUI = { $Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Background) }
        & $RefreshUI
    
        Try {
            $ComputerName = $env:COMPUTERNAME
            $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $DisplayDate = Get-Date -Format "dd-MMM-yyyy hh:mm tt"
        
            # Setup Folders
            $DownloadsPath = [System.Environment]::GetFolderPath('UserProfile') + "\Downloads\AssetVault"
            $CachePath = "$env:TEMP\IT_Asset_Cache"
        
            if (-not (Test-Path $DownloadsPath)) { New-Item -ItemType Directory -Force -Path $DownloadsPath | Out-Null }
            if (-not (Test-Path $CachePath)) { New-Item -ItemType Directory -Force -Path $CachePath | Out-Null }
        
            $SafeAssetTag = $txtAssetTag.Text.Trim() -replace '[\\/:\*\?"<>\|]', '_'
            if ($SafeAssetTag -eq "") { $SafeAssetTag = "Asset" }
            $TimestampStrict = Get-Date -Format "ddMMyyyyHHmmss"

            $HtmlFile = "$CachePath\${SafeAssetTag}_${TimestampStrict}.html"
            $PdfFile = "$DownloadsPath\${SafeAssetTag}_${TimestampStrict}.pdf"

            # ==========================================
            # 1. ASSET INFORMATION
            # ==========================================
            $OS = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
            $System = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
            $BIOS = Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue
        
            $logAsset.Text = "[COMPLETED] Asset Information Extracted!"
            $logAsset.Foreground = "#16A34A"
            $logHW.Text = "[Processing...] Querying Hardware Metrics"
            $logHW.Foreground = "#CA8A04"
            & $RefreshUI

            # ==========================================
            # 2. HARDWARE INVENTORY
            # ==========================================
            $CPU = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue
            $RAMBytes = (Get-CimInstance Win32_PhysicalMemory -ErrorAction SilentlyContinue | Measure-Object Capacity -Sum).Sum
            $RAMGB = if ($null -ne $RAMBytes) { "{0:N2} GB" -f ($RAMBytes / 1GB) } else { "Unknown" }
            $GPUName = "Unknown"; $GPUVRAM = "Unknown"; $GPUDriver = "Unknown"
            Try {
                $GPU = Get-CimInstance Win32_VideoController
                $GPUName = ($GPU.Name) -join ", "
                $GPUVRAM = ($GPU | ForEach-Object { "{0:N2} GB" -f ($_.AdapterRAM / 1GB) }) -join ", "
                $GPUDriver = ($GPU.DriverVersion) -join ", "
            }
            Catch {}
        
            $MobMfg = "Unknown"; $MobModel = "Unknown"
            Try { $Mob = Get-CimInstance Win32_BaseBoard; $MobMfg = $Mob.Manufacturer; $MobModel = $Mob.Product } Catch {}
        
            $RamSpeed = "Unknown"; Try { $RamSpeed = (Get-CimInstance Win32_PhysicalMemory | Select -First 1).Speed } Catch {}
        
            $StorageRows = ""
            Get-CimInstance Win32_DiskDrive -ErrorAction SilentlyContinue | ForEach-Object {
                $sz = "{0:N2}" -f ($_.Size / 1GB)
                $StorageRows += "<tr><td>$($_.Model)</td><td>$($_.SerialNumber)</td><td>$sz GB</td><td>$($_.InterfaceType)</td><td>$($_.MediaType)</td></tr>"
            }
        
            $logHW.Text = "[COMPLETED] Hardware Inventory Scanned!"
            $logHW.Foreground = "#16A34A"
            $logHealth.Text = "[Processing...] Storage and System Health Rules"
            $logHealth.Foreground = "#CA8A04"
            & $RefreshUI

            # ==========================================
            # 3. HARDWARE HEALTH DIAGNOSTICS
            # ==========================================
            $DiskHealthRows = ""
            Try {
                Get-PhysicalDisk -ErrorAction SilentlyContinue | ForEach-Object {
                    $DiskHealthRows += "<tr><td>$($_.FriendlyName)</td><td>$($_.HealthStatus)</td><td>$($_.OperationalStatus)</td></tr>"
                }
            }
            Catch { $DiskHealthRows = "<tr><td colspan='3'>Disk health queries not supported or run without admin rights.</td></tr>" }
        
            $logHealth.Text = "[COMPLETED] Storage Health Analyzed!"
            $logHealth.Foreground = "#16A34A"
            $logSec.Text = "[Processing...] Security Audit (TPM / BitLocker)"
            $logSec.Foreground = "#CA8A04"
            & $RefreshUI

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
            }
            Catch {}
        
            # ==========================================
            # 5. OS STATUS & 6. SECURITY AUDIT
            # ==========================================
            $Firmware = $env:firmware_type
            $TimeZone = (Get-TimeZone).Id
        
            $BitLocker = "Not Checked"
            Try { $BitLocker = (Get-BitLockerVolume -MountPoint 'C:' -ErrorAction SilentlyContinue).VolumeStatus } Catch {}
        
            $Admins = "Unknown"
            Try { $Admins = (Get-LocalGroupMember -Group "Administrators" | Where-Object PrincipalSource -eq "Local" | Select-Object -ExpandProperty Name) -join ", " } Catch {}
        
            $logSec.Text = "[COMPLETED] Security Data Collected!"
            $logSec.Foreground = "#16A34A"
            $logNet.Text = "[Processing...] Auditing Networks and Interfaces"
            $logNet.Foreground = "#CA8A04"
            & $RefreshUI

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
            }
            Catch {}

            $logNet.Text = "[COMPLETED] Network Configurations Secured!"
            $logNet.Foreground = "#16A34A"
            $logPerf.Text = "[Processing...] Perf Logs & Software Inventory"
            $logPerf.Foreground = "#CA8A04"
            & $RefreshUI

            # ==========================================
            # 8. PERFORMANCE & 9. EVENT LOGS
            # ==========================================
            $BootTime = $OS.LastBootUpTime
        
            $EventRows = ""
            Try {
                Get-WinEvent -FilterHashtable @{LogName = 'System', 'Application'; Level = 2 } -MaxEvents 5 -ErrorAction SilentlyContinue | ForEach-Object {
                    $EventRows += "<tr><td>$($_.TimeCreated)</td><td>$($_.ProviderName)</td><td>$($_.Id)</td><td style='font-size: 11px;'>$($_.Message -replace "`n", " ")</td></tr>"
                }
            }
            Catch { $EventRows = "<tr><td colspan='4'>No recent critical events or access denied.</td></tr>" }

            # ==========================================
            # 11. PERIPHERALS
            # ==========================================
            $Monitors = "Unknown"; Try { $Monitors = (Get-CimInstance WmiMonitorID -Namespace root\wmi | ForEach-Object { [System.Text.Encoding]::ASCII.GetString($_.UserFriendlyName) -replace "`0", "" }) -join ", " } Catch {}
            $Printers = "Unknown"; Try { $Printers = (Get-CimInstance Win32_Printer | Where-Object Default -eq $true | Select-Object -ExpandProperty Name) -join ", " } Catch {}

            $logPerf.Text = "[COMPLETED] Performance Rules Captured!"
            $logPerf.Foreground = "#16A34A"
            $logBatt.Text = "[Processing...] Compiling PowerCfg Diagnostics"
            $logBatt.Foreground = "#CA8A04"
            & $RefreshUI

            # ==========================================
            # 10. BATTERY & POWER (Powercfg execution)
            # ==========================================
            $BattHtml = "$CachePath\batt_temp_$Timestamp.html"
            $EnergyHtml = "$CachePath\energy_temp_$Timestamp.html"
        
            Start-Process powercfg -ArgumentList "/batteryreport /output `"$BattHtml`"" -NoNewWindow -Wait
            Start-Process powercfg -ArgumentList "/energy /output `"$EnergyHtml`"" -NoNewWindow -Wait
        
            $BattContent = "<p>Battery report unavailable.</p>"
            $EnergyContent = "<p>Energy report unavailable.</p>"
            if (Test-Path $BattHtml) { 
                $raw = Get-Content $BattHtml -Raw
                if ($raw -match '(?si)<body[^>]*>(.*?)</body>') { 
                    $BattContent = $matches[1]
                    $splitPoint = $BattContent.IndexOf("<h2>Recent usage</h2>")
                    if ($splitPoint -gt -1) { $BattContent = $BattContent.Substring(0, $splitPoint) }
                } 
                Remove-Item $BattHtml -Force 
            }
            if (Test-Path $EnergyHtml) { $raw = Get-Content $EnergyHtml -Raw; if ($raw -match '(?si)<body[^>]*>(.*?)</body>') { $EnergyContent = $matches[1] }; Remove-Item $EnergyHtml -Force }

            $ValStaff = if ($txtTechnician.Text.Trim() -ne "") { $txtTechnician.Text } else { "_________________" }
            $ValUser = if ($txtUser.Text.Trim() -ne "") { $txtUser.Text } else { "_________________" }

            # Asset Age Calculator
            $AgeString = "N/A"
            $ReplaceString = "N/A"
            if ($null -ne $parsedDate) {
                $timespan = (Get-Date) - $parsedDate
                $totalMonths = [math]::Floor($timespan.TotalDays / 30.436875)
                $years = [math]::Floor($totalMonths / 12)
                $months = $totalMonths % 12
                $AgeString = "$years Years, $months Months"
            
                $targetMonths = 60 - $totalMonths # 5 year standard
                if ($targetMonths -lt 0) {
                    $ReplaceString = "Overdue ($([math]::Abs($targetMonths)) months past)"
                }
                else {
                    $r_years = [math]::Floor($targetMonths / 12)
                    $r_months = $targetMonths % 12
                    $ReplaceString = "$r_years Years, $r_months Months remaining"
                }
            }

            # Injecting Logo Base64
            $AppPathX = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
            $IconX = [System.Drawing.Icon]::ExtractAssociatedIcon($AppPathX)
            $BmpX = $IconX.ToBitmap()
            $MsX = New-Object System.IO.MemoryStream
            $BmpX.Save($MsX, [System.Drawing.Imaging.ImageFormat]::Png)
            $B64Logo = [Convert]::ToBase64String($MsX.ToArray())

            $HtmlBody = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title>AssetVault IT Hardware Report</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Verdana, sans-serif; margin: 40px; color: #333; line-height: 1.5; font-size: 13px; }
        .header-container { display: flex; align-items: center; justify-content: center; margin-bottom: 5px; }
        .header-logo { width: 48px; height: 48px; margin-right: 15px; }
        h1 { text-align: center; color: #1E3A8A; margin: 0; font-size: 28px; }
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

        .powercfg { zoom: 0.65; margin: 0; border: 1px solid #eee; padding: 0; background: #fdfdfd; width: 100%; overflow-x: hidden; }
        .powercfg body { margin: 0; padding: 0; }
        .page-break { page-break-before: always; }
    </style>
</head>
<body>

    <div class="header-container">
        <img class="header-logo" src="data:image/png;base64,$B64Logo" alt="AssetVault Logo"/>
        <h1>AssetVault</h1>
    </div>
    <div class="subtitle">Diagnostic &amp; Hardware Report</div>
    
    <h2>1. Asset Information</h2>
    <table>
        <tr><th>Report Generated</th><td>$DisplayDate (Local Time)</td></tr>
        <tr><th>Computer Name / Domain</th><td>$ComputerName / $Domain</td></tr>
        <tr><th>Asset Tag</th><td>$($txtAssetTag.Text)</td></tr>
        <tr><th>Manufacturer</th><td>$($System.Manufacturer)</td></tr>
        <tr><th>Model</th><td>$($System.Model)</td></tr>
        <tr><th>Serial Number</th><td>$($BIOS.SerialNumber)</td></tr>
        <tr><th>UUID</th><td>$UUID</td></tr>
        <tr><th>Chassis Type</th><td>$Chassis</td></tr>
        <tr><th>Assigned User</th><td>$($txtUser.Text)</td></tr>
        <tr><th>Department / Location</th><td>$($txtDepartment.Text) / $($txtLocation.Text)</td></tr>
        <tr><th>Organization</th><td>$($txtOrganization.Text)</td></tr>
        <tr><th>Staff Name</th><td>$($txtTechnician.Text)</td></tr>
    </table>

    <h2>2. Hardware Inventory</h2>
    <table>
        <tr><th>Processor (CPU)</th><td>$($CPU.Name)</td></tr>
        <tr><th>Cores / Threads</th><td>$($CPU.NumberOfCores) Cores / $($CPU.NumberOfLogicalProcessors) Threads</td></tr>
        <tr><th>Installed RAM</th><td>$RAMGB ($RamSpeed MHz)</td></tr>
        <tr><th>Motherboard</th><td>$MobMfg / $MobModel</td></tr>
        <tr><th>Graphics (GPU)</th><td>$GPUName ($GPUVRAM VRAM)</td></tr>
    </table>

    <h2>3. Storage &amp; Health Diagnostics</h2>
    <table>
        <tr><th style="width:25%">Drive Model</th><th style="width:25%">Serial Number</th><th style="width:20%">Capacity</th><th style="width:15%">Interface</th><th style="width:15%">Type</th></tr>
        $StorageRows
    </table>
    <table>
        <tr><th style="width:33.3%">Drive Health Target</th><th style="width:33.3%">SMART Health Status</th><th style="width:33.3%">Operational Status</th></tr>
        $DiskHealthRows
    </table>

    <h2>4. Operating System Status</h2>
    <table>
        <tr><th>Operating System</th><td>$($OS.Caption)</td></tr>
        <tr><th>OS Build</th><td>$($OS.BuildNumber)</td></tr>
        <tr><th>Install Date / TimeZone</th><td>$($OS.InstallDate) / $TimeZone</td></tr>
        <tr><th>Boot Mode / Arch</th><td>$Firmware / $($OS.OSArchitecture)</td></tr>
    </table>

    <h2>5. Security Audit</h2>
    <table>
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

    <h2>9. Peripherals Inventory</h2>
    <table>
        <tr><th>Connected Monitors</th><td>$Monitors</td></tr>
        <tr><th>Default Printer</th><td>$Printers</td></tr>
    </table>

    <h2>10. Performance Metrics</h2>
    <table>
        <tr><th>Last Boot Time</th><td>$BootTime</td></tr>
    </table>

    <!-- Inject Battery Report -->
    <div class="page-break">
        <h2>11. Battery Diagnostics</h2>
        <div class="powercfg">
            $BattContent
        </div>
    </div>

    <!-- Inject Energy Report -->
    <div class="page-break">
        <h2>12. Energy Diagnostics</h2>
        <div class="powercfg">
            $EnergyContent
        </div>
    </div>

    <!-- Terms and Signatures Box at the very end -->
    <div class="page-break">
        <h2>13. Asset Lifecycle Data</h2>
        <table>
            <tr><th>Purchase Date</th><td>$($txtPurchaseDate.Text)</td></tr>
            <tr><th>Asset Cost</th><td>₹ $($txtAssetCost.Text)</td></tr>
            <tr><th>Invoice Number</th><td>$($txtInvoice.Text)</td></tr>
            <tr><th>System Age</th><td>$AgeString</td></tr>
            <tr><th>Time To Replacement Target</th><td>$ReplaceString</td></tr>
        </table>
        <div class="agreement-box">
            <div class="agreement-title">USER ACKNOWLEDGEMENT</div>
            <div class="agreement-text">
                <p>By signing or accepting this report, the user confirms that:</p>
                <ul>
                    <li>The device has been received and recorded correctly.</li>
                    <li>The system information captured in this report is acknowledged.</li>
                    <li>The user understands the responsibilities associated with the assigned IT asset.</li>
                    <li>The user agrees to comply with the organization’s IT asset usage policies.</li>
                </ul>
            </div>
            
            <div class="sign-grid">
                <div class="sign-col" style="width: 50%;">
                    <div class="sign-line"></div>
                    <div class="sign-label">User Signature</div>
                    <div class="sign-name">$ValUser</div>
                </div>
                <div class="sign-col" style="width: 50%;">
                    <div class="sign-line"></div>
                    <div class="sign-label">IT Administrator Signature</div>
                    <div class="sign-name">$ValStaff</div>
                </div>
            </div>
        </div>
        <div style="margin-top: 30px; text-align: center; font-size: 0.85em; color: #777; border-top: 1px solid #eee; padding-top: 15px;">
            Report generated by $ValStaff  $DisplayDate<br/>
            Secure Report via AssetVault IT Asset Management Platform developed by Arbor Solutions
        </div>
    </div>

</body>
</html>
"@


            $HtmlBody | Out-File -FilePath $HtmlFile -Encoding UTF8
        
            # Edge needs the file URI to be absolute and safe for local HTML
            $SafeHtmlPath = $HtmlFile
            Export-HtmlToPdf $SafeHtmlPath $PdfFile
        
            $RetryCounter = 10
            while (-not (Test-Path $PdfFile) -and $RetryCounter -gt 0) {
                Start-Sleep -Seconds 1
                $RetryCounter--
            }
        
            if (Test-Path $PdfFile) {
                $logPdf.Text = "✅ [SUCCESS] Saved PDF to Downloads folder!`nPath: $PdfFile"
                $logPdf.Foreground = "#16A34A"
                $btnFinish.Visibility = 'Visible'
            
                Start-Process $PdfFile -ErrorAction SilentlyContinue
            }
            else {
                $logPdf.Text = "❌ [FAILED] PDF generation failed."
                $logPdf.Foreground = "#DC2626"
                $btnFinish.Visibility = 'Visible'
            }

        }
        Catch {
            $logPdf.Text = "❌ [ERROR] $_"
            $logPdf.Foreground = "#DC2626"
            $btnFinish.Visibility = 'Visible'
        }
    })

Try {
    Add-Type -AssemblyName System.Drawing
    $AppPath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($AppPath)
    $Bmp = $Icon.ToBitmap()
    $Ms = New-Object System.IO.MemoryStream
    $Bmp.Save($Ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $ImageSource = New-Object System.Windows.Media.Imaging.BitmapImage
    $ImageSource.BeginInit()
    $ImageSource.StreamSource = $Ms
    $ImageSource.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
    $ImageSource.EndInit()
    $imgLogo.Source = $ImageSource
    $Window.Icon = $ImageSource
}
Catch {}

$Window.ShowDialog() | Out-Null



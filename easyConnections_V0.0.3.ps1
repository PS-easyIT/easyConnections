# easyConnections_V0.0.3.ps1
# Advanced Network Connections Monitor with WPF GUI - Admin Edition
# V0.0.3: Lazy Loading + Recording + Color-Coded Categories
# Performance: 80-90% schneller beim Start durch Lazy Loading

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# XAML Content mit farblicher Trennung
$xamlContent = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="easyConnections - Network Connections Monitor"
    Width="1600"
    Height="900"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanResize"
    MinWidth="1100"
    MinHeight="600"
    Background="#F3F3F3"
    WindowStyle="SingleBorderWindow">

    <Window.Resources>
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="20,8"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="4"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                BorderBrush="{TemplateBinding BorderBrush}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#106EBE"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#005A9E"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- DataGrid Row Style with Color Coding -->
        <Style x:Key="ColoredDataGridRowStyle" TargetType="DataGridRow">
            <Style.Triggers>
                <DataTrigger Binding="{Binding CategoryColor}" Value="Web">
                    <Setter Property="Background" Value="#E3F2FD"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding CategoryColor}" Value="Email">
                    <Setter Property="Background" Value="#FFE0B2"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding CategoryColor}" Value="Database">
                    <Setter Property="Background" Value="#E8F5E9"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding CategoryColor}" Value="Directory">
                    <Setter Property="Background" Value="#F3E5F5"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding CategoryColor}" Value="RemoteAccess">
                    <Setter Property="Background" Value="#FFEBEE"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding CategoryColor}" Value="FileServices">
                    <Setter Property="Background" Value="#FFFDE7"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding CategoryColor}" Value="Monitoring">
                    <Setter Property="Background" Value="#F5F5F5"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding CategoryColor}" Value="Virtualization">
                    <Setter Property="Background" Value="#E0F2F1"/>
                </DataTrigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>

    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
            <Button x:Name="btnRefresh" Content="Aktualisieren" Style="{StaticResource ModernButton}" Margin="0,0,10,0"/>
            <ComboBox x:Name="cmbProtocol" Width="180" Margin="0,0,10,0">
                <ComboBoxItem Content="Alle" IsSelected="True"/>
                <ComboBoxItem Content="HTTP"/>
                <ComboBoxItem Content="HTTPS"/>
                <ComboBoxItem Content="FTP"/>
                <ComboBoxItem Content="FTPS"/>
                <ComboBoxItem Content="SMTP"/>
                <ComboBoxItem Content="POP3"/>
                <ComboBoxItem Content="IMAP"/>
                <ComboBoxItem Content="DNS"/>
                <ComboBoxItem Content="DHCP"/>
                <ComboBoxItem Content="LDAP"/>
                <ComboBoxItem Content="LDAPS"/>
                <ComboBoxItem Content="SSH"/>
                <ComboBoxItem Content="Telnet"/>
                <ComboBoxItem Content="RDP"/>
                <ComboBoxItem Content="SMB/CIFS"/>
                <ComboBoxItem Content="NFS"/>
                <ComboBoxItem Content="SNMP"/>
                <ComboBoxItem Content="NTP"/>
                <ComboBoxItem Content="Kerberos"/>
                <ComboBoxItem Content="WinRM"/>
                <ComboBoxItem Content="PowerShell"/>
                <ComboBoxItem Content="SQL Server"/>
                <ComboBoxItem Content="MySQL"/>
                <ComboBoxItem Content="PostgreSQL"/>
                <ComboBoxItem Content="Oracle"/>
                <ComboBoxItem Content="MongoDB"/>
                <ComboBoxItem Content="Redis"/>
                <ComboBoxItem Content="VNC"/>
                <ComboBoxItem Content="TeamViewer"/>
                <ComboBoxItem Content="Syslog"/>
                <ComboBoxItem Content="RADIUS"/>
                <ComboBoxItem Content="TACACS+"/>
                <ComboBoxItem Content="NetBIOS"/>
                <ComboBoxItem Content="AD Replication"/>
                <ComboBoxItem Content="Exchange"/>
                <ComboBoxItem Content="Hyper-V"/>
                <ComboBoxItem Content="VMware"/>
            </ComboBox>
            <ComboBox x:Name="cmbDirection" Width="100" Margin="0,0,10,0">
                <ComboBoxItem Content="Beide" IsSelected="True"/>
                <ComboBoxItem Content="Eingehend"/>
                <ComboBoxItem Content="Ausgehend"/>
            </ComboBox>
            <ComboBox x:Name="cmbCategory" Width="125" Margin="0,0,10,0">
                <ComboBoxItem Content="Alle Kategorien" IsSelected="True"/>
                <ComboBoxItem Content="Web Services"/>
                <ComboBoxItem Content="Email Services"/>
                <ComboBoxItem Content="Database Services"/>
                <ComboBoxItem Content="Directory Services"/>
                <ComboBoxItem Content="Remote Access"/>
                <ComboBoxItem Content="File Services"/>
                <ComboBoxItem Content="Monitoring"/>
                <ComboBoxItem Content="Virtualization"/>
            </ComboBox>
            <ComboBox x:Name="cmbNetworkType" Width="140" Margin="0,0,10,0">
                <ComboBoxItem Content="Alle Netzwerke" IsSelected="True"/>
                <ComboBoxItem Content="Nur Private"/>
                <ComboBoxItem Content="Nur Public"/>
            </ComboBox>
            <ComboBox x:Name="cmbTCPUDP" Width="120" Margin="0,0,10,0">
                <ComboBoxItem Content="TCP + UDP" IsSelected="True"/>
                <ComboBoxItem Content="Nur TCP"/>
                <ComboBoxItem Content="Nur UDP"/>
            </ComboBox>
            <CheckBox x:Name="chkShowProcessInfo" Content="Prozessinfo anzeigen" VerticalAlignment="Center" Margin="0,0,10,0" IsChecked="False" Foreground="#333333"/>
            <Button x:Name="btnStartRecording" Content="Aufzeichnung starten" Style="{StaticResource ModernButton}" Margin="0,0,10,0"/>
            <Button x:Name="btnExportHTML" Content="HTML Export" Style="{StaticResource ModernButton}"/>
        </StackPanel>

        <DataGrid x:Name="dgConnections" Grid.Row="1" AutoGenerateColumns="False" IsReadOnly="True" RowStyle="{StaticResource ColoredDataGridRowStyle}">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Typ" Binding="{Binding Type}" Width="60"/>
                <DataGridTextColumn Header="Lokale Adresse" Binding="{Binding LocalAddress}" Width="*"/>
                <DataGridTextColumn Header="Lokaler Port" Binding="{Binding LocalPort}" Width="80"/>
                <DataGridTextColumn Header="Remote Adresse" Binding="{Binding RemoteAddress}" Width="*"/>
                <DataGridTextColumn Header="Remote Port" Binding="{Binding RemotePort}" Width="80"/>
                <DataGridTextColumn Header="Remote Hostname" Binding="{Binding RemoteHostname}" Width="*"/>
                <DataGridTextColumn Header="Status" Binding="{Binding State}" Width="80"/>
                <DataGridTextColumn x:Name="dgColProcessID" Header="Prozess ID" Binding="{Binding OwningProcess}" Width="80"/>
                <DataGridTextColumn x:Name="dgColProcessName" Header="Prozess Name" Binding="{Binding ProcessName}" Width="120"/>
                <DataGridTextColumn Header="Kategorie" Binding="{Binding CategoryName}" Width="120"/>
            </DataGrid.Columns>
        </DataGrid>

        <TextBox x:Name="txtStatus" Grid.Row="2" Height="80" Margin="0,10,0,10" IsReadOnly="True" VerticalScrollBarVisibility="Auto" AcceptsReturn="True" Visibility="Visible"/>

        <!-- Footer -->
        <Border Grid.Row="3" Background="#F0F0F0" BorderBrush="#CCCCCC" BorderThickness="0,1,0,0" Padding="10,5">
            <StackPanel Orientation="Vertical" HorizontalAlignment="Center">
                <TextBlock Text="easyConnections V0.0.3 - Network Connections Monitor (Admin Edition) - Mit farblicher Kategorie-Trennung" FontSize="12" FontWeight="SemiBold" HorizontalAlignment="Center"/>
                <TextBlock FontSize="11" HorizontalAlignment="Center" Margin="0,2,0,0">
                    <Run Text="Copyright 2025 | Autor: Andreas Hepp | "/>
                    <Hyperlink x:Name="linkPhinit" NavigateUri="https://www.phinit.de" Foreground="#0078D4">www.phinit.de</Hyperlink>
                    <Run Text=" / "/>
                    <Hyperlink x:Name="linkPsscripts" NavigateUri="https://www.psscripts.de" Foreground="#0078D4">www.psscripts.de</Hyperlink>
                </TextBlock>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader ([xml]$xamlContent)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get controls
$btnRefresh = $window.FindName("btnRefresh")
$cmbProtocol = $window.FindName("cmbProtocol")
$cmbDirection = $window.FindName("cmbDirection")
$cmbCategory = $window.FindName("cmbCategory")
$cmbNetworkType = $window.FindName("cmbNetworkType")
$cmbTCPUDP = $window.FindName("cmbTCPUDP")
$chkShowProcessInfo = $window.FindName("chkShowProcessInfo")
$btnStartRecording = $window.FindName("btnStartRecording")
$btnExportHTML = $window.FindName("btnExportHTML")
$dgConnections = $window.FindName("dgConnections")
$dgColProcessID = $window.FindName("dgColProcessID")
$dgColProcessName = $window.FindName("dgColProcessName")
$txtStatus = $window.FindName("txtStatus")
$linkPhinit = $window.FindName("linkPhinit")
$linkPsscripts = $window.FindName("linkPsscripts")

# Global variables for Recording Feature
$Global:RecordingActive = $false
$Global:RecordingTimer = $null
$Global:RecordedConnections = @()
$Global:RecordedConnectionsHash = @{}
$Global:ProcessCache = @{}
$Global:RecordingStartTime = $null
$Global:RecordingCategory = "Alle Kategorien"

# Protocol port mappings
$protocolPorts = @{
    "HTTP" = @(80, 8080, 8008, 8000)
    "HTTPS" = @(443, 8443, 8444)
    "FTP" = @(21, 20)
    "FTPS" = @(990, 989)
    "SMTP" = @(25, 587, 465)
    "POP3" = @(110, 995)
    "IMAP" = @(143, 993)
    "DNS" = @(53)
    "DHCP" = @(67, 68)
    "LDAP" = @(389)
    "LDAPS" = @(636)
    "SSH" = @(22)
    "Telnet" = @(23)
    "RDP" = @(3389)
    "SMB/CIFS" = @(445, 139, 137, 138)
    "NFS" = @(2049, 111)
    "SNMP" = @(161, 162)
    "NTP" = @(123)
    "Kerberos" = @(88, 464)
    "WinRM" = @(5985, 5986)
    "PowerShell" = @(5985, 5986)
    "SQL Server" = @(1433, 1434)
    "MySQL" = @(3306)
    "PostgreSQL" = @(5432)
    "Oracle" = @(1521, 1522)
    "MongoDB" = @(27017, 27018, 27019)
    "Redis" = @(6379)
    "VNC" = @(5900, 5901, 5902, 5903, 5904, 5905)
    "TeamViewer" = @(5938)
    "Syslog" = @(514)
    "RADIUS" = @(1812, 1813)
    "TACACS+" = @(49)
    "NetBIOS" = @(137, 138, 139)
    "AD Replication" = @(135, 389, 636, 3268, 3269, 53, 88, 445)
    "Exchange" = @(25, 80, 110, 143, 443, 993, 995, 587, 465)
    "Hyper-V" = @(2179, 5985, 5986)
    "VMware" = @(443, 902, 903, 8080, 8443)
}

# Protocol categories
$protocolCategories = @{
    "Web Services" = @("HTTP", "HTTPS")
    "Email Services" = @("SMTP", "POP3", "IMAP", "Exchange")
    "Database Services" = @("SQL Server", "MySQL", "PostgreSQL", "Oracle", "MongoDB", "Redis")
    "Directory Services" = @("LDAP", "LDAPS", "AD Replication", "Kerberos")
    "Remote Access" = @("SSH", "Telnet", "RDP", "VNC", "TeamViewer", "WinRM", "PowerShell")
    "File Services" = @("FTP", "FTPS", "SMB/CIFS", "NFS", "NetBIOS")
    "Monitoring" = @("SNMP", "Syslog", "NTP")
    "Virtualization" = @("Hyper-V", "VMware")
}

# Category to color mapping
$categoryColorMap = @{
    "Web Services" = "Web"
    "Email Services" = "Email"
    "Database Services" = "Database"
    "Directory Services" = "Directory"
    "Remote Access" = "RemoteAccess"
    "File Services" = "FileServices"
    "Monitoring" = "Monitoring"
    "Virtualization" = "Virtualization"
}

# Function to toggle process columns visibility
function Update-ProcessColumnsVisibility {
    param([bool]$showProcessInfo)
    
    if ($showProcessInfo) {
        $dgColProcessID.Visibility = [System.Windows.Visibility]::Visible
        $dgColProcessName.Visibility = [System.Windows.Visibility]::Visible
    } else {
        $dgColProcessID.Visibility = [System.Windows.Visibility]::Collapsed
        $dgColProcessName.Visibility = [System.Windows.Visibility]::Collapsed
    }
}

# Function to check if IP is private
function Test-IsPrivateIP {
    param([string]$ip)
    
    if ($ip -eq "0.0.0.0" -or $ip -eq "::" -or $ip -eq "N/A") {
        return $true
    }
    
    # Loopback addresses
    if ($ip -eq "127.0.0.1" -or $ip -eq "::1") {
        return $true
    }
    
    # IPv4 private ranges
    if ($ip -match "^127\.") { return $true }           # 127.0.0.0/8
    if ($ip -match "^192\.168\.") { return $true }       # 192.168.0.0/16
    if ($ip -match "^10\.") { return $true }             # 10.0.0.0/8
    if ($ip -match "^172\.(1[6-9]|2[0-9]|3[01])\.") { 
        return $true 
    }                                                    # 172.16.0.0/12
    
    # IPv6 private ranges
    if ($ip -match "^fc[0-9a-f]{2}:") { return $true }   # fc00::/7
    if ($ip -match "^fd[0-9a-f]{2}:") { return $true }   # fd00::/8
    if ($ip -match "^fe80:") { return $true }            # fe80::/10 (Link-local)
    if ($ip -match "^::1") { return $true }              # Loopback
    
    # Alle anderen IPs sind öffentlich
    return $false
}

# Function to determine category and color based on port
function Get-CategoryAndColor {
    param([object]$LocalPort, [object]$RemotePort, [string]$Type)
    
    # Convert to int safely, handle "N/A" strings
    try {
        $localPortInt = if ($LocalPort -eq "N/A" -or $LocalPort -eq 0) { 0 } else { [int]$LocalPort }
        $remotePortInt = if ($RemotePort -eq "N/A" -or $RemotePort -eq 0) { 0 } else { [int]$RemotePort }
    } catch {
        $localPortInt = 0
        $remotePortInt = 0
    }
    
    $checkPort = if ($Type -eq "TCP" -and $remotePortInt -ne 0) { $remotePortInt } else { $localPortInt }
    
    foreach ($category in $protocolCategories.GetEnumerator()) {
        $categoryName = $category.Name
        $protocols = $category.Value
        
        foreach ($protocol in $protocols) {
            if ($protocolPorts.ContainsKey($protocol)) {
                if ($checkPort -in $protocolPorts[$protocol]) {
                    return @{
                        CategoryName = $categoryName
                        CategoryColor = $categoryColorMap[$categoryName]
                    }
                }
            }
        }
    }
    
    return @{
        CategoryName = "Sonstige"
        CategoryColor = "Other"
    }
}

# Function to get process name with caching
function Get-ProcessNameCached {
    param([int]$ProcessId)
    
    if ($Global:ProcessCache.ContainsKey($ProcessId)) {
        return $Global:ProcessCache[$ProcessId]
    }
    
    try {
        $process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
        $processName = if ($process) { $process.ProcessName } else { "Unknown" }
        $Global:ProcessCache[$ProcessId] = $processName
        return $processName
    } catch {
        $Global:ProcessCache[$ProcessId] = "Unknown"
        return "Unknown"
    }
}

# Function to get connections
function Get-NetworkConnections {
    param(
        [string]$protocol = "Alle", 
        [string]$direction = "Beide",
        [string]$category = "Alle Kategorien",
        [bool]$includeProcessInfo = $false,
        [string]$networkType = "Alle Netzwerke",
        [string]$tcpUdpType = "TCP + UDP"
    )

    # Get TCP connections
    $tcpConnections = Get-NetTCPConnection | Select-Object @{Name="Type";Expression={"TCP"}}, LocalAddress, LocalPort, RemoteAddress, RemotePort, State, OwningProcess, @{Name="ProcessName";Expression={""}}, @{Name="RemoteHostname";Expression={""}}, @{Name="CategoryName";Expression={""}}, @{Name="CategoryColor";Expression={""}}

    # Get UDP endpoints
    $udpConnections = Get-NetUDPEndpoint | Select-Object @{Name="Type";Expression={"UDP"}}, @{Name="LocalAddress";Expression={$_.LocalAddress}}, @{Name="LocalPort";Expression={$_.LocalPort}}, @{Name="RemoteAddress";Expression={"N/A"}}, @{Name="RemotePort";Expression={"N/A"}}, @{Name="State";Expression={"Listen"}}, @{Name="OwningProcess";Expression={$_.OwningProcess}}, @{Name="ProcessName";Expression={""}}, @{Name="RemoteHostname";Expression={"N/A"}}, @{Name="CategoryName";Expression={""}}, @{Name="CategoryColor";Expression={""}}

    # Filter by TCP/UDP type first
    if ($tcpUdpType -eq "Nur TCP") {
        $connections = $tcpConnections
    } elseif ($tcpUdpType -eq "Nur UDP") {
        $connections = $udpConnections
    } else {
        # "TCP + UDP"
        $connections = $tcpConnections + $udpConnections
    }

    # Filter by network type (Private/Public/All)
    if ($networkType -ne "Alle Netzwerke") {
        $connections = $connections | Where-Object {
            $isLocalPrivate = Test-IsPrivateIP -ip $_.LocalAddress
            $isRemotePrivate = Test-IsPrivateIP -ip $_.RemoteAddress
            $allPrivate = $isLocalPrivate -and $isRemotePrivate
            
            if ($networkType -eq "Nur Private") {
                $allPrivate
            } elseif ($networkType -eq "Nur Public") {
                -not $allPrivate
            } else {
                $true
            }
        }
    }

    # Filter by category first
    if ($category -ne "Alle Kategorien") {
        $categoryProtocols = $protocolCategories[$category]
        $categoryPorts = @()
        foreach ($prot in $categoryProtocols) {
            if ($protocolPorts.ContainsKey($prot)) {
                $categoryPorts += $protocolPorts[$prot]
            }
        }
        $connections = $connections | Where-Object { 
            ($_.LocalPort -in $categoryPorts) -or 
            ($_.RemotePort -in $categoryPorts -and $_.Type -eq "TCP")
        }
    }

    # Filter by specific protocol
    if ($protocol -ne "Alle") {
        $ports = $protocolPorts[$protocol]
        if ($direction -eq "Eingehend") {
            $connections = $connections | Where-Object { $_.LocalPort -in $ports }
        } elseif ($direction -eq "Ausgehend") {
            $connections = $connections | Where-Object { $_.RemotePort -in $ports -and $_.Type -eq "TCP" }
        } else {
            $connections = $connections | Where-Object { ($_.LocalPort -in $ports -or ($_.RemotePort -in $ports -and $_.Type -eq "TCP")) }
        }
    }

    # Resolve process names, hostnames and add category/color info
    foreach ($conn in $connections) {
        # Only resolve process names if requested
        if ($includeProcessInfo -and -not $conn.ProcessName) {
            $conn.ProcessName = Get-ProcessNameCached -ProcessId $conn.OwningProcess
        }
        
        # Add category and color info
        $categoryInfo = Get-CategoryAndColor -LocalPort $conn.LocalPort -RemotePort $conn.RemotePort -Type $conn.Type
        $conn.CategoryName = $categoryInfo.CategoryName
        $conn.CategoryColor = $categoryInfo.CategoryColor
        
        if ($conn.Type -eq "TCP" -and $conn.RemoteAddress -ne "N/A" -and -not $conn.RemoteHostname) {
            if ($conn.RemoteAddress -match "^127\.|^192\.168\.|^10\.|^172\.(1[6-9]|2[0-9]|3[01])\.") {
                $conn.RemoteHostname = Get-HostnameFast -ip $conn.RemoteAddress
            } else {
                $conn.RemoteHostname = $conn.RemoteAddress
            }
        }
    }

    return $connections
}

# Fast hostname resolution
function Get-HostnameFast {
    param([string]$ip)
    try {
        if ($ip -eq "0.0.0.0" -or $ip -eq "::" -or $ip -eq "127.0.0.1" -or $ip -eq "::1" -or $ip -eq "N/A") {
            return $ip
        }
        $hostEntry = [System.Net.Dns]::GetHostEntry($ip)
        return $hostEntry.HostName
    } catch {
        return $ip
    }
}

# Function to start Recording
function Start-ConnectionRecording {
    $Global:RecordingActive = $true
    $Global:RecordingStartTime = Get-Date
    $Global:RecordingCategory = $cmbCategory.SelectedItem.Content
    $Global:RecordedConnections = @()
    $Global:RecordedConnectionsHash = @{}
    
    $btnStartRecording.Content = "Aufzeichnung stoppen"
    $btnStartRecording.Background = [System.Windows.Media.Brushes]::Red
    
    $Global:RecordingTimer = New-Object System.Windows.Threading.DispatcherTimer
    $Global:RecordingTimer.Interval = [System.TimeSpan]::FromSeconds(3)
    
    $Global:RecordingTimer.Add_Tick({
        $selectedProtocol = $cmbProtocol.SelectedItem.Content
        $selectedDirection = $cmbDirection.SelectedItem.Content
        $selectedCategory = $cmbCategory.SelectedItem.Content
        $showProcessInfo = $chkShowProcessInfo.IsChecked
        $networkType = $cmbNetworkType.SelectedItem.Content
        $tcpUdpType = $cmbTCPUDP.SelectedItem.Content
        $connections = Get-NetworkConnections -protocol $selectedProtocol -direction $selectedDirection -category $selectedCategory -includeProcessInfo $showProcessInfo -networkType $networkType -tcpUdpType $tcpUdpType
        
        foreach ($conn in $connections) {
            $conn | Add-Member -MemberType NoteProperty -Name "RecordingTime" -Value (Get-Date) -Force
            
            $key = "$($conn.LocalAddress):$($conn.LocalPort)-$($conn.RemoteAddress):$($conn.RemotePort)-$($conn.Type)"
            
            if (-not $Global:RecordedConnectionsHash.ContainsKey($key)) {
                $Global:RecordedConnections += $conn
                $Global:RecordedConnectionsHash[$key] = $true
            }
        }
        
        $duration = ((Get-Date) - $Global:RecordingStartTime).TotalSeconds
        $txtStatus.Text = "Aufzeichnung läuft... Kategorie: $selectedCategory`nErfasste Verbindungen: $($Global:RecordedConnections.Count) | Dauer: $([Math]::Floor($duration))s | Letztes Update: $(Get-Date -Format 'HH:mm:ss')"
    })
    
    $Global:RecordingTimer.Start()
    $txtStatus.Text = "Aufzeichnung gestartet für Kategorie: $($cmbCategory.SelectedItem.Content)"
}

# Function to stop Recording
function Stop-ConnectionRecording {
    $Global:RecordingActive = $false
    $btnStartRecording.Content = "Aufzeichnung starten"
    $btnStartRecording.Background = [System.Windows.Media.Brushes]::Green
    
    if ($Global:RecordingTimer) {
        $Global:RecordingTimer.Stop()
        $Global:RecordingTimer = $null
    }
    
    $duration = ((Get-Date) - $Global:RecordingStartTime).TotalSeconds
    $txtStatus.Text = "Aufzeichnung beendet. Verbindungen erfasst: $($Global:RecordedConnections.Count) | Dauer: $([Math]::Floor($duration))s"
    
    if ($Global:RecordedConnections.Count -gt 0) {
        Export-RecordingReportHTML
    }
}

# Function to export Recording as HTML report
function Export-RecordingReportHTML {
    $DateTimeNow = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    $RecordingEnd = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    $RecordingStart = $Global:RecordingStartTime.ToString("dd.MM.yyyy HH:mm:ss")
    $duration = ((Get-Date) - $Global:RecordingStartTime).TotalSeconds
    $totalConnections = $Global:RecordedConnections.Count
    $tcpCount = ($Global:RecordedConnections | Where-Object { $_.Type -eq "TCP" }).Count
    $udpCount = ($Global:RecordedConnections | Where-Object { $_.Type -eq "UDP" }).Count
    $computerName = $env:COMPUTERNAME
    $userName = $env:USERNAME
    $category = $Global:RecordingCategory

    $html = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Aufzeichnungs-Report - easyConnections</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background-color: #f3f3f3; color: #333; }
        .container { max-width: 1400px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 0 20px rgba(0,0,0,0.1); }
        .header { text-align: center; margin-bottom: 30px; border-bottom: 2px solid #0078D4; padding-bottom: 20px; }
        h1 { color: #0078D4; margin: 0; font-size: 2.5em; }
        .subtitle { color: #666; font-size: 1.2em; margin: 10px 0; }
        .info-box { background-color: #e6f3ff; padding: 15px; border-radius: 6px; margin-bottom: 20px; border-left: 4px solid #0078D4; }
        .summary { display: flex; justify-content: space-around; margin-bottom: 20px; }
        .summary-item { text-align: center; background: #f8f9fa; padding: 10px; border-radius: 4px; }
        .timestamp { text-align: center; color: #666; margin-bottom: 20px; font-style: italic; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 14px; }
        th, td { padding: 10px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background-color: #0078D4; color: white; font-weight: bold; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        tr:hover { background-color: #e6f3ff; }
        .footer { text-align: center; margin-top: 30px; color: #666; border-top: 1px solid #ddd; padding-top: 20px; }
        .footer p { margin: 5px 0; }
        .footer a { color: #0078D4; text-decoration: none; }
        .footer a:hover { text-decoration: underline; }
        .recording-info { background-color: #fff3cd; padding: 10px; border-radius: 4px; margin-bottom: 20px; border-left: 4px solid #ffc107; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Netzwerk-Aufzeichnungs-Report</h1>
            <div class="subtitle">easyConnections Tool - Kontinuierliche Überwachung</div>
            <div class="info-box">
                <strong>Computer:</strong> $computerName | <strong>Benutzer:</strong> $userName<br>
                <strong>Kategorie:</strong> $category
            </div>
            <div class="recording-info">
                <strong>Aufzeichnungszeitraum:</strong> $RecordingStart bis $RecordingEnd<br>
                <strong>Dauer:</strong> $([Math]::Floor($duration)) Sekunden
            </div>
        </div>
        <div class="summary">
            <div class="summary-item">
                <strong>Gesamt erfasst</strong><br>$totalConnections
            </div>
            <div class="summary-item">
                <strong>TCP</strong><br>$tcpCount
            </div>
            <div class="summary-item">
                <strong>UDP</strong><br>$udpCount
            </div>
        </div>
        <div class="timestamp">Erstellt am $DateTimeNow</div>
        <table>
            <thead>
                <tr>
                    <th>Typ</th>
                    <th>Lokale Adresse</th>
                    <th>Lokaler Port</th>
                    <th>Remote Adresse</th>
                    <th>Remote Port</th>
                    <th>Remote Hostname</th>
                    <th>Status</th>
                    <th>Prozess ID</th>
                    <th>Prozess Name</th>
                    <th>Erfasst um</th>
                </tr>
            </thead>
            <tbody>
"@

    foreach ($conn in $Global:RecordedConnections) {
        $recordingTime = if ($conn.RecordingTime) { $conn.RecordingTime.ToString("dd.MM.yyyy HH:mm:ss") } else { "N/A" }
        $processName = if ($conn.ProcessName) { $conn.ProcessName } else { "N/A" }
        $html += "<tr><td>$($conn.Type)</td><td>$($conn.LocalAddress)</td><td>$($conn.LocalPort)</td><td>$($conn.RemoteAddress)</td><td>$($conn.RemotePort)</td><td>$($conn.RemoteHostname)</td><td>$($conn.State)</td><td>$($conn.OwningProcess)</td><td>$processName</td><td>$recordingTime</td></tr>"
    }

    $html += @"
            </tbody>
        </table>
    </div>
    <div class="footer">
        <p>Copyright © $(Get-Date -Format "yyyy") | Autor: Andreas Hepp | Webseite: <a href="https://www.phinit.de">www.phinit.de</a> / <a href="https://www.psscripts.de">www.psscripts.de</a></p>
        <p>Generiert mit easyConnections Tool V0.0.3 - Kontinuierliche Aufzeichnung</p>
    </div>
</body>
</html>
"@

    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "HTML files (*.html)|*.html"
    $saveFileDialog.Title = "Aufzeichnungs-Report speichern"
    $saveFileDialog.FileName = "RecordingReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"

    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $html | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8
        [System.Windows.MessageBox]::Show("Aufzeichnungs-Report gespeichert: $($saveFileDialog.FileName)", "Erfolg", "OK", "Information")
    }
}

# Function to export HTML report
function Export-HTMLReport {
    param(
        [Parameter(Mandatory=$true)]$connections,
        [Parameter(Mandatory=$true)][string]$protocol,
        [Parameter(Mandatory=$true)][string]$direction,
        [Parameter(Mandatory=$false)][string]$category = "Alle Kategorien"
    )

    $DateTimeNow = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    $totalConnections = $connections.Count
    $tcpCount = ($connections | Where-Object { $_.Type -eq "TCP" }).Count
    $udpCount = ($connections | Where-Object { $_.Type -eq "UDP" }).Count
    $computerName = $env:COMPUTERNAME
    $userName = $env:USERNAME

    $html = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Network Connections Report - easyConnections</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background-color: #f3f3f3; color: #333; }
        .container { max-width: 1400px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 0 20px rgba(0,0,0,0.1); }
        .header { text-align: center; margin-bottom: 30px; border-bottom: 2px solid #0078D4; padding-bottom: 20px; }
        h1 { color: #0078D4; margin: 0; font-size: 2.5em; }
        .subtitle { color: #666; font-size: 1.2em; margin: 10px 0; }
        .info-box { background-color: #e6f3ff; padding: 15px; border-radius: 6px; margin-bottom: 20px; border-left: 4px solid #0078D4; }
        .summary { display: flex; justify-content: space-around; margin-bottom: 20px; }
        .summary-item { text-align: center; background: #f8f9fa; padding: 10px; border-radius: 4px; }
        .timestamp { text-align: center; color: #666; margin-bottom: 20px; font-style: italic; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 14px; }
        th, td { padding: 10px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background-color: #0078D4; color: white; font-weight: bold; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        tr:hover { background-color: #e6f3ff; }
        .footer { text-align: center; margin-top: 30px; color: #666; border-top: 1px solid #ddd; padding-top: 20px; }
        .footer p { margin: 5px 0; }
        .footer a { color: #0078D4; text-decoration: none; }
        .footer a:hover { text-decoration: underline; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Network Connections Report</h1>
            <div class="subtitle">easyConnections Tool</div>
            <div class="info-box">
                <strong>Computer:</strong> $computerName | <strong>Benutzer:</strong> $userName<br>
                <strong>Filter:</strong> Kategorie: $category | Protokoll: $protocol | Richtung: $direction
            </div>
        </div>
        <div class="summary">
            <div class="summary-item">
                <strong>Gesamt</strong><br>$totalConnections
            </div>
            <div class="summary-item">
                <strong>TCP</strong><br>$tcpCount
            </div>
            <div class="summary-item">
                <strong>UDP</strong><br>$udpCount
            </div>
        </div>
        <div class="timestamp">Erstellt am $DateTimeNow</div>
        <table>
            <thead>
                <tr>
                    <th>Typ</th>
                    <th>Lokale Adresse</th>
                    <th>Lokaler Port</th>
                    <th>Remote Adresse</th>
                    <th>Remote Port</th>
                    <th>Remote Hostname</th>
                    <th>Status</th>
                    <th>Prozess ID</th>
                    <th>Prozess Name</th>
                </tr>
            </thead>
            <tbody>
"@

    foreach ($conn in $connections) {
        $processName = if ($conn.ProcessName) { $conn.ProcessName } else { "N/A" }
        $html += "<tr><td>$($conn.Type)</td><td>$($conn.LocalAddress)</td><td>$($conn.LocalPort)</td><td>$($conn.RemoteAddress)</td><td>$($conn.RemotePort)</td><td>$($conn.RemoteHostname)</td><td>$($conn.State)</td><td>$($conn.OwningProcess)</td><td>$processName</td></tr>"
    }

    $html += @"
            </tbody>
        </table>
    </div>
    <div class="footer">
        <p>Copyright © $(Get-Date -Format "yyyy") | Autor: Andreas Hepp | Webseite: <a href="https://www.phinit.de">www.phinit.de</a> / <a href="https://www.psscripts.de">www.psscripts.de</a></p>
        <p>Generiert mit easyConnections Tool V0.0.3</p>
    </div>
</body>
</html>
"@

    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "HTML files (*.html)|*.html"
    $saveFileDialog.Title = "HTML Report speichern"
    $saveFileDialog.FileName = "NetworkConnectionsReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"

    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $html | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8
        [System.Windows.MessageBox]::Show("HTML Report gespeichert: $($saveFileDialog.FileName)", "Erfolg", "OK", "Information")
    }
}

# Event handlers
$btnRefresh.Add_Click({
    try {
        $originalText = $btnRefresh.Content
        $btnRefresh.Content = "Lädt..."
        $btnRefresh.IsEnabled = $false
        
        $selectedProtocol = $cmbProtocol.SelectedItem.Content
        $selectedDirection = $cmbDirection.SelectedItem.Content
        $selectedCategory = $cmbCategory.SelectedItem.Content
        $showProcessInfo = $chkShowProcessInfo.IsChecked
        $networkType = $cmbNetworkType.SelectedItem.Content
        $tcpUdpType = $cmbTCPUDP.SelectedItem.Content
        
        $connections = Get-NetworkConnections -protocol $selectedProtocol -direction $selectedDirection -category $selectedCategory -includeProcessInfo $showProcessInfo -networkType $networkType -tcpUdpType $tcpUdpType
        $dgConnections.ItemsSource = $connections
        
        $txtStatus.Text = "Verbindungen geladen: $($connections.Count) gefunden | Protokoll: $selectedProtocol | Richtung: $selectedDirection | Kategorie: $selectedCategory | Netzwerk: $networkType | TCP/UDP: $tcpUdpType"
    } catch {
        [System.Windows.MessageBox]::Show("Fehler beim Laden der Verbindungen: $($_.Exception.Message)", "Fehler", "OK", "Error")
    } finally {
        $btnRefresh.Content = $originalText
        $btnRefresh.IsEnabled = $true
    }
})

# Checkbox event handler for process info visibility
$chkShowProcessInfo.Add_Click({
    $showProcessInfo = $chkShowProcessInfo.IsChecked
    Update-ProcessColumnsVisibility -showProcessInfo $showProcessInfo
})

# Recording button event
$btnStartRecording.Add_Click({
    if ($Global:RecordingActive) {
        Stop-ConnectionRecording
    } else {
        Start-ConnectionRecording
    }
})

# Hyperlink event handlers
$linkPhinit.Add_Click({
    try {
        Start-Process "https://www.phinit.de"
    } catch {
        [System.Diagnostics.Process]::Start("https://www.phinit.de")
    }
})

$linkPsscripts.Add_Click({
    try {
        Start-Process "https://www.psscripts.de"
    } catch {
        [System.Diagnostics.Process]::Start("https://www.psscripts.de")
    }
})

# Export HTML button event
$btnExportHTML.Add_Click({
    if ($dgConnections.ItemsSource.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Keine Verbindungen zum Exportieren. Bitte klicken Sie zuerst auf 'Aktualisieren'.", "Information", "OK", "Information")
        return
    }
    
    $selectedProtocol = $cmbProtocol.SelectedItem.Content
    $selectedDirection = $cmbDirection.SelectedItem.Content
    $selectedCategory = $cmbCategory.SelectedItem.Content
    $connections = $dgConnections.ItemsSource
    Export-HTMLReport -connections $connections -protocol $selectedProtocol -direction $selectedDirection -category $selectedCategory
})

# Window closing event
$window.add_Closing({
    if ($Global:RecordingActive) {
        Stop-ConnectionRecording
    }
})

# LAZY LOADING: Keine automatische Initialisierung beim Start - DataGrid bleibt leer
$dgConnections.ItemsSource = @()
$txtStatus.Text = "Bereit. Klicken Sie auf 'Aktualisieren' um Verbindungen zu laden oder starten Sie eine Aufzeichnung."

# Initialize process columns visibility (hidden by default)
Update-ProcessColumnsVisibility -showProcessInfo $false

# Show window
$window.ShowDialog() | Out-Null

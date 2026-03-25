#Requires -Version 5.1
<#
.SYNOPSIS
    Intune Device Lookup — search users and inspect their managed devices.
.DESCRIPTION
    WPF GUI that authenticates to Microsoft Graph, searches Entra ID users,
    lists their Intune-managed devices, and shows detailed device information.
.NOTES
    Requires the Microsoft.Graph.Authentication module.
#>

# ── Prerequisites ──
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    $answer = [System.Windows.MessageBox]::Show(
        "The 'Microsoft.Graph.Authentication' module is required but not installed.`nInstall it now?",
        "Missing Module", "YesNo", "Warning"
    )
    if ($answer -eq 'Yes') {
        try { Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber }
        catch {
            [System.Windows.MessageBox]::Show("Failed to install module: $_", "Error", "OK", "Error")
            return
        }
    } else { return }
}

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

# ═══════════════════════════════════════════════════════════
#  XAML
# ═══════════════════════════════════════════════════════════
[xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Intune Device Lookup"
    Height="740" Width="1100"
    MinHeight="600" MinWidth="920"
    WindowStartupLocation="CenterScreen"
    Background="#F0F2F5"
    FontFamily="Segoe UI">

    <Window.Resources>
        <!-- Primary button -->
        <Style x:Key="BtnPrimary" TargetType="Button">
            <Setter Property="Background" Value="#5B5FC7"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="18,8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Name="bd" Background="{TemplateBinding Background}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#4B4FC0"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="bd" Property="Background" Value="#B4B6D1"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Card border -->
        <Style x:Key="Card" TargetType="Border">
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="CornerRadius" Value="8"/>
            <Setter Property="BorderBrush" Value="#E5E7EB"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Effect">
                <Setter.Value>
                    <DropShadowEffect ShadowDepth="1" BlurRadius="6" Opacity="0.07" Color="Black"/>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Clean ListBox -->
        <Style x:Key="ListClean" TargetType="ListBox">
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Padding" Value="0"/>
            <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Disabled"/>
        </Style>

        <!-- ListBox item -->
        <Style x:Key="ItemClean" TargetType="ListBoxItem">
            <Setter Property="Padding" Value="12,10"/>
            <Setter Property="Margin" Value="2,1"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBoxItem">
                        <Border Name="Bd" Background="{TemplateBinding Background}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}"
                                Margin="{TemplateBinding Margin}">
                            <ContentPresenter/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#F3F4F6"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#EEF2FF"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- ═══ HEADER ═══ -->
        <Border Grid.Row="0" Background="#FFFFFF" BorderBrush="#E5E7EB" BorderThickness="0,0,0,1">
            <Border.Effect>
                <DropShadowEffect ShadowDepth="1" BlurRadius="4" Opacity="0.04" Color="Black"/>
            </Border.Effect>
            <Grid Margin="24,14">
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <Border Width="30" Height="30" CornerRadius="6" Background="#EEF2FF" Margin="0,0,12,0">
                        <TextBlock Text="&#x2B22;" FontSize="14" Foreground="#5B5FC7"
                                   HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <TextBlock Text="Intune Device Lookup" FontSize="17" FontWeight="SemiBold"
                               Foreground="#1A1B25" VerticalAlignment="Center"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
                    <Ellipse Name="StatusDot" Width="8" Height="8" Fill="#D1D5DB"
                             VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <TextBlock Name="StatusText" Text="Not connected" FontSize="12"
                               Foreground="#9CA3AF" VerticalAlignment="Center" Margin="0,0,16,0"/>
                    <Button Name="LoginButton" Content="Sign In" Style="{StaticResource BtnPrimary}"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- ═══ BODY ═══ -->
        <Grid Grid.Row="1" Margin="20,16,20,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="340"/>
                <ColumnDefinition Width="16"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- ─── LEFT PANEL ─── -->
            <Grid Grid.Column="0">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <!-- Search -->
                <Border Grid.Row="0" Style="{StaticResource Card}" Padding="16" Margin="0,0,0,12">
                    <StackPanel>
                        <TextBlock Text="Search User" FontSize="13" FontWeight="SemiBold"
                                   Foreground="#374151" Margin="0,0,0,8"/>
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <Border Grid.Column="0" CornerRadius="6" BorderBrush="#D1D5DB"
                                    BorderThickness="1" Background="#FAFBFC" Margin="0,0,8,0">
                                <TextBox Name="SearchBox" BorderThickness="0" Background="Transparent"
                                         FontSize="13" Padding="10,8" VerticalAlignment="Center"
                                         Foreground="#1A1B25" IsEnabled="False"/>
                            </Border>
                            <Button Grid.Column="1" Name="SearchButton" Content="Search"
                                    Style="{StaticResource BtnPrimary}" IsEnabled="False"/>
                        </Grid>
                        <TextBlock Name="SearchHint" Text="Sign in to start searching"
                                   FontSize="11" Foreground="#9CA3AF" Margin="2,6,0,0"/>
                    </StackPanel>
                </Border>

                <!-- User results -->
                <Border Grid.Row="1" Name="UserPanel" Style="{StaticResource Card}"
                        Padding="12" Margin="0,0,0,12" Visibility="Collapsed" MaxHeight="220">
                    <StackPanel>
                        <TextBlock Text="USERS" FontSize="11" FontWeight="SemiBold"
                                   Foreground="#9CA3AF" Margin="4,0,0,6"/>
                        <ListBox Name="UserList" Style="{StaticResource ListClean}"
                                 ItemContainerStyle="{StaticResource ItemClean}" MaxHeight="170">
                            <ListBox.ItemTemplate>
                                <DataTemplate>
                                    <StackPanel>
                                        <TextBlock Text="{Binding DisplayName}" FontSize="13"
                                                   FontWeight="Medium" Foreground="#1A1B25"/>
                                        <TextBlock Text="{Binding UPN}" FontSize="11" Foreground="#9CA3AF"/>
                                    </StackPanel>
                                </DataTemplate>
                            </ListBox.ItemTemplate>
                        </ListBox>
                    </StackPanel>
                </Border>

                <!-- Device results -->
                <Border Grid.Row="2" Name="DevicePanel" Style="{StaticResource Card}"
                        Padding="12" Visibility="Collapsed">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" Name="DeviceHeader" Text="DEVICES" FontSize="11"
                                   FontWeight="SemiBold" Foreground="#9CA3AF" Margin="4,0,0,6"/>
                        <ListBox Grid.Row="1" Name="DeviceList" Style="{StaticResource ListClean}"
                                 ItemContainerStyle="{StaticResource ItemClean}">
                            <ListBox.ItemTemplate>
                                <DataTemplate>
                                    <StackPanel>
                                        <TextBlock Text="{Binding DeviceName}" FontSize="13"
                                                   FontWeight="Medium" Foreground="#1A1B25"/>
                                        <TextBlock Text="{Binding OS}" FontSize="11" Foreground="#9CA3AF"/>
                                        <TextBlock Text="{Binding LastSync}" FontSize="10" Foreground="#B4B6D1"/>
                                    </StackPanel>
                                </DataTemplate>
                            </ListBox.ItemTemplate>
                        </ListBox>
                    </Grid>
                </Border>
            </Grid>

            <!-- ─── RIGHT PANEL ─── -->
            <Border Grid.Column="2" Style="{StaticResource Card}" Padding="32">
                <Grid>
                    <!-- Empty state -->
                    <StackPanel Name="EmptyState" VerticalAlignment="Center" HorizontalAlignment="Center">
                        <Border Width="56" Height="56" CornerRadius="28" Background="#F3F4F6"
                                HorizontalAlignment="Center" Margin="0,0,0,14">
                            <TextBlock Text="&#xE7F4;" FontFamily="Segoe MDL2 Assets" FontSize="22"
                                       Foreground="#B4B6D1" HorizontalAlignment="Center"
                                       VerticalAlignment="Center"/>
                        </Border>
                        <TextBlock Text="Select a device to view details" FontSize="14"
                                   Foreground="#9CA3AF" HorizontalAlignment="Center"/>
                    </StackPanel>

                    <!-- Loading state -->
                    <StackPanel Name="LoadingState" VerticalAlignment="Center"
                                HorizontalAlignment="Center" Visibility="Collapsed">
                        <ProgressBar IsIndeterminate="True" Height="3" Width="200"
                                     Foreground="#5B5FC7" Background="#E5E7EB" BorderThickness="0"
                                     Margin="0,0,0,12"/>
                        <TextBlock Text="Loading device details..." FontSize="13"
                                   Foreground="#5B5FC7" HorizontalAlignment="Center"/>
                    </StackPanel>

                    <!-- Detail view -->
                    <Grid Name="DetailPanel" Visibility="Collapsed">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>

                        <!-- Header -->
                        <StackPanel Grid.Row="0">
                            <TextBlock Name="DetailTitle" FontSize="22" FontWeight="SemiBold"
                                       Foreground="#1A1B25" Margin="0,0,0,4"/>
                            <TextBlock Name="DetailSubtitle" FontSize="12"
                                       Foreground="#9CA3AF" Margin="0,0,0,16"/>
                        </StackPanel>

                        <!-- Tab buttons -->
                        <Grid Grid.Row="1">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>

                            <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,0">
                                <Border Name="TabInfoBtn" Cursor="Hand" Padding="12,8"
                                        Background="Transparent" CornerRadius="6,6,0,0"
                                        BorderBrush="#5B5FC7" BorderThickness="0,0,0,2" Margin="0,0,4,0">
                                    <TextBlock Text="Details" FontSize="13" FontWeight="SemiBold"
                                               Foreground="#5B5FC7"/>
                                </Border>
                                <Border Name="TabAppsBtn" Cursor="Hand" Padding="12,8"
                                        Background="Transparent" CornerRadius="6,6,0,0"
                                        BorderBrush="Transparent" BorderThickness="0,0,0,2" Margin="0,0,4,0">
                                    <TextBlock Text="Applications" FontSize="13" FontWeight="Medium"
                                               Foreground="#9CA3AF"/>
                                </Border>
                                <Border Name="TabComplianceBtn" Cursor="Hand" Padding="12,8"
                                        Background="Transparent" CornerRadius="6,6,0,0"
                                        BorderBrush="Transparent" BorderThickness="0,0,0,2">
                                    <TextBlock Text="Compliance" FontSize="13" FontWeight="Medium"
                                               Foreground="#9CA3AF"/>
                                </Border>
                            </StackPanel>

                            <Border Grid.Row="1" Height="1" Background="#E5E7EB" Margin="0,0,0,16"/>

                            <!-- TAB: Details -->
                            <ScrollViewer Grid.Row="2" Name="TabInfoPanel"
                                          VerticalScrollBarVisibility="Auto">
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="190"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>

                                    <TextBlock Grid.Row="0" Grid.Column="0" Text="Device Name"
                                               Style="{x:Null}" FontSize="13" Foreground="#6B7280"
                                               FontWeight="Medium" Margin="0,0,0,18"/>
                                    <TextBlock Grid.Row="0" Grid.Column="1" Name="ValDeviceName"
                                               FontSize="13" Foreground="#1A1B25" Margin="0,0,0,18"/>

                                    <TextBlock Grid.Row="1" Grid.Column="0" Text="OS Version"
                                               FontSize="13" Foreground="#6B7280" FontWeight="Medium"
                                               Margin="0,0,0,18"/>
                                    <TextBlock Grid.Row="1" Grid.Column="1" Name="ValOSVersion"
                                               FontSize="13" Foreground="#1A1B25" Margin="0,0,0,18"/>

                                    <TextBlock Grid.Row="2" Grid.Column="0" Text="OS Build"
                                               FontSize="13" Foreground="#6B7280" FontWeight="Medium"
                                               Margin="0,0,0,18"/>
                                    <TextBlock Grid.Row="2" Grid.Column="1" Name="ValOSBuild"
                                               FontSize="13" Foreground="#1A1B25" Margin="0,0,0,18"/>

                                    <TextBlock Grid.Row="3" Grid.Column="0" Text="Last Logged On User"
                                               FontSize="13" Foreground="#6B7280" FontWeight="Medium"
                                               Margin="0,0,0,18"/>
                                    <TextBlock Grid.Row="3" Grid.Column="1" Name="ValLastLogon"
                                               FontSize="13" Foreground="#1A1B25" Margin="0,0,0,18"
                                               TextWrapping="Wrap"/>

                                    <TextBlock Grid.Row="4" Grid.Column="0" Text="Primary User"
                                               FontSize="13" Foreground="#6B7280" FontWeight="Medium"
                                               Margin="0,0,0,18"/>
                                    <TextBlock Grid.Row="4" Grid.Column="1" Name="ValPrimaryUser"
                                               FontSize="13" Foreground="#1A1B25" Margin="0,0,0,18"
                                               TextWrapping="Wrap"/>

                                    <TextBlock Grid.Row="5" Grid.Column="0" Text="All Logged-In Users"
                                               FontSize="13" Foreground="#6B7280" FontWeight="Medium"
                                               Margin="0,0,0,18" VerticalAlignment="Top"/>
                                    <TextBlock Grid.Row="5" Grid.Column="1" Name="ValAllUsers"
                                               FontSize="13" Foreground="#1A1B25" Margin="0,0,0,18"
                                               TextWrapping="Wrap"/>

                                    <TextBlock Grid.Row="6" Grid.Column="0" Text="Enrollment Profile"
                                               FontSize="13" Foreground="#6B7280" FontWeight="Medium"
                                               Margin="0,0,0,18"/>
                                    <TextBlock Grid.Row="6" Grid.Column="1" Name="ValEnrollProfile"
                                               FontSize="13" Foreground="#1A1B25" Margin="0,0,0,18"
                                               TextWrapping="Wrap"/>

                                    <TextBlock Grid.Row="7" Grid.Column="0" Text="Last Enrolled"
                                               FontSize="13" Foreground="#6B7280" FontWeight="Medium"
                                               Margin="0,0,0,18"/>
                                    <TextBlock Grid.Row="7" Grid.Column="1" Name="ValLastEnrolled"
                                               FontSize="13" Foreground="#1A1B25" Margin="0,0,0,18"/>

                                    <TextBlock Grid.Row="8" Grid.Column="0" Text="LAPS Password"
                                               FontSize="13" Foreground="#6B7280" FontWeight="Medium"
                                               Margin="0,0,0,18" VerticalAlignment="Center"/>
                                    <StackPanel Grid.Row="8" Grid.Column="1" Orientation="Horizontal"
                                                Margin="0,0,0,18">
                                        <TextBlock Name="ValLapsPassword" FontSize="13"
                                                   Foreground="#1A1B25" VerticalAlignment="Center"
                                                   FontFamily="Consolas" Text="N/A"/>
                                        <Button Name="LapsRevealBtn" Content="Reveal"
                                                Margin="10,0,0,0" Padding="8,2"
                                                FontSize="11" Cursor="Hand"
                                                Background="#EEF2FF" Foreground="#5B5FC7"
                                                BorderThickness="0" Visibility="Collapsed">
                                            <Button.Template>
                                                <ControlTemplate TargetType="Button">
                                                    <Border Background="{TemplateBinding Background}"
                                                            CornerRadius="4" Padding="{TemplateBinding Padding}">
                                                        <ContentPresenter HorizontalAlignment="Center"
                                                                          VerticalAlignment="Center"/>
                                                    </Border>
                                                </ControlTemplate>
                                            </Button.Template>
                                        </Button>
                                    </StackPanel>
                                </Grid>
                            </ScrollViewer>

                            <!-- TAB: Applications -->
                            <Grid Grid.Row="2" Name="TabAppsPanel" Visibility="Collapsed">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>
                                <Border Grid.Row="0" CornerRadius="6" BorderBrush="#D1D5DB"
                                        BorderThickness="1" Background="#FAFBFC" Margin="0,0,0,10">
                                    <TextBox Name="AppsSearchBox" BorderThickness="0"
                                             Background="Transparent" FontSize="13" Padding="10,8"
                                             VerticalAlignment="Center" Foreground="#1A1B25"/>
                                </Border>
                                <TextBlock Grid.Row="1" Name="AppsStatus" Text=""
                                           FontSize="12" Foreground="#9CA3AF" Margin="0,0,0,10"/>
                                <ListBox Grid.Row="2" Name="AppsList"
                                         Style="{StaticResource ListClean}"
                                         ItemContainerStyle="{StaticResource ItemClean}">
                                    <ListBox.ItemTemplate>
                                        <DataTemplate>
                                            <Grid>
                                                <Grid.ColumnDefinitions>
                                                    <ColumnDefinition Width="*"/>
                                                    <ColumnDefinition Width="Auto"/>
                                                </Grid.ColumnDefinitions>
                                                <StackPanel Grid.Column="0">
                                                    <TextBlock Text="{Binding Name}" FontSize="13"
                                                               FontWeight="Medium" Foreground="#1A1B25"/>
                                                    <TextBlock Text="{Binding Publisher}" FontSize="11"
                                                               Foreground="#9CA3AF"/>
                                                </StackPanel>
                                                <StackPanel Grid.Column="1" VerticalAlignment="Center"
                                                            Margin="12,0,0,0">
                                                    <TextBlock Text="{Binding Version}" FontSize="11"
                                                               Foreground="#6B7280" HorizontalAlignment="Right"/>
                                                    <TextBlock Text="{Binding InstallDate}" FontSize="10"
                                                               Foreground="#9CA3AF" HorizontalAlignment="Right"/>
                                                </StackPanel>
                                            </Grid>
                                        </DataTemplate>
                                    </ListBox.ItemTemplate>
                                </ListBox>
                            </Grid>

                            <!-- TAB: Compliance -->
                            <ScrollViewer Grid.Row="2" Name="TabCompliancePanel" Visibility="Collapsed"
                                          VerticalScrollBarVisibility="Auto">
                                <StackPanel>
                                    <TextBlock Name="ComplianceStatus" Text="" FontSize="12"
                                               Foreground="#9CA3AF" Margin="0,0,0,10"/>
                                    <StackPanel Name="ComplianceList"/>
                                </StackPanel>
                            </ScrollViewer>
                        </Grid>
                    </Grid>
                </Grid>
            </Border>
        </Grid>

        <!-- ═══ FOOTER ═══ -->
        <Border Grid.Row="2" Background="#FFFFFF" BorderBrush="#E5E7EB" BorderThickness="0,1,0,0"
                Padding="24,8">
            <TextBlock FontSize="11" Foreground="#9CA3AF" HorizontalAlignment="Center">
                <Run Text="Intune Device Lookup"/>
                <Run Text="·" Foreground="#D1D5DB"/>
                <Run Text="v1.0"/>
                <Run Text="·" Foreground="#D1D5DB"/>
                <Run Text="Jacob Düring Bakkeli"/>
                <Run Text="·" Foreground="#D1D5DB"/>
                <Run Text="2026"/>
            </TextBlock>
        </Border>

        <!-- ═══ OVERLAY (login progress) ═══ -->
        <Border Name="LoginOverlay" Grid.RowSpan="3" Background="#80000000" Visibility="Collapsed">
            <Border Background="#FFFFFF" CornerRadius="12" Padding="40"
                    HorizontalAlignment="Center" VerticalAlignment="Center" MinWidth="340">
                <Border.Effect>
                    <DropShadowEffect ShadowDepth="2" BlurRadius="20" Opacity="0.12" Color="Black"/>
                </Border.Effect>
                <StackPanel HorizontalAlignment="Center">
                    <TextBlock Name="OverlayText" Text="Connecting..."
                               FontSize="14" Foreground="#374151" HorizontalAlignment="Center"
                               TextWrapping="Wrap" TextAlignment="Center" MaxWidth="380"/>
                    <ProgressBar IsIndeterminate="True" Height="3" Margin="0,14,0,0"
                                 Foreground="#5B5FC7" Background="#E5E7EB"
                                 BorderThickness="0" Width="260"/>
                </StackPanel>
            </Border>
        </Border>
    </Grid>
</Window>
'@

# ═══════════════════════════════════════════════════════════
#  BUILD WINDOW
# ═══════════════════════════════════════════════════════════
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Grab named controls
$ui = @{}
@(
    'LoginButton','StatusDot','StatusText',
    'SearchBox','SearchButton','SearchHint',
    'UserPanel','UserList',
    'DevicePanel','DeviceHeader','DeviceList',
    'EmptyState','LoadingState','DetailPanel',
    'DetailTitle','DetailSubtitle',
    'ValDeviceName','ValOSVersion','ValOSBuild','ValLastLogon',
    'ValPrimaryUser','ValAllUsers','ValEnrollProfile','ValLastEnrolled','ValLapsPassword','LapsRevealBtn',
    'TabInfoBtn','TabAppsBtn','TabComplianceBtn',
    'TabInfoPanel','TabAppsPanel','TabCompliancePanel',
    'AppsList','AppsStatus','AppsSearchBox',
    'ComplianceList','ComplianceStatus',
    'LoginOverlay','OverlayText'
) | ForEach-Object { $ui[$_] = $window.FindName($_) }

# ═══════════════════════════════════════════════════════════
#  HELPERS
# ═══════════════════════════════════════════════════════════
$script:connected = $false

function Push-UI {
    # Force WPF to process pending render work
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [action]{}, [System.Windows.Threading.DispatcherPriority]::Background
    )
}

function Show-Overlay ([string]$Text) {
    $ui['OverlayText'].Text = $Text
    $ui['LoginOverlay'].Visibility = 'Visible'
    Push-UI
}

function Hide-Overlay { $ui['LoginOverlay'].Visibility = 'Collapsed' }

function Invoke-Graph {
    param([string]$Uri, [hashtable]$Headers = @{}, [int]$Retries = 2)
    for ($attempt = 1; $attempt -le $Retries; $attempt++) {
        try {
            return (Invoke-MgGraphRequest -Uri $Uri -Method GET -Headers $Headers -ErrorAction Stop)
        } catch {
            $msg = $_.Exception.Message
            # Retry on 500/503/timeout, but not on 4xx
            if ($attempt -lt $Retries -and ($msg -match '5\d{2}|timeout|ServiceUnavailable')) {
                Start-Sleep -Seconds (2 * $attempt)
                continue
            }
            [System.Windows.MessageBox]::Show(
                "Graph API error:`n$msg", "Error", "OK", "Error") | Out-Null
            return $null
        }
    }
}

function Reset-DetailPanel {
    $ui['DetailPanel'].Visibility  = 'Collapsed'
    $ui['LoadingState'].Visibility = 'Collapsed'
    $ui['EmptyState'].Visibility   = 'Visible'
}

# Track the currently selected device ID for lazy-loaded tabs
$script:currentDeviceId    = $null
$script:appsLoaded         = $false
$script:complianceLoaded   = $false

# ═══════════════════════════════════════════════════════════
#  TAB SWITCHING
# ═══════════════════════════════════════════════════════════
function Switch-Tab ([string]$Tab) {
    $brushConv = [System.Windows.Media.BrushConverter]::new()
    $active    = $brushConv.ConvertFrom('#5B5FC7')
    $passive   = $brushConv.ConvertFrom('#9CA3AF')
    $clear     = $brushConv.ConvertFrom('Transparent')

    # Deactivate all tabs
    foreach ($b in @('TabInfoBtn','TabAppsBtn','TabComplianceBtn')) {
        $ui[$b].BorderBrush             = $clear
        ($ui[$b].Child).Foreground      = $passive
        ($ui[$b].Child).FontWeight      = [System.Windows.FontWeights]::Medium
    }
    foreach ($p in @('TabInfoPanel','TabAppsPanel','TabCompliancePanel')) {
        $ui[$p].Visibility = 'Collapsed'
    }

    # Activate chosen tab
    $ui["Tab${Tab}Btn"].BorderBrush        = $active
    ($ui["Tab${Tab}Btn"].Child).Foreground = $active
    ($ui["Tab${Tab}Btn"].Child).FontWeight = [System.Windows.FontWeights]::SemiBold
    $ui["Tab${Tab}Panel"].Visibility       = 'Visible'

    # Lazy-load tab data
    if ($Tab -eq 'Apps' -and -not $script:appsLoaded -and $script:currentDeviceId) {
        Load-DeviceApps $script:currentDeviceId
    }
    if ($Tab -eq 'Compliance' -and -not $script:complianceLoaded -and $script:currentDeviceId) {
        Load-DeviceCompliance $script:currentDeviceId
    }
}

$ui['TabInfoBtn'].Add_MouseLeftButtonDown({ Switch-Tab 'Info' })
$ui['TabAppsBtn'].Add_MouseLeftButtonDown({ Switch-Tab 'Apps' })
$ui['TabComplianceBtn'].Add_MouseLeftButtonDown({ Switch-Tab 'Compliance' })

function Load-DeviceApps ([string]$deviceId) {
    $ui['AppsList'].Items.Clear()
    $ui['AppsStatus'].Text = 'Loading applications...'
    Push-UI

    # Filter out package-style app names (e.g. Microsoft.BingSearch, MicrosoftWindows.Client.WebExperience)
    # These use Publisher.AppName dot-notation. Keep only human-readable names like "Google Chrome".
    $packagePattern = '^\w+\.\w+\.'          # matches Vendor.Something.More (3+ dot segments)
    $packagePattern2 = '^Microsoft\.\w+'     # matches Microsoft.Anything
    $packagePattern3 = '^MicrosoftWindows\.'  # matches MicrosoftWindows.Anything
    $packagePattern4 = '^Windows\.'           # matches Windows.Anything
    $packagePattern5 = '^MSIX\\\\|^ms-resource:'  # MSIX paths or resource refs

    # Detected apps (apps actually found on the device)
    $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/${deviceId}/detectedApps?`$top=500"
    $apps = Invoke-Graph -Uri $uri
    if ($apps -and $apps.value -and $apps.value.Count -gt 0) {
        $filtered = $apps.value | Where-Object {
            $n = $_.displayName
            $n -and
            $n -notmatch $packagePattern -and
            $n -notmatch $packagePattern2 -and
            $n -notmatch $packagePattern3 -and
            $n -notmatch $packagePattern4 -and
            $n -notmatch $packagePattern5
        }
        $sorted = $filtered | Sort-Object { $_.displayName }
        $script:allFilteredApps = @()
        foreach ($app in $sorted) {
            $instDate = ''
            if ($app.lastUpdatedDateTime) {
                $instDate = ([datetime]$app.lastUpdatedDateTime).ToString('yyyy-MM-dd')
            }
            $obj = [PSCustomObject]@{
                Name        = $app.displayName
                Publisher   = if ($app.publisher) { $app.publisher } else { '' }
                Version     = if ($app.version) { $app.version } else { '' }
                InstallDate = $instDate
            }
            $script:allFilteredApps += $obj
            $ui['AppsList'].Items.Add($obj) | Out-Null
        }
        $count = $script:allFilteredApps.Count
        $totalHidden = $apps.value.Count - $count
        $hint = "$count application(s)"
        if ($totalHidden -gt 0) { $hint += "  ($totalHidden built-in hidden)" }
        $ui['AppsStatus'].Text = $hint
    } else {
        $ui['AppsStatus'].Text = 'No applications found on this device'
    }
    $script:appsLoaded = $true
}

# Store all filtered apps for search
$script:allFilteredApps = @()

function Filter-AppsList {
    $query = $ui['AppsSearchBox'].Text.Trim()
    $ui['AppsList'].Items.Clear()
    $matched = 0
    foreach ($app in $script:allFilteredApps) {
        if ([string]::IsNullOrEmpty($query) -or
            $app.Name -like "*$query*" -or
            $app.Publisher -like "*$query*") {
            $ui['AppsList'].Items.Add($app) | Out-Null
            $matched++
        }
    }
    $total = $script:allFilteredApps.Count
    if ([string]::IsNullOrEmpty($query)) {
        $ui['AppsStatus'].Text = "$total application(s)"
    } else {
        $ui['AppsStatus'].Text = "$matched of $total application(s) matching '$query'"
    }
}

$ui['AppsSearchBox'].Add_TextChanged({ Filter-AppsList })

# ═══════════════════════════════════════════════════════════
#  DEVICE COMPLIANCE
# ═══════════════════════════════════════════════════════════
# ── Friendly names for well-known compliance setting identifiers ──
$script:settingNameMap = @{
    # Windows – OS / build
    'osMinimumVersion'                                 = 'Minimum OS Version'
    'osMaximumVersion'                                 = 'Maximum OS Version'
    'mobileOsMinimumVersion'                           = 'Minimum Mobile OS Version'
    'mobileOsMaximumVersion'                           = 'Maximum Mobile OS Version'
    'osMinimumBuildVersion'                            = 'Minimum OS Build'
    'osMaximumBuildVersion'                            = 'Maximum OS Build'
    'validOperatingSystemBuildRanges'                  = 'Allowed OS Build Ranges'
    'earlyLaunchAntiMalwareDriverEnabled'              = 'Early-Launch Anti-Malware Driver'

    # Windows – security / encryption
    'bitLockerEnabled'                                 = 'BitLocker Encryption'
    'secureBootEnabled'                                = 'Secure Boot'
    'codeIntegrityEnabled'                             = 'Code Integrity'
    'storageRequireEncryption'                         = 'Storage Encryption Required'
    'systemIntegrityProtectionEnabled'                 = 'System Integrity Protection (SIP)'
    'deviceThreatProtectionEnabled'                    = 'Device Threat Protection'
    'deviceThreatProtectionRequiredSecurityLevel'      = 'Threat Protection Security Level'
    'advancedThreatProtectionRequiredSecurityLevel'    = 'Advanced Threat Protection Security Level'
    'windowsDefenderMalwareProtectionEnabled'          = 'Windows Defender (Malware Protection)'
    'antiSpywareRequired'                              = 'Anti-Spyware Required'
    'antivirusRequired'                                = 'Antivirus Required'
    'realTimeProtectionEnabled'                        = 'Real-Time Protection'
    'signatureOutOfDate'                               = 'Defender Signatures Up-to-Date'
    'defenderEnabled'                                  = 'Windows Defender Enabled'
    'defenderVersion'                                  = 'Windows Defender Version'
    'firewallEnabled'                                  = 'Firewall Enabled'
    'firewallBlockAllIncoming'                         = 'Firewall – Block All Incoming'
    'firewallEnableStealthMode'                        = 'Firewall – Stealth Mode'

    # Windows – password / PIN
    'passwordRequired'                                 = 'Password Required'
    'passwordBlockSimple'                              = 'Block Simple Passwords'
    'passwordRequiredToUnlockFromIdle'                 = 'Password Required to Unlock from Idle'
    'passwordMinutesOfInactivityBeforeLock'            = 'Idle Timeout Before Lock (minutes)'
    'passwordExpirationDays'                           = 'Password Expiration (days)'
    'passwordMinimumLength'                            = 'Minimum Password Length'
    'passwordMinimumCharacterSetCount'                 = 'Minimum Character Sets'
    'passwordRequiredType'                             = 'Required Password Type'
    'passwordPreviousPasswordBlockCount'               = 'Reuse of Previous Passwords Blocked'
    'pinRequired'                                      = 'PIN Required'

    # Windows – TPM / health
    'tpmRequired'                                      = 'TPM Required'
    'deviceHealthAttestationRequired'                  = 'Device Health Attestation Required'
    'requireHealthyDeviceReport'                       = 'Healthy Device Report Required'
    'activeFirewallRequired'                           = 'Active Firewall Required'
    'configurationManagerComplianceRequired'           = 'Config Manager Compliance Required'

    # Windows – user / account
    'userPrincipalName'                                = 'User Principal Name'
    'userExists'                                       = 'User Account Exists'

    # iOS / macOS – common
    'managedEmailProfileRequired'                      = 'Managed Email Profile Required'
    'jailBroken'                                       = 'Jailbreak Detected'
    'deviceThreatProtectionEnabled.iOS'                = 'Device Threat Protection (iOS)'
    'passcodeRequired'                                 = 'Passcode Required'
    'passcodeBlockSimple'                              = 'Block Simple Passcodes'
    'passcodeMinimumLength'                            = 'Minimum Passcode Length'
    'passcodeMinutesOfInactivityBeforeLock'            = 'Idle Timeout Before Lock (minutes)'
    'passcodeExpirationDays'                           = 'Passcode Expiration (days)'
    'passcodePreviousPasscodeBlockCount'               = 'Reuse of Previous Passcodes Blocked'
    'passcodeRequiredType'                             = 'Required Passcode Type'
    'securityBlockJailbrokenDevices'                   = 'Block Jailbroken Devices'
    'osVersionRange'                                   = 'OS Version Range'

    # Android
    'securityRequireGooglePlayServices'                = 'Google Play Services Required'
    'securityRequireUpToDateSecurityProviders'         = 'Up-to-Date Security Providers Required'
    'securityRequireCompanyPortalAppIntegrity'         = 'Company Portal App Integrity Required'
    'securityRequireSafetyNetAttestationBasicIntegrity'= 'SafetyNet Basic Integrity Required'
    'securityRequireSafetyNetAttestationCertifiedDevice'= 'SafetyNet Certified Device Required'
    'securityRequireVerifyApps'                        = 'Verify Apps Required'
    'securityDisableUsbDebugging'                      = 'USB Debugging Disabled'
    'securityPreventInstallAppsFromUnknownSources'     = 'Block Apps from Unknown Sources'
    'deviceThreatProtectionEnabled.android'            = 'Device Threat Protection (Android)'
    'passwordRequiredType.android'                     = 'Required Password Type (Android)'
    'workProfilePasswordRequired'                      = 'Work Profile Password Required'
    'workProfilePasswordMinimumLength'                 = 'Work Profile Minimum Password Length'
    'workProfilePasswordRequiredType'                  = 'Work Profile Required Password Type'
    'storageRequireEncryption.android'                 = 'Storage Encryption Required (Android)'
}

function Get-FriendlySettingName ([string]$raw) {
    # 1. Exact lookup (strip any leading namespace prefix like "DefaultDeviceCompliancePolicy.")
    $trimmed = $raw -replace '^.*\.(?=[A-Z])', ''   # remove "Prefix." if next char is uppercase
    if ($script:settingNameMap.ContainsKey($trimmed)) { return $script:settingNameMap[$trimmed] }
    if ($script:settingNameMap.ContainsKey($raw))     { return $script:settingNameMap[$raw] }

    # 2. Strip known OData class-path prefixes (e.g. "windows10CompliancePolicy.bitLockerEnabled")
    $leaf = $raw -replace '^[a-z]\w+Policy\.', '' `
                  -replace '^[a-z]\w+CompliancePolicy\.', '' `
                  -replace '^[a-z]\w+Configuration\.', ''
    if ($script:settingNameMap.ContainsKey($leaf)) { return $script:settingNameMap[$leaf] }

    # 3. camelCase / PascalCase → words (insert space before each upper-case run)
    $spaced = [regex]::Replace($leaf, '(?<=[a-z])(?=[A-Z])', ' ')
    # Capitalise first letter and return
    return ($spaced.Substring(0,1).ToUpper() + $spaced.Substring(1))
}

function Load-DeviceCompliance ([string]$deviceId) {
    $ui['ComplianceList'].Children.Clear()
    $ui['ComplianceStatus'].Text = 'Loading compliance policies...'
    Push-UI

    $bc = [System.Windows.Media.BrushConverter]::new()

    # Colour map per state
    $stateColour = @{
        'compliant'     = '#16A34A'
        'nonCompliant'  = '#DC2626'
        'error'         = '#EA580C'
        'inGracePeriod' = '#D97706'
        'unknown'       = '#9CA3AF'
        'notApplicable' = '#9CA3AF'
    }
    $defaultColour = '#9CA3AF'

    $iconMap = @{
        'compliant'     = [string][char]0x2713   # ✓
        'nonCompliant'  = [string][char]0x2717   # ✗
        'error'         = [string][char]0x2717   # ✗
        'inGracePeriod' = '~'
        'notApplicable' = [string][char]0x2014   # —
    }
    $dashIcon = [string][char]0x2014  # —

    $uri  = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/${deviceId}/deviceCompliancePolicyStates"
    $resp = Invoke-Graph -Uri $uri
    if (-not $resp -or -not $resp.value -or $resp.value.Count -eq 0) {
        $ui['ComplianceStatus'].Text = 'No compliance policies assigned to this device'
        $script:complianceLoaded = $true
        return
    }

    $uniquePolicies = $resp.value |
        Group-Object { $_.id } |
        ForEach-Object { $_.Group[0] }

    $compliantCount = 0
    $totalCount     = $uniquePolicies.Count

    foreach ($policy in ($uniquePolicies | Sort-Object { $_.displayName })) {
        $pState  = if ($policy.state) { $policy.state } else { 'unknown' }
        $pColour = if ($stateColour.ContainsKey($pState)) { $stateColour[$pState] } else { $defaultColour }
        $pIcon   = if ($iconMap.ContainsKey($pState)) { $iconMap[$pState] } else { $dashIcon }
        if ($pState -eq 'compliant') { $compliantCount++ }

        # ── Card border ──────────────────────────────────────────
        $card = [System.Windows.Controls.Border]::new()
        $card.Margin          = [System.Windows.Thickness]::new(0,0,0,10)
        $card.Background      = $bc.ConvertFrom('#FAFBFC')
        $card.BorderBrush     = $bc.ConvertFrom('#E5E7EB')
        $card.BorderThickness = [System.Windows.Thickness]::new(1)
        $card.CornerRadius    = [System.Windows.CornerRadius]::new(8)
        $card.Padding         = [System.Windows.Thickness]::new(14,12,14,12)

        $cardBody = [System.Windows.Controls.StackPanel]::new()

        # ── Policy header row ─────────────────────────────────────
        $headerRow = [System.Windows.Controls.StackPanel]::new()
        $headerRow.Orientation = [System.Windows.Controls.Orientation]::Horizontal
        $headerRow.Margin      = [System.Windows.Thickness]::new(0,0,0,8)

        $hIcon = [System.Windows.Controls.TextBlock]::new()
        $hIcon.Text              = $pIcon
        $hIcon.FontSize          = 14
        $hIcon.FontWeight        = [System.Windows.FontWeights]::Bold
        $hIcon.Foreground        = $bc.ConvertFrom($pColour)
        $hIcon.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        $hIcon.Margin            = [System.Windows.Thickness]::new(0,0,8,0)

        $hName = [System.Windows.Controls.TextBlock]::new()
        $hName.Text              = $policy.displayName
        $hName.FontSize          = 13
        $hName.FontWeight        = [System.Windows.FontWeights]::SemiBold
        $hName.Foreground        = $bc.ConvertFrom('#1A1B25')
        $hName.TextWrapping      = [System.Windows.TextWrapping]::Wrap
        $hName.VerticalAlignment = [System.Windows.VerticalAlignment]::Center

        $headerRow.Children.Add($hIcon) | Out-Null
        $headerRow.Children.Add($hName) | Out-Null
        $cardBody.Children.Add($headerRow) | Out-Null

        # ── Separator ─────────────────────────────────────────────
        $sep = [System.Windows.Controls.Border]::new()
        $sep.Height      = 1
        $sep.Background  = $bc.ConvertFrom('#E5E7EB')
        $sep.Margin      = [System.Windows.Thickness]::new(0,0,0,8)
        $cardBody.Children.Add($sep) | Out-Null

        # ── Settings ──────────────────────────────────────────────
        $settingsAdded = $false
        try {
            $sUri  = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/${deviceId}" +
                     "/deviceCompliancePolicyStates/$($policy.id)/settingStates"
            $sResp = Invoke-MgGraphRequest -Uri $sUri -Method GET -ErrorAction Stop

            if ($sResp -and $sResp.value -and $sResp.value.Count -gt 0) {
                $uniqueSettings = $sResp.value |
                    Group-Object { if ($_.settingName) { $_.settingName } else { $_.setting } } |
                    ForEach-Object { $_.Group[0] }
                foreach ($s in ($uniqueSettings | Sort-Object { $_.settingName })) {
                    $sState  = if ($s.state) { $s.state } else { 'unknown' }
                    $sColour = if ($stateColour.ContainsKey($sState)) { $stateColour[$sState] } else { $defaultColour }
                    $sIcon   = if ($iconMap.ContainsKey($sState)) { $iconMap[$sState] } else { $dashIcon }
                    $rawName = if ($s.settingName) { $s.settingName } else { $s.setting }
                    $friendlyName = Get-FriendlySettingName $rawName

                    # Row grid: [icon 18] [name *] [state badge Auto]
                    $row = [System.Windows.Controls.Grid]::new()
                    $row.Margin = [System.Windows.Thickness]::new(0,3,0,3)

                    $col0 = [System.Windows.Controls.ColumnDefinition]::new()
                    $col0.Width = [System.Windows.GridLength]::new(18)
                    $col1 = [System.Windows.Controls.ColumnDefinition]::new()
                    $col1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
                    $col2 = [System.Windows.Controls.ColumnDefinition]::new()
                    $col2.Width = [System.Windows.GridLength]::Auto
                    $row.ColumnDefinitions.Add($col0)
                    $row.ColumnDefinitions.Add($col1)
                    $row.ColumnDefinitions.Add($col2)

                    $rIcon = [System.Windows.Controls.TextBlock]::new()
                    $rIcon.Text              = $sIcon
                    $rIcon.FontSize          = 12
                    $rIcon.FontWeight        = [System.Windows.FontWeights]::SemiBold
                    $rIcon.Foreground        = $bc.ConvertFrom($sColour)
                    $rIcon.VerticalAlignment = [System.Windows.VerticalAlignment]::Top
                    $rIcon.Margin            = [System.Windows.Thickness]::new(0,1,0,0)
                    [System.Windows.Controls.Grid]::SetColumn($rIcon, 0)

                    $rName = [System.Windows.Controls.TextBlock]::new()
                    $rName.Text         = $friendlyName
                    $rName.FontSize     = 12
                    $rName.Foreground   = $bc.ConvertFrom('#374151')
                    $rName.TextWrapping = [System.Windows.TextWrapping]::Wrap
                    [System.Windows.Controls.Grid]::SetColumn($rName, 1)

                    $rState = [System.Windows.Controls.TextBlock]::new()
                    $rState.Text              = $sState
                    $rState.FontSize          = 11
                    $rState.Foreground        = $bc.ConvertFrom($sColour)
                    $rState.Margin            = [System.Windows.Thickness]::new(10,1,0,0)
                    $rState.VerticalAlignment = [System.Windows.VerticalAlignment]::Top
                    [System.Windows.Controls.Grid]::SetColumn($rState, 2)

                    $row.Children.Add($rIcon)  | Out-Null
                    $row.Children.Add($rName)  | Out-Null
                    $row.Children.Add($rState) | Out-Null
                    $cardBody.Children.Add($row) | Out-Null
                    $settingsAdded = $true
                }
            }
        } catch {
            $errNote = [System.Windows.Controls.TextBlock]::new()
            $errNote.Text       = "Could not load settings: $($_.Exception.Message)"
            $errNote.FontSize   = 11
            $errNote.Foreground = $bc.ConvertFrom('#EA580C')
            $errNote.TextWrapping = [System.Windows.TextWrapping]::Wrap
            $cardBody.Children.Add($errNote) | Out-Null
            $settingsAdded = $true
        }

        if (-not $settingsAdded) {
            $noSettings = [System.Windows.Controls.TextBlock]::new()
            $noSettings.Text       = 'No setting details available'
            $noSettings.FontSize   = 11
            $noSettings.Foreground = $bc.ConvertFrom('#9CA3AF')
            $cardBody.Children.Add($noSettings) | Out-Null
        }

        $card.Child = $cardBody
        $ui['ComplianceList'].Children.Add($card) | Out-Null
        Push-UI
    }

    $t = $totalCount
    $ui['ComplianceStatus'].Text = "$compliantCount of $t polic$(if ($t -eq 1) { 'y' } else { 'ies' }) compliant"
    $script:complianceLoaded = $true
}

# ═══════════════════════════════════════════════════════════
#  LOGIN / LOGOUT
# ═══════════════════════════════════════════════════════════
$ui['LoginButton'].Add_Click({
    if ($script:connected) {
        # ── Sign out ──
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        $script:connected = $false
        $ui['StatusDot'].Fill       = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#D1D5DB')
        $ui['StatusText'].Text      = 'Not connected'
        $ui['LoginButton'].Content  = 'Sign In'
        $ui['SearchBox'].IsEnabled  = $false
        $ui['SearchButton'].IsEnabled = $false
        $ui['SearchBox'].Text       = ''
        $ui['SearchHint'].Text      = 'Sign in to start searching'
        $ui['UserPanel'].Visibility   = 'Collapsed'
        $ui['DevicePanel'].Visibility = 'Collapsed'
        Reset-DetailPanel
        return
    }

    # ── Sign in ──
    Show-Overlay 'Opening browser for sign-in...'
    try {
        $ctx = Get-MgContext -ErrorAction SilentlyContinue
        if (-not $ctx) {
            # Run Connect-MgGraph in a background runspace so the browser
            # popup can complete without the WPF UI thread blocking it.
            $rs = [runspacefactory]::CreateRunspace()
            $rs.Open()
            $ps = [powershell]::Create().AddScript({
                Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
                Connect-MgGraph -Scopes @(
                    'DeviceManagementManagedDevices.Read.All',
                    'User.Read.All',
                    'DeviceManagementServiceConfig.Read.All'
                ) -NoWelcome -ErrorAction Stop
            })
            $ps.Runspace = $rs
            $handle = $ps.BeginInvoke()

            # Keep the WPF dispatcher alive while waiting
            while (-not $handle.IsCompleted) {
                Push-UI
                Start-Sleep -Milliseconds 100
            }

            # Collect result / errors
            $ps.EndInvoke($handle)
            if ($ps.HadErrors) {
                $errMsg = ($ps.Streams.Error | ForEach-Object { $_.ToString() }) -join "`n"
                throw $errMsg
            }
            $ps.Dispose()
            $rs.Close()
        }

        $script:connected = $true
        $ui['StatusDot'].Fill       = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#16A34A')

        # Fetch tenant display name
        $tenantName = 'Connected'
        try {
            $org = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/v1.0/organization?$select=displayName' -Method GET -ErrorAction Stop
            if ($org.value -and $org.value.Count -gt 0 -and $org.value[0].displayName) {
                $tenantName = $org.value[0].displayName
            }
        } catch { <# fall back to generic "Connected" #> }

        $ui['StatusText'].Text      = $tenantName
        $ui['LoginButton'].Content  = 'Sign Out'
        $ui['SearchBox'].IsEnabled  = $true
        $ui['SearchButton'].IsEnabled = $true
        $ui['SearchHint'].Text      = 'Search by name, email, or device name'
    } catch {
        [System.Windows.MessageBox]::Show(
            "Sign-in failed:`n$($_.Exception.Message)", "Authentication Error", "OK", "Error") | Out-Null
    } finally {
        Hide-Overlay
    }
})

# ═══════════════════════════════════════════════════════════
#  USER SEARCH
# ═══════════════════════════════════════════════════════════
function Start-UserSearch {
    $q = $ui['SearchBox'].Text.Trim()
    if ([string]::IsNullOrEmpty($q)) { return }

    $ui['SearchHint'].Text        = 'Searching...'
    $ui['UserPanel'].Visibility   = 'Collapsed'
    $ui['DevicePanel'].Visibility = 'Collapsed'
    Reset-DetailPanel
    Push-UI

    $safe = $q -replace "'", "''"
    $allUsers = @{}  # keyed by user id to deduplicate

    # ── 1. Search users by name / UPN / mail (partial match) ──
    $userUri = "https://graph.microsoft.com/v1.0/users?" +
        "`$filter=startswith(displayName,'$safe') or startswith(userPrincipalName,'$safe') or startswith(mail,'$safe')" +
        "&`$select=id,displayName,userPrincipalName,mail" +
        "&`$top=20&`$count=true"

    $userResult = Invoke-Graph -Uri $userUri -Headers @{ ConsistencyLevel = 'eventual' }
    if ($userResult -and $userResult.value) {
        foreach ($u in $userResult.value) {
            if (-not $allUsers.ContainsKey($u.id)) {
                $allUsers[$u.id] = [PSCustomObject]@{
                    DisplayName = $u.displayName
                    UPN         = $u.userPrincipalName
                    Id          = $u.id
                    Source      = 'user'
                }
            }
        }
    }

    # ── 2. Search devices by name, then resolve their users ──
    $devUri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?" +
        "`$filter=contains(deviceName,'$safe')" +
        "&`$select=id,deviceName,userPrincipalName,userDisplayName" +
        "&`$top=10"

    $devResult = Invoke-Graph -Uri $devUri
    if ($devResult -and $devResult.value) {
        foreach ($d in $devResult.value) {
            if ($d.userPrincipalName) {
                # Look up the full user to get their ID
                $safeUpn = $d.userPrincipalName -replace "'", "''"
                $uLookup = Invoke-Graph -Uri ("https://graph.microsoft.com/v1.0/users?" +
                    "`$filter=userPrincipalName eq '$safeUpn'" +
                    "&`$select=id,displayName,userPrincipalName")
                if ($uLookup -and $uLookup.value -and $uLookup.value.Count -gt 0) {
                    $u = $uLookup.value[0]
                    if (-not $allUsers.ContainsKey($u.id)) {
                        $allUsers[$u.id] = [PSCustomObject]@{
                            DisplayName = "$($u.displayName)  [$($d.deviceName)]"
                            UPN         = $u.userPrincipalName
                            Id          = $u.id
                            Source      = 'device'
                        }
                    }
                }
            }
        }
    }

    if ($allUsers.Count -eq 0) {
        $ui['SearchHint'].Text = 'No users or devices found'
        return
    }

    $ui['SearchHint'].Text = "$($allUsers.Count) result(s) found"
    $ui['UserList'].Items.Clear()

    foreach ($u in $allUsers.Values | Sort-Object DisplayName) {
        $ui['UserList'].Items.Add($u) | Out-Null
    }

    $ui['UserPanel'].Visibility = 'Visible'

    if ($allUsers.Count -eq 1) {
        $ui['UserList'].SelectedIndex = 0
    }
}

$ui['SearchButton'].Add_Click({ Start-UserSearch })
$ui['SearchBox'].Add_KeyDown({ if ($_.Key -eq 'Return') { Start-UserSearch } })

# ═══════════════════════════════════════════════════════════
#  USER SELECTED → LOAD DEVICES
# ═══════════════════════════════════════════════════════════
$ui['UserList'].Add_SelectionChanged({
    $sel = $ui['UserList'].SelectedItem
    if (-not $sel) { return }

    $ui['DeviceList'].Items.Clear()
    $ui['DevicePanel'].Visibility = 'Visible'
    $ui['DeviceHeader'].Text = "DEVICES"
    Reset-DetailPanel
    Push-UI

    $safeUpn = $sel.UPN -replace "'", "''"
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?" +
           "`$filter=userPrincipalName eq '$safeUpn'" +
           "&`$select=id,deviceName,operatingSystem,osVersion,serialNumber,model,manufacturer,lastSyncDateTime" +
           "&`$top=50"

    $result = Invoke-Graph -Uri $uri
    if (-not $result -or -not $result.value -or $result.value.Count -eq 0) {
        $ui['DeviceHeader'].Text = "No devices found"
        return
    }

    $ui['DeviceHeader'].Text = "$($result.value.Count) DEVICE(S)"

    $sorted = $result.value | Sort-Object {
        if ($_.lastSyncDateTime) { [datetime]$_.lastSyncDateTime } else { [datetime]::MinValue }
    } -Descending

    foreach ($d in $sorted) {
        $lastSync = if ($d.lastSyncDateTime) {
            'Last sync: ' + ([datetime]$d.lastSyncDateTime).ToString('yyyy-MM-dd HH:mm')
        } else { 'Last sync: unknown' }
        $ui['DeviceList'].Items.Add([PSCustomObject]@{
            DeviceName = $d.deviceName
            OS         = "$($d.operatingSystem) $($d.osVersion)"
            DeviceId   = $d.id
            Serial     = $d.serialNumber
            LastSync   = $lastSync
        }) | Out-Null
    }
})

# ═══════════════════════════════════════════════════════════
#  DEVICE SELECTED → LOAD DETAILS
# ═══════════════════════════════════════════════════════════
$ui['DeviceList'].Add_SelectionChanged({
    $sel = $ui['DeviceList'].SelectedItem
    if (-not $sel) { return }

    $ui['EmptyState'].Visibility   = 'Collapsed'
    $ui['DetailPanel'].Visibility  = 'Collapsed'
    $ui['LoadingState'].Visibility = 'Visible'
    Push-UI

    $id = $sel.DeviceId

    # ── Step 1: Core device via v1.0 (stable, no 500s) ──
    # Fetch without $select — avoids Bad Request from unsupported property names
    $dev = Invoke-Graph -Uri ("https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/${id}")

    if (-not $dev) {
        $ui['LoadingState'].Visibility = 'Collapsed'
        $ui['EmptyState'].Visibility   = 'Visible'
        return
    }

    # ── OS parsing ──
    $osBuild   = if ($dev.osVersion) { $dev.osVersion } else { 'N/A' }
    $osFriendly = $dev.operatingSystem
    if ($osBuild -match '^10\.0\.(\d+)') {
        $major = [int]$Matches[1]
        $osFriendly = if ($major -ge 22000) { 'Windows 11' } else { 'Windows 10' }
    }

    # ── Primary user (from device record) ──
    $primaryUser = 'N/A'
    if ($dev.userDisplayName -and $dev.userPrincipalName) {
        $primaryUser = "$($dev.userDisplayName) ($($dev.userPrincipalName))"
    } elseif ($dev.userPrincipalName) {
        $primaryUser = $dev.userPrincipalName
    }

    # ── Step 2: usersLoggedOn via beta (separate call, graceful failure) ──
    $lastLogon    = 'N/A'
    $allUsersText = 'N/A'

    try {
        $betaDev  = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices/${id}?`$select=usersLoggedOn" -Method GET -ErrorAction Stop
        $loggedOn = $betaDev.usersLoggedOn
    } catch {
        $loggedOn = $null
    }

    if ($loggedOn -and $loggedOn.Count -gt 0) {
        $sorted = $loggedOn | Sort-Object { [datetime]$_.lastLogOnDateTime } -Descending
        $resolved = @()

        foreach ($entry in $sorted) {
            $uid  = $entry.userId
            $time = ([datetime]$entry.lastLogOnDateTime).ToString('yyyy-MM-dd HH:mm')
            $name = $uid
            try {
                $uInfo = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/users/${uid}?`$select=displayName,userPrincipalName" -Method GET -ErrorAction SilentlyContinue
                if ($uInfo) { $name = "$($uInfo.displayName) ($($uInfo.userPrincipalName))" }
            } catch { <# SID or deleted user – keep raw id #> }
            $resolved += [PSCustomObject]@{ Name = $name; Time = $time }
        }

        if ($resolved.Count -gt 0) {
            $lastLogon    = "$($resolved[0].Name)`n$($resolved[0].Time)"
            $allUsersText = ($resolved | ForEach-Object { "$($_.Name) - $($_.Time)" }) -join "`n"
        }
    }

    # ── Enrolled date ──
    $enrolled = 'N/A'
    if ($dev.enrolledDateTime) {
        $enrolled = ([datetime]$dev.enrolledDateTime).ToString('yyyy-MM-dd HH:mm')
    }

    # ── Enrollment profile ──
    # enrollmentProfileName is a standard v1.0 property
    $enrollProfile = if ($dev.enrollmentProfileName) { $dev.enrollmentProfileName } else { 'N/A' }

    # ── Store device ID for lazy-loaded tabs ──
    $script:currentDeviceId    = $id
    $script:currentAadDeviceId = if ($dev.azureADDeviceId) { $dev.azureADDeviceId } else { $null }
    $script:lapsPassword       = $null
    $script:appsLoaded         = $false
    $script:complianceLoaded   = $false

    # ── Populate UI ──
    $ui['DetailTitle'].Text    = $dev.deviceName
    $ui['DetailSubtitle'].Text = "$($dev.manufacturer) $($dev.model) — S/N $($dev.serialNumber)"

    $ui['ValDeviceName'].Text   = $dev.deviceName
    $ui['ValOSVersion'].Text    = $osFriendly
    $ui['ValOSBuild'].Text      = $osBuild
    $ui['ValLastLogon'].Text     = $lastLogon
    $ui['ValPrimaryUser'].Text  = $primaryUser
    $ui['ValAllUsers'].Text      = $allUsersText
    $ui['ValEnrollProfile'].Text  = $enrollProfile
    $ui['ValLastEnrolled'].Text   = $enrolled

    # ── LAPS ──
    $ui['ValLapsPassword'].Text       = '••••••••  (click Reveal)'
    $ui['ValLapsPassword'].Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#9CA3AF')
    $ui['LapsRevealBtn'].Visibility   = 'Collapsed'
    $ui['LapsRevealBtn'].Content      = 'Reveal'

    if ($script:currentAadDeviceId) {
        try {
            $lapsResp = Invoke-MgGraphRequest `
                -Uri "https://graph.microsoft.com/beta/directory/deviceLocalCredentials/$($script:currentAadDeviceId)?`$select=credentials" `
                -Method GET -ErrorAction Stop
            if ($lapsResp.credentials -and $lapsResp.credentials.Count -gt 0) {
                # Pick the most recently backed-up credential
                $latest = $lapsResp.credentials | Sort-Object backupDateTime -Descending | Select-Object -First 1
                $script:lapsPassword = [System.Text.Encoding]::UTF8.GetString(
                    [System.Convert]::FromBase64String($latest.passwordBase64))
                $ui['LapsRevealBtn'].Visibility = 'Visible'
            } else {
                $ui['ValLapsPassword'].Text = 'Not configured'
            }
        } catch {
            $msg = $_.Exception.Message
            if ($msg -match '403|Forbidden|Authorization') {
                $ui['ValLapsPassword'].Text = 'No permission (DeviceLocalCredentials.Read.All required)'
            } else {
                $ui['ValLapsPassword'].Text = 'Not available'
            }
        }
    } else {
        $ui['ValLapsPassword'].Text = 'No Azure AD Device ID'
    }

    # Reset to Details tab
    Switch-Tab 'Info'
    $ui['AppsList'].Items.Clear()
    $ui['AppsSearchBox'].Text = ''
    $script:allFilteredApps = @()
    $ui['ComplianceList'].Children.Clear()
    $ui['ComplianceStatus'].Text = ''

    $ui['LoadingState'].Visibility = 'Collapsed'
    $ui['DetailPanel'].Visibility  = 'Visible'
})

# ═══════════════════════════════════════════════════════════
#  LAPS REVEAL / HIDE TOGGLE
# ═══════════════════════════════════════════════════════════
$ui['LapsRevealBtn'].Add_Click({
    $bc = [System.Windows.Media.BrushConverter]::new()
    if ($ui['LapsRevealBtn'].Content -eq 'Reveal') {
        $ui['ValLapsPassword'].Text       = $script:lapsPassword
        $ui['ValLapsPassword'].Foreground = $bc.ConvertFrom('#1A1B25')
        $ui['LapsRevealBtn'].Content      = 'Hide'
    } else {
        $ui['ValLapsPassword'].Text       = '••••••••  (click Reveal)'
        $ui['ValLapsPassword'].Foreground = $bc.ConvertFrom('#9CA3AF')
        $ui['LapsRevealBtn'].Content      = 'Reveal'
    }
})

# ═══════════════════════════════════════════════════════════
#  LAUNCH
# ═══════════════════════════════════════════════════════════
$window.ShowDialog() | Out-Null

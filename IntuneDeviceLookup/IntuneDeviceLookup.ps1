#Requires -Version 5.1
<#
.SYNOPSIS
    Intune Device Lookup — search users and inspect their managed devices.
.DESCRIPTION
    WPF GUI that authenticates to Microsoft Graph, searches Entra ID users,
    lists their Intune-managed devices, and shows detailed device information.
.NOTES
    Author:  Jacob Düring Bakkeli
    Version: 1.0.0
    Date:    2026-03-27
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
    Background="{DynamicResource ThBg}"
    FontFamily="Segoe UI">

    <Window.Resources>
        <!-- ── Theme brushes (light defaults; swapped at runtime for dark mode) ── -->
        <SolidColorBrush x:Key="ThBg"         Color="#F0F2F5"/>
        <SolidColorBrush x:Key="ThSurface"    Color="#FFFFFF"/>
        <SolidColorBrush x:Key="ThSurfaceAlt" Color="#FAFBFC"/>
        <SolidColorBrush x:Key="ThBorder"     Color="#E5E7EB"/>
        <SolidColorBrush x:Key="ThBorderSub"  Color="#D1D5DB"/>
        <SolidColorBrush x:Key="ThTxt1"       Color="#1A1B25"/>
        <SolidColorBrush x:Key="ThTxt2"       Color="#374151"/>
        <SolidColorBrush x:Key="ThTxt3"       Color="#9CA3AF"/>
        <SolidColorBrush x:Key="ThTxt4"       Color="#B4B6D1"/>
        <SolidColorBrush x:Key="ThTxt5"       Color="#6B7280"/>
        <SolidColorBrush x:Key="ThHover"      Color="#F3F4F6"/>
        <SolidColorBrush x:Key="ThSelected"   Color="#EEF2FF"/>
        <SolidColorBrush x:Key="ThInputBg"    Color="#FAFBFC"/>
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

        <!-- Secondary button -->
        <Style x:Key="BtnSecondary" TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource ThHover}"/>
            <Setter Property="Foreground" Value="{DynamicResource ThTxt2}"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="Medium"/>
            <Setter Property="Padding" Value="14,8"/>
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
                                <Setter TargetName="bd" Property="Background" Value="#E5E7EB"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Card border -->
        <Style x:Key="Card" TargetType="Border">
            <Setter Property="Background" Value="{DynamicResource ThSurface}"/>
            <Setter Property="CornerRadius" Value="8"/>
            <Setter Property="BorderBrush" Value="{DynamicResource ThBorder}"/>
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
        <Border Grid.Row="0" Background="{DynamicResource ThSurface}" BorderBrush="{DynamicResource ThBorder}" BorderThickness="0,0,0,1">
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
                               Foreground="{DynamicResource ThTxt1}" VerticalAlignment="Center"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
                    <Ellipse Name="StatusDot" Width="8" Height="8" Fill="#D1D5DB"
                             VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <TextBlock Name="StatusText" Text="Not connected" FontSize="12"
                               Foreground="{DynamicResource ThTxt3}" VerticalAlignment="Center" Margin="0,0,16,0"/>
                    <Button Name="CheckPermsButton" Content="Check Permissions"
                            Style="{StaticResource BtnSecondary}" Margin="0,0,8,0"/>
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
                                   Foreground="{DynamicResource ThTxt2}" Margin="0,0,0,8"/>
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <Border Grid.Column="0" CornerRadius="6" BorderBrush="{DynamicResource ThBorderSub}"
                                    BorderThickness="1" Background="{DynamicResource ThInputBg}" Margin="0,0,8,0">
                                <TextBox Name="SearchBox" BorderThickness="0" Background="Transparent"
                                         FontSize="13" Padding="10,8" VerticalAlignment="Center"
                                         Foreground="{DynamicResource ThTxt1}" IsEnabled="False"/>
                            </Border>
                            <Button Grid.Column="1" Name="SearchButton" Content="Search"
                                    Style="{StaticResource BtnPrimary}" IsEnabled="False"/>
                        </Grid>
                        <TextBlock Name="SearchHint" Text="Sign in to start searching"
                                   FontSize="11" Foreground="{DynamicResource ThTxt3}" Margin="2,6,0,0"/>
                    </StackPanel>
                </Border>

                <!-- User results -->
                <Border Grid.Row="1" Name="UserPanel" Style="{StaticResource Card}"
                        Padding="12" Margin="0,0,0,12" Visibility="Collapsed" MaxHeight="220">
                    <StackPanel>
                        <TextBlock Text="USERS" FontSize="11" FontWeight="SemiBold"
                                   Foreground="{DynamicResource ThTxt3}" Margin="4,0,0,6"/>
                        <ListBox Name="UserList" Style="{StaticResource ListClean}"
                                 ItemContainerStyle="{DynamicResource ItemClean}" MaxHeight="170">
                            <ListBox.ItemTemplate>
                                <DataTemplate>
                                    <StackPanel>
                                        <TextBlock Text="{Binding DisplayName}" FontSize="13"
                                                   FontWeight="Medium" Foreground="{DynamicResource ThTxt1}"/>
                                        <TextBlock Text="{Binding UPN}" FontSize="11" Foreground="{DynamicResource ThTxt3}"/>
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
                                   FontWeight="SemiBold" Foreground="{DynamicResource ThTxt3}" Margin="4,0,0,6"/>
                        <ListBox Grid.Row="1" Name="DeviceList" Style="{StaticResource ListClean}"
                                 ItemContainerStyle="{DynamicResource ItemClean}">
                            <ListBox.ItemTemplate>
                                <DataTemplate>
                                    <Grid>
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="Auto"/>
                                            <ColumnDefinition Width="*"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBlock Grid.Column="0" Text="&#x25CF;" FontSize="9"
                                                   Foreground="{Binding ComplianceColor}"
                                                   VerticalAlignment="Top" Margin="0,4,8,0"
                                                   ToolTip="{Binding ComplianceLabel}"/>
                                        <StackPanel Grid.Column="1">
                                            <TextBlock Text="{Binding DeviceName}" FontSize="13"
                                                       FontWeight="Medium" Foreground="{DynamicResource ThTxt1}"/>
                                            <TextBlock Text="{Binding OS}" FontSize="11" Foreground="{DynamicResource ThTxt3}"/>
                                            <TextBlock Text="{Binding LastSync}" FontSize="10" Foreground="{DynamicResource ThTxt4}"/>
                                        </StackPanel>
                                    </Grid>
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
                        <Border Width="56" Height="56" CornerRadius="28" Background="{DynamicResource ThHover}"
                                HorizontalAlignment="Center" Margin="0,0,0,14">
                            <TextBlock Text="&#xE7F4;" FontFamily="Segoe MDL2 Assets" FontSize="22"
                                       Foreground="{DynamicResource ThTxt4}" HorizontalAlignment="Center"
                                       VerticalAlignment="Center"/>
                        </Border>
                        <TextBlock Text="Select a device to view details" FontSize="14"
                                   Foreground="{DynamicResource ThTxt3}" HorizontalAlignment="Center"/>
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
                            <Grid Margin="0,0,0,4">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto" MaxWidth="340"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Name="DetailTitle" FontSize="22" FontWeight="SemiBold"
                                           Foreground="{DynamicResource ThTxt1}" VerticalAlignment="Center"
                                           TextTrimming="CharacterEllipsis"/>
                                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center"
                                            Margin="20,0,0,0">
                                    <!-- Device Actions group -->
                                    <Border CornerRadius="8" BorderBrush="{DynamicResource ThBorder}" BorderThickness="1"
                                            Background="{DynamicResource ThSurfaceAlt}" Padding="12,8" Margin="0,0,12,0">
                                        <StackPanel>
                                            <TextBlock Text="DEVICE ACTIONS" FontSize="10" FontWeight="Bold"
                                                       Foreground="{DynamicResource ThTxt3}" Margin="0,0,0,6"/>
                                            <StackPanel Orientation="Horizontal">
                                                <Border Name="RefreshBtn" Cursor="Hand"
                                                        CornerRadius="6" Padding="10,5" VerticalAlignment="Center" Margin="4,0,0,0"
                                                        Background="Transparent" BorderThickness="0">
                                                    <StackPanel Orientation="Horizontal">
                                                        <TextBlock Text="&#x1F504;" FontSize="11"
                                                                   Foreground="{DynamicResource ThTxt2}" Margin="0,0,5,0"
                                                                   VerticalAlignment="Center"/>
                                                        <TextBlock Text="Refresh" FontSize="11" FontWeight="Medium"
                                                                   Foreground="{DynamicResource ThTxt2}" VerticalAlignment="Center"/>
                                                    </StackPanel>
                                                </Border>
                                                <Border Name="SyncDeviceBtn" Cursor="Hand"
                                                        CornerRadius="6" Padding="10,5" VerticalAlignment="Center" Margin="4,0,0,0"
                                                        Background="Transparent" BorderThickness="0">
                                                    <StackPanel Orientation="Horizontal">
                                                        <TextBlock Text="&#x21BB;" FontSize="13" FontWeight="SemiBold"
                                                                   Foreground="{DynamicResource ThTxt2}" Margin="0,0,5,0"
                                                                   VerticalAlignment="Center"/>
                                                        <TextBlock Text="Sync" FontSize="11" FontWeight="Medium"
                                                                   Foreground="{DynamicResource ThTxt2}" VerticalAlignment="Center"/>
                                                    </StackPanel>
                                                </Border>
                                            </StackPanel>
                                        </StackPanel>
                                    </Border>
                                    <!-- Remote Actions group -->
                                    <Border CornerRadius="8" BorderBrush="{DynamicResource ThBorder}" BorderThickness="1"
                                            Background="{DynamicResource ThSurfaceAlt}" Padding="12,8">
                                        <StackPanel>
                                            <TextBlock Text="REMOTE ACTIONS" FontSize="10" FontWeight="Bold"
                                                       Foreground="{DynamicResource ThTxt3}" Margin="0,0,0,6"/>
                                            <StackPanel Orientation="Horizontal">
                                                <Border Name="FreshStartBtn" Cursor="Hand"
                                                        CornerRadius="6" Padding="10,5" VerticalAlignment="Center"
                                                        Background="Transparent" BorderThickness="0">
                                                    <StackPanel Orientation="Horizontal">
                                                        <TextBlock Text="&#x267B;" FontSize="11"
                                                                   Foreground="{DynamicResource ThTxt2}" Margin="0,0,5,0"
                                                                   VerticalAlignment="Center"/>
                                                        <TextBlock Text="Fresh Start" FontSize="11" FontWeight="Medium"
                                                                   Foreground="{DynamicResource ThTxt2}" VerticalAlignment="Center"/>
                                                    </StackPanel>
                                                </Border>
                                            </StackPanel>
                                        </StackPanel>
                                    </Border>
                                </StackPanel>
                                <Border Grid.Column="2" Name="ExportPdfBtn" Cursor="Hand"
                                        CornerRadius="6" Padding="10,5" VerticalAlignment="Center" Margin="6,0,0,0"
                                        Background="{DynamicResource ThHover}" BorderBrush="{DynamicResource ThBorder}"
                                        BorderThickness="1">
                                    <StackPanel Orientation="Horizontal">
                                        <TextBlock Text="&#x1F4C4;" FontSize="11"
                                                   Foreground="{DynamicResource ThTxt2}" Margin="0,0,5,0"
                                                   VerticalAlignment="Center"/>
                                        <TextBlock Text="Export PDF" FontSize="11" FontWeight="Medium"
                                                   Foreground="{DynamicResource ThTxt2}" VerticalAlignment="Center"/>
                                    </StackPanel>
                                </Border>
                            </Grid>
                            <TextBlock Name="DetailSubtitle" FontSize="12"
                                       Foreground="{DynamicResource ThTxt3}" Margin="0,0,0,16"/>
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
                                               Foreground="{DynamicResource ThTxt3}"/>
                                </Border>
                                <Border Name="TabComplianceBtn" Cursor="Hand" Padding="12,8"
                                        Background="Transparent" CornerRadius="6,6,0,0"
                                        BorderBrush="Transparent" BorderThickness="0,0,0,2" Margin="0,0,4,0">
                                    <TextBlock Text="Compliance" FontSize="13" FontWeight="Medium"
                                               Foreground="{DynamicResource ThTxt3}"/>
                                </Border>
                                <Border Name="TabSecurityBtn" Cursor="Hand" Padding="12,8"
                                        Background="Transparent" CornerRadius="6,6,0,0"
                                        BorderBrush="Transparent" BorderThickness="0,0,0,2" Margin="0,0,4,0">
                                    <TextBlock Text="Security" FontSize="13" FontWeight="Medium"
                                               Foreground="{DynamicResource ThTxt3}"/>
                                </Border>
                                <Border Name="TabGroupsBtn" Cursor="Hand" Padding="12,8"
                                        Background="Transparent" CornerRadius="6,6,0,0"
                                        BorderBrush="Transparent" BorderThickness="0,0,0,2" Margin="0,0,4,0">
                                    <TextBlock Text="Groups" FontSize="13" FontWeight="Medium"
                                               Foreground="{DynamicResource ThTxt3}"/>
                                </Border>

                            </StackPanel>

                            <Border Grid.Row="1" Height="1" Background="{DynamicResource ThBorder}" Margin="0,0,0,16"/>

                            <!-- TAB: Details -->
                            <ScrollViewer Grid.Row="2" Name="TabInfoPanel"
                                          VerticalScrollBarVisibility="Auto">
                                <StackPanel>
                                    <!-- ── Device Health Card ── -->
                                    <Border CornerRadius="8" BorderBrush="{DynamicResource ThBorder}" BorderThickness="1"
                                            Background="{DynamicResource ThSurfaceAlt}" Padding="16,14" Margin="0,0,0,20">
                                        <Grid>
                                            <Grid.ColumnDefinitions>
                                                <ColumnDefinition Width="*"/>
                                                <ColumnDefinition Width="*"/>
                                                <ColumnDefinition Width="Auto"/>
                                            </Grid.ColumnDefinitions>
                                            <!-- Left column: Compliance, Encryption, Last Sync -->
                                            <StackPanel Grid.Column="0" Margin="0,0,12,0">
                                                <TextBlock Text="DEVICE HEALTH" FontSize="10" FontWeight="SemiBold"
                                                           Foreground="{DynamicResource ThTxt3}" Margin="0,0,0,12"/>
                                                <!-- Compliance -->
                                                <Grid Margin="0,0,0,8">
                                                    <Grid.ColumnDefinitions>
                                                        <ColumnDefinition Width="14"/>
                                                        <ColumnDefinition Width="90"/>
                                                        <ColumnDefinition Width="*"/>
                                                    </Grid.ColumnDefinitions>
                                                    <TextBlock Name="HthCompliantIcon" Text="&#x25CF;" FontSize="9"
                                                               Foreground="#9CA3AF" VerticalAlignment="Center"/>
                                                    <TextBlock Grid.Column="1" Text="Compliance" FontSize="12"
                                                               Foreground="{DynamicResource ThTxt2}" VerticalAlignment="Center"/>
                                                    <TextBlock Grid.Column="2" Name="HthCompliantVal" Text="&#x2014;"
                                                               FontSize="12" Foreground="#9CA3AF" VerticalAlignment="Center"/>
                                                </Grid>
                                                <!-- Encryption -->
                                                <Grid Margin="0,0,0,8">
                                                    <Grid.ColumnDefinitions>
                                                        <ColumnDefinition Width="14"/>
                                                        <ColumnDefinition Width="90"/>
                                                        <ColumnDefinition Width="*"/>
                                                    </Grid.ColumnDefinitions>
                                                    <TextBlock Name="HthEncryptedIcon" Text="&#x25CF;" FontSize="9"
                                                               Foreground="#9CA3AF" VerticalAlignment="Center"/>
                                                    <TextBlock Grid.Column="1" Text="Encryption" FontSize="12"
                                                               Foreground="{DynamicResource ThTxt2}" VerticalAlignment="Center"/>
                                                    <TextBlock Grid.Column="2" Name="HthEncryptedVal" Text="&#x2014;"
                                                               FontSize="12" Foreground="#9CA3AF" VerticalAlignment="Center"/>
                                                </Grid>
                                                <!-- Last Sync -->
                                                <Grid>
                                                    <Grid.ColumnDefinitions>
                                                        <ColumnDefinition Width="14"/>
                                                        <ColumnDefinition Width="90"/>
                                                        <ColumnDefinition Width="*"/>
                                                    </Grid.ColumnDefinitions>
                                                    <TextBlock Name="HthSyncIcon" Text="&#x25CF;" FontSize="9"
                                                               Foreground="#9CA3AF" VerticalAlignment="Center"/>
                                                    <TextBlock Grid.Column="1" Text="Last Sync" FontSize="12"
                                                               Foreground="{DynamicResource ThTxt2}" VerticalAlignment="Center"/>
                                                    <TextBlock Grid.Column="2" Name="HthSyncVal" Text="&#x2014;"
                                                               FontSize="12" Foreground="#9CA3AF" VerticalAlignment="Center"/>
                                                </Grid>
                                            </StackPanel>
                                            <!-- Right column: Management, Defender/AV -->
                                            <StackPanel Grid.Column="1" Margin="12,0,0,0">
                                                <TextBlock Text="&#x200B;" FontSize="10" FontWeight="SemiBold"
                                                           Margin="0,0,0,12"/>
                                                <!-- Management State -->
                                                <Grid Margin="0,0,0,8">
                                                    <Grid.ColumnDefinitions>
                                                        <ColumnDefinition Width="14"/>
                                                        <ColumnDefinition Width="90"/>
                                                        <ColumnDefinition Width="*"/>
                                                    </Grid.ColumnDefinitions>
                                                    <TextBlock Name="HthMgmtIcon" Text="&#x25CF;" FontSize="9"
                                                               Foreground="#9CA3AF" VerticalAlignment="Center"/>
                                                    <TextBlock Grid.Column="1" Text="Management" FontSize="12"
                                                               Foreground="{DynamicResource ThTxt2}" VerticalAlignment="Center"/>
                                                    <TextBlock Grid.Column="2" Name="HthMgmtVal" Text="&#x2014;"
                                                               FontSize="12" Foreground="#9CA3AF" VerticalAlignment="Center"/>
                                                </Grid>
                                                <!-- Defender / AV -->
                                                <Grid>
                                                    <Grid.ColumnDefinitions>
                                                        <ColumnDefinition Width="14"/>
                                                        <ColumnDefinition Width="90"/>
                                                        <ColumnDefinition Width="*"/>
                                                    </Grid.ColumnDefinitions>
                                                    <TextBlock Name="HthDefenderIcon" Text="&#x25CF;" FontSize="9"
                                                               Foreground="#9CA3AF" VerticalAlignment="Center"/>
                                                    <TextBlock Grid.Column="1" Text="Defender / AV" FontSize="12"
                                                               Foreground="{DynamicResource ThTxt2}" VerticalAlignment="Center"/>
                                                    <TextBlock Grid.Column="2" Name="HthDefenderVal" Text="&#x2014;"
                                                               FontSize="12" Foreground="#9CA3AF" VerticalAlignment="Center"/>
                                                </Grid>
                                            </StackPanel>
                                            <!-- Score badge -->
                                            <Border Grid.Column="2" Name="HealthScoreBadge" CornerRadius="10"
                                                    MinWidth="80" Padding="16,12" VerticalAlignment="Center"
                                                    Margin="20,0,0,0" Background="#9CA3AF">
                                                <StackPanel HorizontalAlignment="Center">
                                                    <TextBlock Name="HealthScoreText" Text="&#x2014;" FontSize="22"
                                                               FontWeight="Bold" HorizontalAlignment="Center"
                                                               Foreground="White"/>
                                                    <TextBlock Name="HealthScoreLabel" Text="&#x2014;" FontSize="10"
                                                               FontWeight="SemiBold" HorizontalAlignment="Center"
                                                               Foreground="White" Opacity="0.9"/>
                                                </StackPanel>
                                            </Border>
                                        </Grid>
                                    </Border>
                                    <!-- ── Device Properties ── -->
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
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>

                                    <TextBlock Grid.Row="0" Grid.Column="0" Text="Device Name"
                                               Style="{x:Null}" FontSize="13" Foreground="{DynamicResource ThTxt5}"
                                               FontWeight="Medium" Margin="0,0,0,18"/>
                                    <TextBlock Grid.Row="0" Grid.Column="1" Name="ValDeviceName"
                                               FontSize="13" Foreground="{DynamicResource ThTxt1}" Margin="0,0,0,18"/>

                                    <TextBlock Grid.Row="1" Grid.Column="0" Text="OS Version"
                                               FontSize="13" Foreground="{DynamicResource ThTxt5}" FontWeight="Medium"
                                               Margin="0,0,0,18"/>
                                    <TextBlock Grid.Row="1" Grid.Column="1" Name="ValOSVersion"
                                               FontSize="13" Foreground="{DynamicResource ThTxt1}" Margin="0,0,0,18"/>

                                    <TextBlock Grid.Row="2" Grid.Column="0" Text="OS Build"
                                               FontSize="13" Foreground="{DynamicResource ThTxt5}" FontWeight="Medium"
                                               Margin="0,0,0,18"/>
                                    <TextBlock Grid.Row="2" Grid.Column="1" Name="ValOSBuild"
                                               FontSize="13" Foreground="{DynamicResource ThTxt1}" Margin="0,0,0,18"/>

                                    <TextBlock Grid.Row="3" Grid.Column="0" Text="Last Logged On User"
                                               FontSize="13" Foreground="{DynamicResource ThTxt5}" FontWeight="Medium"
                                               Margin="0,0,0,18"/>
                                    <TextBlock Grid.Row="3" Grid.Column="1" Name="ValLastLogon"
                                               FontSize="13" Foreground="{DynamicResource ThTxt1}" Margin="0,0,0,18"
                                               TextWrapping="Wrap"/>

                                    <TextBlock Grid.Row="4" Grid.Column="0" Text="Primary User"
                                               FontSize="13" Foreground="{DynamicResource ThTxt5}" FontWeight="Medium"
                                               Margin="0,0,0,18"/>
                                    <TextBlock Grid.Row="4" Grid.Column="1" Name="ValPrimaryUser"
                                               FontSize="13" Foreground="{DynamicResource ThTxt1}" Margin="0,0,0,18"
                                               TextWrapping="Wrap"/>

                                    <TextBlock Grid.Row="5" Grid.Column="0" Text="All Logged-In Users"
                                               FontSize="13" Foreground="{DynamicResource ThTxt5}" FontWeight="Medium"
                                               Margin="0,0,0,18" VerticalAlignment="Top"/>
                                    <TextBlock Grid.Row="5" Grid.Column="1" Name="ValAllUsers"
                                               FontSize="13" Foreground="{DynamicResource ThTxt1}" Margin="0,0,0,18"
                                               TextWrapping="Wrap"/>

                                    <TextBlock Grid.Row="6" Grid.Column="0" Text="Installed Enrollment Profile"
                                               FontSize="13" Foreground="{DynamicResource ThTxt5}" FontWeight="Medium"
                                               Margin="0,0,0,18"/>
                                    <TextBlock Grid.Row="6" Grid.Column="1" Name="ValEnrollProfile"
                                               FontSize="13" Foreground="{DynamicResource ThTxt1}" Margin="0,0,0,18"
                                               TextWrapping="Wrap"/>

                                    <TextBlock Grid.Row="7" Grid.Column="0" Text="Assigned Enrollment Profile"
                                               FontSize="13" Foreground="{DynamicResource ThTxt5}" FontWeight="Medium"
                                               Margin="0,0,0,18"/>
                                    <TextBlock Grid.Row="7" Grid.Column="1" Name="ValAssignedEnrollProfile"
                                               FontSize="13" Foreground="{DynamicResource ThTxt1}" Margin="0,0,0,18"
                                               TextWrapping="Wrap"/>

                                    <TextBlock Grid.Row="8" Grid.Column="0" Text="Autopilot Group Tag"
                                               FontSize="13" Foreground="{DynamicResource ThTxt5}" FontWeight="Medium"
                                               Margin="0,0,0,18" VerticalAlignment="Center"/>
                                    <StackPanel Grid.Row="8" Grid.Column="1" Orientation="Horizontal"
                                                Margin="0,0,0,18" VerticalAlignment="Center">
                                        <TextBlock Name="ValGroupTag"
                                                   FontSize="13" Foreground="{DynamicResource ThTxt1}"
                                                   TextWrapping="Wrap" VerticalAlignment="Center"/>
                                        <Border Name="GroupTagStatus" CornerRadius="4"
                                                Padding="6,2" Margin="8,0,0,0"
                                                VerticalAlignment="Center" Visibility="Collapsed">
                                            <TextBlock Name="GroupTagStatusText" FontSize="10"
                                                       FontWeight="SemiBold"/>
                                        </Border>
                                        <Border Name="ChangeGroupTagBtn" Cursor="Hand"
                                                CornerRadius="4" Padding="8,2" Margin="10,0,0,0"
                                                Background="#EEF2FF" VerticalAlignment="Center">
                                            <TextBlock Text="Change" FontSize="11" FontWeight="Medium"
                                                       Foreground="#5B5FC7"/>
                                        </Border>
                                    </StackPanel>

                                    <TextBlock Grid.Row="9" Grid.Column="0" Text="Last Enrolled"
                                               FontSize="13" Foreground="{DynamicResource ThTxt5}" FontWeight="Medium"
                                               Margin="0,0,0,18"/>
                                    <TextBlock Grid.Row="9" Grid.Column="1" Name="ValLastEnrolled"
                                               FontSize="13" Foreground="{DynamicResource ThTxt1}" Margin="0,0,0,18"/>

                                    <TextBlock Grid.Row="10" Grid.Column="0" Text="Last Password Change"
                                               FontSize="13" Foreground="{DynamicResource ThTxt5}" FontWeight="Medium"
                                               Margin="0,0,0,18"/>
                                    <TextBlock Grid.Row="10" Grid.Column="1" Name="ValLastPasswordChange"
                                               FontSize="13" Foreground="{DynamicResource ThTxt1}" Margin="0,0,0,18"/>

                                    <TextBlock Grid.Row="11" Grid.Column="0" Text="BIOS / UEFI Version"
                                               FontSize="13" Foreground="{DynamicResource ThTxt5}" FontWeight="Medium"
                                               Margin="0,0,0,18"/>
                                    <TextBlock Grid.Row="11" Grid.Column="1" Name="ValBiosVersion"
                                               FontSize="13" Foreground="{DynamicResource ThTxt1}" Margin="0,0,0,18"
                                               TextWrapping="Wrap"/>

                                    <TextBlock Grid.Row="12" Grid.Column="0" Text="Defender / AV"
                                               FontSize="13" Foreground="{DynamicResource ThTxt5}" FontWeight="Medium"
                                               Margin="0,0,0,18"/>
                                    <TextBlock Grid.Row="12" Grid.Column="1" Name="ValDefenderStatus"
                                               FontSize="13" Foreground="{DynamicResource ThTxt1}" Margin="0,0,0,18"
                                               TextWrapping="Wrap"/>

                                </Grid>
                                </StackPanel>
                            </ScrollViewer>

                            <!-- TAB: Applications -->
                            <Grid Grid.Row="2" Name="TabAppsPanel" Visibility="Collapsed">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>
                                <Border Grid.Row="0" CornerRadius="6" BorderBrush="{DynamicResource ThBorderSub}"
                                        BorderThickness="1" Background="{DynamicResource ThInputBg}" Margin="0,0,0,10">
                                    <TextBox Name="AppsSearchBox" BorderThickness="0"
                                             Background="Transparent" FontSize="13" Padding="10,8"
                                             VerticalAlignment="Center" Foreground="{DynamicResource ThTxt1}"/>
                                </Border>
                                <Grid Grid.Row="1" Margin="0,0,0,10">
                                    <TextBlock Name="AppsStatus" Text=""
                                               FontSize="12" Foreground="{DynamicResource ThTxt3}"
                                               VerticalAlignment="Center"/>
                                    <Border Name="AppsToggleHiddenBtn" HorizontalAlignment="Right"
                                            VerticalAlignment="Center" Cursor="Hand"
                                            CornerRadius="4" Padding="8,3"
                                            Background="{DynamicResource ThHover}"
                                            BorderBrush="{DynamicResource ThBorder}" BorderThickness="1"
                                            Visibility="Collapsed">
                                        <TextBlock Name="AppsToggleHiddenText" Text="Show hidden"
                                                   FontSize="11" FontWeight="Medium"
                                                   Foreground="{DynamicResource ThTxt2}"/>
                                    </Border>
                                </Grid>
                                <ListBox Grid.Row="2" Name="AppsList"
                                         Style="{StaticResource ListClean}"
                                         ItemContainerStyle="{DynamicResource ItemClean}">
                                    <ListBox.ItemTemplate>
                                        <DataTemplate>
                                            <Grid>
                                                <Grid.ColumnDefinitions>
                                                    <ColumnDefinition Width="*"/>
                                                    <ColumnDefinition Width="Auto"/>
                                                </Grid.ColumnDefinitions>
                                                <StackPanel Grid.Column="0">
                                                    <TextBlock Text="{Binding Name}" FontSize="13"
                                                               FontWeight="Medium" Foreground="{DynamicResource ThTxt1}"/>
                                                    <TextBlock Text="{Binding Publisher}" FontSize="11"
                                                               Foreground="{DynamicResource ThTxt3}"/>
                                                </StackPanel>
                                                <StackPanel Grid.Column="1" VerticalAlignment="Center"
                                                            Margin="12,0,0,0">
                                                    <TextBlock Text="{Binding Version}" FontSize="11"
                                                               Foreground="{DynamicResource ThTxt5}" HorizontalAlignment="Right"/>
                                                    <TextBlock Text="{Binding InstallDate}" FontSize="10"
                                                               Foreground="{DynamicResource ThTxt3}" HorizontalAlignment="Right"/>
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
                                    <!-- Compliance sub-tab bar -->
                                    <StackPanel Orientation="Horizontal" Margin="0,0,0,12">
                                        <Border Name="CompSubOverviewBtn" CornerRadius="12" Padding="10,4"
                                                Background="#3B82F6" Cursor="Hand">
                                            <TextBlock Text="Overview" FontSize="11" FontWeight="Medium" Foreground="White"/>
                                        </Border>
                                        <Border Name="CompSubPoliciesBtn" CornerRadius="12" Padding="10,4"
                                                Margin="6,0,0,0" Background="{DynamicResource ThBorder}" Cursor="Hand">
                                            <TextBlock Text="Policies" FontSize="11" FontWeight="Medium"
                                                       Foreground="{DynamicResource ThTxt2}"/>
                                        </Border>
                                    </StackPanel>
                                    <!-- Sub-panel: Overview -->
                                    <StackPanel Name="CompSubOverviewPanel">
                                        <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                                            <TextBlock Text="Device state:" FontSize="12"
                                                       Foreground="{DynamicResource ThTxt2}"
                                                       VerticalAlignment="Center" Margin="0,0,6,0"/>
                                            <Border Name="DeviceComplianceBadge" CornerRadius="4"
                                                    Padding="6,2" Background="#F3F4F6">
                                                <TextBlock Name="DeviceComplianceText" Text="—"
                                                           FontSize="11" FontWeight="Medium"
                                                           Foreground="#9CA3AF"/>
                                            </Border>
                                        </StackPanel>
                                        <TextBlock Name="ComplianceSyncInfo" Text="" FontSize="11"
                                                   Foreground="{DynamicResource ThTxt3}" Margin="0,0,0,4"/>
                                        <TextBlock Name="ComplianceGraceInfo" Text="" FontSize="11"
                                                   Foreground="#D97706" FontWeight="Medium"
                                                   Margin="0,0,0,8" Visibility="Collapsed"/>
                                        <Border Name="ComplianceMismatchCard" CornerRadius="8" Padding="12,10"
                                                Background="#FEF3C7" BorderBrush="#FDE68A" BorderThickness="1"
                                                Margin="0,4,0,12" Visibility="Collapsed">
                                            <StackPanel>
                                                <TextBlock Name="ComplianceMismatchTitle" Text=""
                                                           FontSize="12" FontWeight="SemiBold" Foreground="#92400E"
                                                           Margin="0,0,0,4"/>
                                                <TextBlock Name="ComplianceMismatchBody" Text=""
                                                           FontSize="11" Foreground="#78350F" TextWrapping="Wrap"/>
                                            </StackPanel>
                                        </Border>
                                        <TextBlock Name="ComplianceStatus" Text="" FontSize="12"
                                                   Foreground="{DynamicResource ThTxt3}" Margin="0,0,0,0"/>
                                    </StackPanel>
                                    <!-- Sub-panel: Policies -->
                                    <StackPanel Name="CompSubPoliciesPanel" Visibility="Collapsed">
                                        <StackPanel Name="ComplianceList"/>
                                    </StackPanel>
                                </StackPanel>
                            </ScrollViewer>

                            <!-- TAB: Security -->
                            <ScrollViewer Grid.Row="2" Name="TabSecurityPanel" Visibility="Collapsed"
                                          VerticalScrollBarVisibility="Auto">
                                <StackPanel Margin="0,0,0,16">
                                    <TextBlock Name="SecurityStatus" Text="" FontSize="12"
                                               Foreground="{DynamicResource ThTxt3}" Margin="0,0,0,12"/>

                                    <!-- LAPS card -->
                                    <Border Background="{DynamicResource ThSurfaceAlt}" BorderBrush="{DynamicResource ThBorder}" BorderThickness="1"
                                            CornerRadius="8" Padding="14,12,14,12" Margin="0,0,0,12">
                                        <StackPanel>
                                            <TextBlock Text="Local Administrator Password (LAPS)"
                                                       FontSize="13" FontWeight="SemiBold"
                                                       Foreground="{DynamicResource ThTxt1}" Margin="0,0,0,10"/>
                                            <Border Height="1" Background="{DynamicResource ThBorder}" Margin="0,0,0,10"/>
                                            <StackPanel Orientation="Horizontal">
                                                <TextBlock Name="ValLapsPassword" FontSize="13"
                                                           FontFamily="Consolas" Foreground="{DynamicResource ThTxt3}"
                                                           VerticalAlignment="Center" Text="N/A"/>
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
                                                <Button Name="LapsCopyBtn" Content="Copy"
                                                        Margin="6,0,0,0" Padding="8,2"
                                                        FontSize="11" Cursor="Hand"
                                                        Background="#F0FDF4" Foreground="#16A34A"
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
                                                <TextBlock Name="LapsCopyCheck" Text="✓" FontSize="13"
                                                           FontWeight="Bold" Foreground="#16A34A"
                                                           VerticalAlignment="Center" Margin="6,0,0,0"
                                                           Opacity="0"/>
                                            </StackPanel>
                                        </StackPanel>
                                    </Border>

                                    <!-- BitLocker card -->
                                    <Border Background="{DynamicResource ThSurfaceAlt}" BorderBrush="{DynamicResource ThBorder}" BorderThickness="1"
                                            CornerRadius="8" Padding="14,12,14,12">
                                        <StackPanel>
                                            <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                                                <TextBlock Text="BitLocker Encryption" FontSize="13"
                                                           FontWeight="SemiBold" Foreground="{DynamicResource ThTxt1}"
                                                           VerticalAlignment="Center"/>
                                                <Border Name="BitLockerStatusBadge" CornerRadius="4"
                                                        Padding="6,2" Margin="10,0,0,0"
                                                        Background="#F3F4F6">
                                                    <TextBlock Name="BitLockerStatusText" Text="Unknown"
                                                               FontSize="11" FontWeight="Medium"
                                                               Foreground="#9CA3AF"/>
                                                </Border>
                                            </StackPanel>
                                            <Border Height="1" Background="{DynamicResource ThBorder}" Margin="0,0,0,10"/>
                                            <TextBlock Name="BitLockerKeysStatus" Text=""
                                                       FontSize="12" Foreground="{DynamicResource ThTxt3}"
                                                       Margin="0,0,0,8"/>
                                            <StackPanel Name="BitLockerKeysList"/>
                                        </StackPanel>
                                    </Border>
                                </StackPanel>
                            </ScrollViewer>

                            <!-- TAB: Groups -->
                            <Grid Grid.Row="2" Name="TabGroupsPanel" Visibility="Collapsed">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>
                                <!-- Search bar -->
                                <Border Grid.Row="0" CornerRadius="6" BorderBrush="{DynamicResource ThBorderSub}"
                                        BorderThickness="1" Background="{DynamicResource ThInputBg}" Margin="0,0,0,10">
                                    <TextBox Name="GroupsSearchBox" BorderThickness="0"
                                             Background="Transparent" FontSize="13" Padding="10,8"
                                             VerticalAlignment="Center" Foreground="{DynamicResource ThTxt1}"/>
                                </Border>
                                <!-- Sub-tab pill bar -->
                                <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,12">
                                    <Border Name="GrpSubDeviceBtn" Cursor="Hand" CornerRadius="6"
                                            Padding="14,5" Margin="0,0,6,0">
                                        <TextBlock Text="Device" FontSize="12" FontWeight="SemiBold"/>
                                    </Border>
                                    <Border Name="GrpSubUserBtn" Cursor="Hand" CornerRadius="6"
                                            Padding="14,5">
                                        <TextBlock Text="User" FontSize="12" FontWeight="SemiBold"/>
                                    </Border>
                                </StackPanel>
                                <!-- Status row -->
                                <TextBlock Grid.Row="2" Name="GroupsActiveStatus" Text=""
                                           FontSize="12" Foreground="{DynamicResource ThTxt3}" Margin="0,0,0,10"/>
                                <!-- Scrollable list area -->
                                <ScrollViewer Grid.Row="3" VerticalScrollBarVisibility="Auto">
                                    <StackPanel>
                                        <StackPanel Name="GrpSubDevicePanel">
                                            <TextBlock Name="GroupsDeviceStatus" Text="" FontSize="12"
                                                       Foreground="{DynamicResource ThTxt3}" Margin="0,0,0,8"/>
                                            <StackPanel Name="GroupsDeviceList"/>
                                        </StackPanel>
                                        <StackPanel Name="GrpSubUserPanel" Visibility="Collapsed">
                                            <TextBlock Name="GroupsUserStatus" Text="" FontSize="12"
                                                       Foreground="{DynamicResource ThTxt3}" Margin="0,0,0,8"/>
                                            <StackPanel Name="GroupsUserList"/>
                                        </StackPanel>
                                    </StackPanel>
                                </ScrollViewer>
                            </Grid>


                        </Grid>
                    </Grid>
                </Grid>
            </Border>
        </Grid>

        <!-- ═══ FOOTER ═══ -->
        <Border Grid.Row="2" Background="{DynamicResource ThSurface}" BorderBrush="{DynamicResource ThBorder}" BorderThickness="0,1,0,0"
                Padding="24,8">
            <Grid>
                <TextBlock FontSize="11" Foreground="{DynamicResource ThTxt3}" HorizontalAlignment="Center"
                           VerticalAlignment="Center">
                    <Run Text="Intune Device Lookup"/>
                    <Run Text="·" Foreground="{DynamicResource ThBorderSub}"/>
                    <Run Text="v1.0"/>
                    <Run Text="·" Foreground="{DynamicResource ThBorderSub}"/>
                    <Run Text="Jacob Düring Bakkeli"/>
                    <Run Text="·" Foreground="{DynamicResource ThBorderSub}"/>
                    <Run Text="2026"/>
                </TextBlock>
                <Border Name="DarkModeToggle" HorizontalAlignment="Right" VerticalAlignment="Center"
                        Background="{DynamicResource ThHover}" CornerRadius="10" Padding="10,4"
                        Cursor="Hand">
                    <TextBlock Name="DarkModeToggleText" Text="🌙  Dark mode" FontSize="11"
                               Foreground="{DynamicResource ThTxt2}"/>
                </Border>
            </Grid>
        </Border>

        <!-- ═══ OVERLAY (login progress) ═══ -->
        <Border Name="LoginOverlay" Grid.RowSpan="3" Background="#80000000" Visibility="Collapsed">
            <Border Background="{DynamicResource ThSurface}" CornerRadius="12" Padding="40"
                    HorizontalAlignment="Center" VerticalAlignment="Center" MinWidth="340">
                <Border.Effect>
                    <DropShadowEffect ShadowDepth="2" BlurRadius="20" Opacity="0.12" Color="Black"/>
                </Border.Effect>
                <StackPanel HorizontalAlignment="Center">
                    <TextBlock Name="OverlayText" Text="Connecting..."
                               FontSize="14" Foreground="{DynamicResource ThTxt2}" HorizontalAlignment="Center"
                               TextWrapping="Wrap" TextAlignment="Center" MaxWidth="380"/>
                    <ProgressBar Name="OverlayBarIndeterminate" IsIndeterminate="True" Height="3" Margin="0,14,0,0"
                                 Foreground="#5B5FC7" Background="#E5E7EB"
                                 BorderThickness="0" Width="260"/>
                    <!-- Determinate progress bar for PDF export -->
                    <Grid Margin="0,14,0,0" Visibility="Collapsed" Name="PdfProgressPanel" Width="260">
                        <Border Background="#E5E7EB" CornerRadius="4" Height="8"/>
                        <Border Name="PdfProgressFill" Background="#5B5FC7" CornerRadius="4" Height="8"
                                HorizontalAlignment="Left" Width="0"/>
                    </Grid>
                    <TextBlock Name="PdfProgressPercent" Text="" FontSize="11" Foreground="{DynamicResource ThTxt3}"
                               HorizontalAlignment="Center" Margin="0,6,0,0" Visibility="Collapsed"/>
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
    'LoginButton','StatusDot','StatusText','CheckPermsButton','SyncDeviceBtn',
    'SearchBox','SearchButton','SearchHint',
    'UserPanel','UserList',
    'DevicePanel','DeviceHeader','DeviceList',
    'EmptyState','LoadingState','DetailPanel',
    'DetailTitle','DetailSubtitle','RefreshBtn','FreshStartBtn','ExportPdfBtn',
    'ValDeviceName','ValOSVersion','ValOSBuild','ValLastLogon',
    'ValPrimaryUser','ValAllUsers','ValEnrollProfile','ValAssignedEnrollProfile','ValGroupTag','GroupTagStatus','GroupTagStatusText','ChangeGroupTagBtn','ValLastEnrolled','ValLastPasswordChange','ValBiosVersion',
    'HealthScoreBadge','HealthScoreText','HealthScoreLabel',
    'HthCompliantIcon','HthCompliantVal',
    'HthEncryptedIcon','HthEncryptedVal',
    'HthSyncIcon','HthSyncVal',
    'HthMgmtIcon','HthMgmtVal',
    'HthDefenderIcon','HthDefenderVal',
    'ValDefenderStatus',
    'TabInfoBtn','TabAppsBtn','TabComplianceBtn','TabSecurityBtn','TabGroupsBtn',
    'TabInfoPanel','TabAppsPanel','TabCompliancePanel','TabSecurityPanel','TabGroupsPanel',
    'GroupsSearchBox','GroupsActiveStatus',
    'GrpSubDeviceBtn','GrpSubUserBtn','GrpSubDevicePanel','GrpSubUserPanel',
    'GroupsDeviceList','GroupsUserList','GroupsDeviceStatus','GroupsUserStatus',

    'AppsList','AppsStatus','AppsSearchBox','AppsToggleHiddenBtn','AppsToggleHiddenText',
    'ComplianceList','ComplianceStatus','DeviceComplianceBadge','DeviceComplianceText',
    'CompSubOverviewBtn','CompSubPoliciesBtn','CompSubOverviewPanel','CompSubPoliciesPanel',
    'ComplianceSyncInfo','ComplianceGraceInfo','ComplianceMismatchCard','ComplianceMismatchTitle','ComplianceMismatchBody',
    'ValLapsPassword','LapsRevealBtn','LapsCopyBtn','LapsCopyCheck',
    'BitLockerStatusBadge','BitLockerStatusText','BitLockerKeysStatus','BitLockerKeysList',
    'SecurityStatus',
    'DarkModeToggle','DarkModeToggleText',
    'LoginOverlay','OverlayText','OverlayBarIndeterminate',
    'PdfProgressPanel','PdfProgressFill','PdfProgressPercent'
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
    $ui['OverlayBarIndeterminate'].Visibility = 'Visible'
    $ui['PdfProgressPanel'].Visibility  = 'Collapsed'
    $ui['PdfProgressPercent'].Visibility = 'Collapsed'
    $ui['LoginOverlay'].Visibility = 'Visible'
    Push-UI
}

function Hide-Overlay { $ui['LoginOverlay'].Visibility = 'Collapsed' }

function Show-PdfProgress ([string]$StepText, [int]$Percent) {
    $ui['OverlayText'].Text = $StepText
    $ui['OverlayBarIndeterminate'].Visibility = 'Collapsed'
    $ui['PdfProgressPanel'].Visibility  = 'Visible'
    $ui['PdfProgressPercent'].Visibility = 'Visible'
    $ui['PdfProgressPercent'].Text = "$Percent%"
    $ui['PdfProgressFill'].Width = [math]::Round(260 * $Percent / 100)
    $ui['LoginOverlay'].Visibility = 'Visible'
    Push-UI
}

# Non-blocking sleep that keeps WPF responsive
function Sleep-UI ([int]$Ms) {
    $end = [datetime]::UtcNow.AddMilliseconds($Ms)
    while ([datetime]::UtcNow -lt $end) {
        Push-UI
        Start-Sleep -Milliseconds 30
    }
}

function Show-CopyConfirmation ([System.Windows.UIElement]$element) {
    # Fade-in then fade-out a checkmark element
    $element.Opacity = 1
    $fadeOut = [System.Windows.Media.Animation.DoubleAnimation]::new()
    $fadeOut.From       = 1.0
    $fadeOut.To         = 0.0
    $fadeOut.Duration   = [System.Windows.Duration]::new([timespan]::FromMilliseconds(800))
    $fadeOut.BeginTime  = [timespan]::FromMilliseconds(400)
    $element.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $fadeOut)
}

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

# ═══════════════════════════════════════════════════════════
#  DEVICE HEALTH CARD
# ═══════════════════════════════════════════════════════════
function Update-HealthCard {
    param($dev, $wps)

    $bc        = [System.Windows.Media.BrushConverter]::new()
    $colGreen  = '#16A34A'
    $colOrange = '#EA580C'
    $colRed    = '#DC2626'
    $colGray   = '#9CA3AF'
    $score     = 0

    # Signal 1: Compliance
    $compState = if ($dev.complianceState) { $dev.complianceState } else { 'unknown' }
    $ic = $colGray; $vt = 'Unknown'; $vc = $colGray
    switch ($compState) {
        'compliant'     { $ic = $colGreen;  $vt = 'Compliant';     $vc = $colGreen;  $score++ }
        'inGracePeriod' { $ic = $colOrange; $vt = 'Grace Period';  $vc = $colOrange }
        'noncompliant'  { $ic = $colRed;    $vt = 'Not Compliant'; $vc = $colRed    }
    }
    $ui['HthCompliantIcon'].Foreground = $bc.ConvertFrom($ic)
    $ui['HthCompliantVal'].Text        = $vt
    $ui['HthCompliantVal'].Foreground  = $bc.ConvertFrom($vc)

    # Signal 2: Encryption
    if ($dev.isEncrypted -eq $true) {
        $ic = $colGreen; $vt = 'Encrypted';     $vc = $colGreen; $score++
    } elseif ($dev.isEncrypted -eq $false) {
        $ic = $colRed;   $vt = 'Not Encrypted'; $vc = $colRed
    } else {
        $ic = $colGray;  $vt = 'N/A';           $vc = $colGray
    }
    $ui['HthEncryptedIcon'].Foreground = $bc.ConvertFrom($ic)
    $ui['HthEncryptedVal'].Text        = $vt
    $ui['HthEncryptedVal'].Foreground  = $bc.ConvertFrom($vc)

    # Signal 3: Last sync freshness
    if ($dev.lastSyncDateTime) {
        $syncDays = [math]::Floor(((Get-Date) - [datetime]$dev.lastSyncDateTime).TotalDays)
        $syncText = if ($syncDays -eq 0) { 'Today' } elseif ($syncDays -eq 1) { '1 day ago' } else { "$syncDays days ago" }
        if      ($syncDays -le 7)  { $ic = $colGreen;  $vc = $colGreen;  $score++ }
        elseif  ($syncDays -le 14) { $ic = $colOrange; $vc = $colOrange; $syncText += ' — sync recommended' }
        else                       { $ic = $colRed;    $vc = $colRed;    $syncText += ' — stale' }
    } else {
        $ic = $colRed; $vc = $colRed; $syncText = 'Never synced'
    }
    $ui['HthSyncIcon'].Foreground = $bc.ConvertFrom($ic)
    $ui['HthSyncVal'].Text        = $syncText
    $ui['HthSyncVal'].Foreground  = $bc.ConvertFrom($vc)

    # Signal 4: Management — show WHERE the device is managed
    $mgmt  = if ($dev.managementState) { $dev.managementState } else { 'unknown' }
    $agent = if ($dev.managementAgent) { $dev.managementAgent } else { '' }

    # Friendly name for the management agent/channel
    $agentLabel = switch ($agent) {
        'mdm'                                 { 'Intune (MDM)' }
        'configurationManagerClient'          { 'SCCM' }
        'configurationManagerClientMdm'       { 'Co-managed (Intune + SCCM)' }
        'configurationManagerClientMdmEas'    { 'Co-managed + EAS' }
        'eas'                                 { 'Exchange ActiveSync' }
        'easMdm'                              { 'Intune + EAS' }
        'easIntuneClient'                     { 'EAS (Intune client)' }
        'intuneClient'                        { 'Intune Client' }
        'jamf'                                { 'Jamf' }
        'googleCloudDevicePolicyController'   { 'Google (Android)' }
        'unknown'                             { 'Unknown' }
        default {
            if ($agent.Length -gt 0) { $agent.Substring(0,1).ToUpper() + $agent.Substring(1) } else { 'Unknown' }
        }
    }

    $actionPendingStates = @('retirePending','retireIssued','wipeRequested','wipePending','wipeIssued')
    $criticalStates      = @('wiped','unhealthy','deletePending')
    if      ($mgmt -eq 'managed')            { $ic = $colGreen;  $vt = $agentLabel;         $vc = $colGreen;  $score++ }
    elseif  ($mgmt -in $actionPendingStates) { $ic = $colOrange; $vt = "$agentLabel — Action Pending";  $vc = $colOrange }
    elseif  ($mgmt -in $criticalStates)      { $ic = $colRed;    $vt = "$agentLabel — Requires Action"; $vc = $colRed    }
    else {
        $ic = $colGray; $vt = $agentLabel; $vc = $colGray
    }
    $ui['HthMgmtIcon'].Foreground = $bc.ConvertFrom($ic)
    $ui['HthMgmtVal'].Text        = $vt
    $ui['HthMgmtVal'].Foreground  = $bc.ConvertFrom($vc)

    # Signal 5: Defender / AV (from windowsProtectionState via beta)
    if ($wps) {
        $rtpProp = $wps.realTimeProtectionEnabled
        $avProp  = $wps.antivirusEnabled
        $malProp = $wps.malwareProtectionEnabled
        $sigOld  = $wps.signatureUpdateOverdue -eq $true
        $hasData = ($null -ne $rtpProp) -or ($null -ne $avProp) -or ($null -ne $malProp)
        $protOn  = ($rtpProp -eq $true)
        $avOn    = ($avProp  -eq $true) -or ($malProp -eq $true)

        if (-not $hasData) {
            $ic = $colGray;   $vt = 'Not reporting';       $vc = $colGray
        } elseif ($protOn -and $avOn -and -not $sigOld) {
            $ic = $colGreen;  $vt = 'Protected';           $vc = $colGreen;  $score++
        } elseif ($sigOld) {
            $ic = $colOrange; $vt = 'Signatures outdated'; $vc = $colOrange
        } elseif (-not $protOn -and $null -ne $rtpProp) {
            $ic = $colRed;    $vt = 'RTP disabled';        $vc = $colRed
        } elseif (-not $avOn -and ($null -ne $avProp -or $null -ne $malProp)) {
            $ic = $colRed;    $vt = 'AV disabled';         $vc = $colRed
        } else {
            $ic = $colOrange; $vt = 'Needs attention';     $vc = $colOrange
        }
    } else {
        $ic = $colGray; $vt = 'N/A'; $vc = $colGray
    }
    $ui['HthDefenderIcon'].Foreground = $bc.ConvertFrom($ic)
    $ui['HthDefenderVal'].Text        = $vt
    $ui['HthDefenderVal'].Foreground  = $bc.ConvertFrom($vc)

    # Overall score badge (out of 5)
    $badgeCfg = switch ($score) {
        5 { @{ Bg='#16A34A'; Label='Healthy'  } }
        4 { @{ Bg='#2563EB'; Label='Good'     } }
        3 { @{ Bg='#D97706'; Label='Fair'     } }
        2 { @{ Bg='#DC2626'; Label='Poor'     } }
        1 { @{ Bg='#991B1B'; Label='Critical' } }
        0 { @{ Bg='#991B1B'; Label='Critical' } }
    }
    $ui['HealthScoreBadge'].Background = $bc.ConvertFrom($badgeCfg.Bg)
    $ui['HealthScoreText'].Text        = "$score/5"
    $ui['HealthScoreLabel'].Text       = $badgeCfg.Label
}

# ═══════════════════════════════════════════════════════════
#  DARK / LIGHT THEME TOGGLE
# ═══════════════════════════════════════════════════════════
function Set-Theme ([bool]$dark) {
    $mkBrush = { param($h) [System.Windows.Media.SolidColorBrush]::new(
        [System.Windows.Media.ColorConverter]::ConvertFromString($h)) }

    if ($dark) {
        $window.Resources['ThBg']         = (& $mkBrush '#111218')
        $window.Resources['ThSurface']    = (& $mkBrush '#1E1F2B')
        $window.Resources['ThSurfaceAlt'] = (& $mkBrush '#252631')
        $window.Resources['ThBorder']     = (& $mkBrush '#363748')
        $window.Resources['ThBorderSub']  = (& $mkBrush '#454659')
        $window.Resources['ThTxt1']       = (& $mkBrush '#E8E9F3')
        $window.Resources['ThTxt2']       = (& $mkBrush '#A8B3C8')
        $window.Resources['ThTxt3']       = (& $mkBrush '#6B7280')
        $window.Resources['ThTxt4']       = (& $mkBrush '#4A4B5E')
        $window.Resources['ThTxt5']       = (& $mkBrush '#9CA3AF')
        $window.Resources['ThHover']      = (& $mkBrush '#2D2E3B')
        $window.Resources['ThSelected']   = (& $mkBrush '#2A2C4A')
        $window.Resources['ThInputBg']    = (& $mkBrush '#252631')
        $ui['DarkModeToggleText'].Text     = '☀  Light mode'
    } else {
        $window.Resources['ThBg']         = (& $mkBrush '#F0F2F5')
        $window.Resources['ThSurface']    = (& $mkBrush '#FFFFFF')
        $window.Resources['ThSurfaceAlt'] = (& $mkBrush '#FAFBFC')
        $window.Resources['ThBorder']     = (& $mkBrush '#E5E7EB')
        $window.Resources['ThBorderSub']  = (& $mkBrush '#D1D5DB')
        $window.Resources['ThTxt1']       = (& $mkBrush '#1A1B25')
        $window.Resources['ThTxt2']       = (& $mkBrush '#374151')
        $window.Resources['ThTxt3']       = (& $mkBrush '#9CA3AF')
        $window.Resources['ThTxt4']       = (& $mkBrush '#B4B6D1')
        $window.Resources['ThTxt5']       = (& $mkBrush '#6B7280')
        $window.Resources['ThHover']      = (& $mkBrush '#F3F4F6')
        $window.Resources['ThSelected']   = (& $mkBrush '#EEF2FF')
        $window.Resources['ThInputBg']    = (& $mkBrush '#FAFBFC')
        $ui['DarkModeToggleText'].Text     = '🌙  Dark mode'
    }

    # Swap ItemClean style — ControlTemplate triggers can't use DynamicResource for named targets
    $hoverColor    = if ($dark) { '#2D2E3B' } else { '#F3F4F6' }
    $selectedColor = if ($dark) { '#2A2C4A' } else { '#EEF2FF' }
    $styleXaml = @"
<Style TargetType="ListBoxItem"
       xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
       xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
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
                        <Setter TargetName="Bd" Property="Background" Value="$hoverColor"/>
                    </Trigger>
                    <Trigger Property="IsSelected" Value="True">
                        <Setter TargetName="Bd" Property="Background" Value="$selectedColor"/>
                    </Trigger>
                </ControlTemplate.Triggers>
            </ControlTemplate>
        </Setter.Value>
    </Setter>
</Style>
"@
    $window.Resources['ItemClean'] = [Windows.Markup.XamlReader]::Parse($styleXaml)
}

# Track the currently selected device ID for lazy-loaded tabs
$script:currentDeviceId         = $null
$script:currentAadDeviceId      = $null
$script:currentPrimaryUpn       = $null
$script:currentAutopilotId      = $null
$script:deviceComplianceState   = $null
$script:deviceLastSync          = $null
$script:deviceGracePeriodExpiry = $null
$script:appsLoaded              = $false
$script:complianceLoaded        = $false
$script:securityLoaded          = $false
$script:groupsLoaded            = $false
$script:deviceIsEncrypted       = $null
$script:isDarkMode              = $false

# ═══════════════════════════════════════════════════════════
#  TAB SWITCHING
# ═══════════════════════════════════════════════════════════
function Switch-Tab ([string]$Tab) {
    $brushConv = [System.Windows.Media.BrushConverter]::new()
    $active    = $brushConv.ConvertFrom('#5B5FC7')
    $passive   = $brushConv.ConvertFrom('#9CA3AF')
    $clear     = $brushConv.ConvertFrom('Transparent')

    # Deactivate all tabs
    foreach ($b in @('TabInfoBtn','TabAppsBtn','TabComplianceBtn','TabSecurityBtn','TabGroupsBtn')) {
        $ui[$b].BorderBrush             = $clear
        ($ui[$b].Child).Foreground      = $passive
        ($ui[$b].Child).FontWeight      = [System.Windows.FontWeights]::Medium
    }
    foreach ($p in @('TabInfoPanel','TabAppsPanel','TabCompliancePanel','TabSecurityPanel','TabGroupsPanel')) {
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
    if ($Tab -eq 'Security' -and -not $script:securityLoaded -and $script:currentDeviceId) {
        Load-DeviceSecurityInfo $script:currentDeviceId
    }
    if ($Tab -eq 'Groups' -and -not $script:groupsLoaded -and $script:currentDeviceId) {
        Load-DeviceGroups $script:currentDeviceId
    }
}

$ui['TabInfoBtn'].Add_MouseLeftButtonDown({ Switch-Tab 'Info' })
$ui['TabAppsBtn'].Add_MouseLeftButtonDown({ Switch-Tab 'Apps' })
$ui['TabComplianceBtn'].Add_MouseLeftButtonDown({ Switch-Tab 'Compliance' })
$ui['TabSecurityBtn'].Add_MouseLeftButtonDown({ Switch-Tab 'Security' })
$ui['TabGroupsBtn'].Add_MouseLeftButtonDown({ Switch-Tab 'Groups' })

function Switch-ComplianceSubTab ([string]$subTab) {
    $bc         = [System.Windows.Media.BrushConverter]::new()
    $activeBg   = $bc.ConvertFrom('#3B82F6')
    $activeFg   = [System.Windows.Media.Brushes]::White
    $inactiveBg = $window.Resources['ThBorder']
    $inactiveFg = $window.Resources['ThTxt2']
    if ($subTab -eq 'Overview') {
        $ui['CompSubOverviewPanel'].Visibility = 'Visible'
        $ui['CompSubPoliciesPanel'].Visibility = 'Collapsed'
        $ui['CompSubOverviewBtn'].Background   = $activeBg
        ($ui['CompSubOverviewBtn'].Child).Foreground = $activeFg
        $ui['CompSubPoliciesBtn'].Background   = $inactiveBg
        ($ui['CompSubPoliciesBtn'].Child).Foreground = $inactiveFg
    } else {
        $ui['CompSubOverviewPanel'].Visibility = 'Collapsed'
        $ui['CompSubPoliciesPanel'].Visibility = 'Visible'
        $ui['CompSubPoliciesBtn'].Background   = $activeBg
        ($ui['CompSubPoliciesBtn'].Child).Foreground = $activeFg
        $ui['CompSubOverviewBtn'].Background   = $inactiveBg
        ($ui['CompSubOverviewBtn'].Child).Foreground = $inactiveFg
    }
}

$ui['CompSubOverviewBtn'].Add_MouseLeftButtonDown({ Switch-ComplianceSubTab 'Overview' })
$ui['CompSubPoliciesBtn'].Add_MouseLeftButtonDown({ Switch-ComplianceSubTab 'Policies' })

function Load-DeviceApps ([string]$deviceId) {
    $ui['AppsList'].Items.Clear()
    $ui['AppsStatus'].Text = 'Loading applications...'
    $ui['AppsToggleHiddenBtn'].Visibility = 'Collapsed'
    $script:showHiddenApps = $false
    $ui['AppsToggleHiddenText'].Text = 'Show hidden'
    Push-UI

    # Filter out package-style app names (e.g. Microsoft.BingSearch, MicrosoftWindows.Client.WebExperience)
    # These use Publisher.AppName dot-notation. Keep only human-readable names like "Google Chrome".
    $packagePattern = '^\w+\.\w+\.'          # matches Vendor.Something.More (3+ dot segments)
    $packagePattern2 = '^Microsoft\.\w+'     # matches Microsoft.Anything
    $packagePattern3 = '^MicrosoftWindows\.'  # matches MicrosoftWindows.Anything
    $packagePattern4 = '^Windows\.'           # matches Windows.Anything
    $packagePattern5 = '^MSIX\\|^ms-resource:'    # MSIX paths or resource refs

    # Detected apps (apps actually found on the device)
    $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/${deviceId}/detectedApps?`$top=500"
    $apps = Invoke-Graph -Uri $uri
    if ($apps -and $apps.value -and $apps.value.Count -gt 0) {
        $sorted = $apps.value | Sort-Object { $_.displayName }
        $script:allFilteredApps = @()
        $script:allHiddenApps   = @()
        foreach ($app in $sorted) {
            $n = $app.displayName
            $isHidden = (-not $n) -or
                        ($n -match $packagePattern) -or
                        ($n -match $packagePattern2) -or
                        ($n -match $packagePattern3) -or
                        ($n -match $packagePattern4) -or
                        ($n -match $packagePattern5)

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
            if ($isHidden) {
                $script:allHiddenApps += $obj
            } else {
                $script:allFilteredApps += $obj
                $ui['AppsList'].Items.Add($obj) | Out-Null
            }
        }
        $count = $script:allFilteredApps.Count
        $totalHidden = $script:allHiddenApps.Count
        $hint = "$count application(s)"
        if ($totalHidden -gt 0) {
            $hint += "  ($totalHidden built-in hidden)"
            $ui['AppsToggleHiddenBtn'].Visibility = 'Visible'
        }
        $ui['AppsStatus'].Text = $hint
    } else {
        $ui['AppsStatus'].Text = 'No applications found on this device'
    }
    $script:appsLoaded = $true
}

# Store all filtered apps for search
$script:allFilteredApps = @()
$script:allHiddenApps   = @()
$script:showHiddenApps  = $false

function Filter-AppsList {
    $query = $ui['AppsSearchBox'].Text.Trim()
    $ui['AppsList'].Items.Clear()
    $matched = 0
    $sourceApps = if ($script:showHiddenApps) {
        @($script:allFilteredApps) + @($script:allHiddenApps) | Sort-Object Name
    } else {
        $script:allFilteredApps
    }
    foreach ($app in $sourceApps) {
        if ([string]::IsNullOrEmpty($query) -or
            $app.Name -like "*$query*" -or
            $app.Publisher -like "*$query*") {
            $ui['AppsList'].Items.Add($app) | Out-Null
            $matched++
        }
    }
    $total = $sourceApps.Count
    if ([string]::IsNullOrEmpty($query)) {
        $ui['AppsStatus'].Text = "$total application(s)"
    } else {
        $ui['AppsStatus'].Text = "$matched of $total application(s) matching '$query'"
    }
}

$ui['AppsSearchBox'].Add_TextChanged({ Filter-AppsList })

$ui['AppsToggleHiddenBtn'].Add_MouseLeftButtonDown({
    $script:showHiddenApps = -not $script:showHiddenApps
    if ($script:showHiddenApps) {
        $ui['AppsToggleHiddenText'].Text = 'Hide hidden'
    } else {
        $ui['AppsToggleHiddenText'].Text = 'Show hidden'
    }
    Filter-AppsList
})

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

    # ── Device-level compliance state badge ──────────────────────────
    $devState = if ($script:deviceComplianceState) { $script:deviceComplianceState } else { 'unknown' }
    $stateMap = @{
        'compliant'     = @{ Label='Compliant';       Bg='#DCFCE7'; Fg='#16A34A' }
        'noncompliant'  = @{ Label='Non-Compliant';   Bg='#FEE2E2'; Fg='#DC2626' }
        'inGracePeriod' = @{ Label='Grace Period';    Bg='#FEF3C7'; Fg='#D97706' }
        'configManager' = @{ Label='Config Manager';  Bg='#EEF2FF'; Fg='#5B5FC7' }
    }
    $cfg = if ($stateMap.ContainsKey($devState)) { $stateMap[$devState] } else { @{ Label=$devState; Bg='#F3F4F6'; Fg='#9CA3AF' } }
    $ui['DeviceComplianceBadge'].Background = $bc.ConvertFrom($cfg.Bg)
    $ui['DeviceComplianceText'].Text        = $cfg.Label
    $ui['DeviceComplianceText'].Foreground  = $bc.ConvertFrom($cfg.Fg)

    # ── Sync info ─────────────────────────────────────────────────────
    $ui['ComplianceSyncInfo'].Text = if ($script:deviceLastSync) {
        'Last Intune sync: ' + $script:deviceLastSync.ToString('yyyy-MM-dd HH:mm')
    } else { '' }

    # ── Grace period info ─────────────────────────────────────────────
    if ($devState -eq 'inGracePeriod' -and $script:deviceGracePeriodExpiry) {
        $ui['ComplianceGraceInfo'].Text       = 'Grace period expires: ' + $script:deviceGracePeriodExpiry.ToString('yyyy-MM-dd HH:mm')
        $ui['ComplianceGraceInfo'].Visibility = 'Visible'
    } else {
        $ui['ComplianceGraceInfo'].Visibility = 'Collapsed'
    }

    # ── Reset mismatch card ───────────────────────────────────────────
    $ui['ComplianceMismatchCard'].Visibility = 'Collapsed'

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
        if ($devState -eq 'noncompliant') {
            $ui['ComplianceMismatchTitle'].Text = 'No compliance policies assigned'
            $ui['ComplianceMismatchBody'].Text  = "This device has no compliance policies assigned. Intune marks devices as non-compliant by default when no compliance policy has been applied — this is controlled by your tenant's Compliance Policy Settings in the Intune portal."
            $ui['ComplianceMismatchCard'].Visibility = 'Visible'
        }
        $script:complianceLoaded = $true
        return
    }

    $uniquePolicies = $resp.value |
        Group-Object { $_.id } |
        ForEach-Object { $_.Group[0] }

    # Map device OS to compatible platformType values
    $compatiblePlatforms = @('all')
    switch -Wildcard ($script:deviceOS) {
        'windows*' { $compatiblePlatforms += @('windows10AndLater','windows81AndLater','windowsPhone81') }
        'ios*'     { $compatiblePlatforms += @('iOS') }
        'ipados*'  { $compatiblePlatforms += @('iOS') }
        'macos*'   { $compatiblePlatforms += @('macOS') }
        'android*' { $compatiblePlatforms += @('android','androidForWork','androidWorkProfile') }
    }

    # Pre-fetch settings for each policy and determine if policy applies to this device
    $applicablePolicies = @()
    foreach ($policy in $uniquePolicies) {
        # Filter by platform: skip policies whose platformType doesn't match the device OS
        $pt = if ($policy.platformType) { $policy.platformType } else { '' }
        if ($pt -and $pt -notin $compatiblePlatforms) { continue }

        $policySettings = @()
        try {
            $sUri  = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/${deviceId}" +
                     "/deviceCompliancePolicyStates/$($policy.id)/settingStates"
            $sResp = Invoke-MgGraphRequest -Uri $sUri -Method GET -ErrorAction Stop
            if ($sResp -and $sResp.value) {
                $policySettings = @($sResp.value |
                    Group-Object { if ($_.settingName) { $_.settingName } else { $_.setting } } |
                    ForEach-Object { $_.Group[0] } |
                    Where-Object { $_.state -ne 'notApplicable' })
            }
        } catch { }
        $applicablePolicies += [PSCustomObject]@{ Policy = $policy; Settings = $policySettings }
    }

    if ($applicablePolicies.Count -eq 0) {
        $ui['ComplianceStatus'].Text = 'No compliance policies deployed to this device'
        $script:complianceLoaded = $true
        return
    }

    $compliantCount = 0
    $totalCount     = $applicablePolicies.Count

    foreach ($entry in ($applicablePolicies | Sort-Object { $_.Policy.displayName })) {
        $policy  = $entry.Policy
        $prefetchedSettings = $entry.Settings
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

        # ── Settings (already pre-fetched) ──────────────────────────────────
        $settingsAdded = $false
        if ($prefetchedSettings -and $prefetchedSettings.Count -gt 0) {
                foreach ($s in ($prefetchedSettings | Sort-Object { $_.settingName })) {
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

    # ── State / policy mismatch detection ────────────────────────────
    $allPoliciesCompliant = ($compliantCount -eq $totalCount -and $totalCount -gt 0)
    if ($devState -eq 'noncompliant' -and $allPoliciesCompliant) {
        $ui['ComplianceMismatchTitle'].Text = "⚠ State evaluation pending"
        $ui['ComplianceMismatchBody'].Text  = "Intune reports this device as Non-Compliant overall, but all $totalCount assigned compliance polic$(if ($totalCount -eq 1) { 'y shows' } else { 'ies show' }) compliant settings.`n`nCommon reasons:`n• The device recently became compliant but Intune hasn't re-evaluated the overall state yet`n• A new compliance policy was recently assigned and is still pending its first evaluation`n• The device hasn't synced with Intune since the last policy change`n`nTip: Trigger a sync from Settings → Accounts → Access Work or School → Sync, or initiate a remote sync from the Intune portal, to force a fresh compliance evaluation."
        $ui['ComplianceMismatchCard'].Visibility = 'Visible'
    } elseif ($devState -eq 'inGracePeriod' -and $allPoliciesCompliant) {
        $graceNote = if ($script:deviceGracePeriodExpiry) {
            " The grace period expires on $($script:deviceGracePeriodExpiry.ToString('yyyy-MM-dd'))."
        } else { '' }
        $ui['ComplianceMismatchTitle'].Text = "ℹ In grace period — policies now compliant"
        $ui['ComplianceMismatchBody'].Text  = "All compliance policies show compliant, but the device is still within a grace period from a previous non-compliant state.$graceNote Once the grace period ends and a sync occurs, the overall state will update to Compliant."
        $ui['ComplianceMismatchCard'].Visibility = 'Visible'
    } elseif ($devState -eq 'compliant' -and $compliantCount -lt $totalCount) {
        $nonCompliant = $totalCount - $compliantCount
        $ui['ComplianceMismatchTitle'].Text = "ℹ State update pending"
        $ui['ComplianceMismatchBody'].Text  = "The device is marked Compliant overall, but $nonCompliant polic$(if ($nonCompliant -eq 1) { 'y shows' } else { 'ies show' }) non-compliant states. This typically resolves automatically after the next compliance evaluation cycle or a device sync."
        $ui['ComplianceMismatchCard'].Visibility = 'Visible'
    }

    $script:complianceLoaded = $true
}

# ═══════════════════════════════════════════════════════════
#  SECURITY TAB (LAPS + BITLOCKER)
# ═══════════════════════════════════════════════════════════
function Load-DeviceSecurityInfo ([string]$deviceId) {
    $ui['SecurityStatus'].Text = ''
    $ui['BitLockerKeysList'].Children.Clear()
    $ui['BitLockerKeysStatus'].Text = 'Loading...'
    Push-UI

    $bc = [System.Windows.Media.BrushConverter]::new()

    # ── LAPS ─────────────────────────────────────────────────
    $ui['ValLapsPassword'].Text       = '••••••••  (click Reveal)'
    $ui['ValLapsPassword'].Foreground = $bc.ConvertFrom('#9CA3AF')
    $ui['LapsRevealBtn'].Visibility   = 'Collapsed'
    $ui['LapsRevealBtn'].Content      = 'Reveal'
    $ui['LapsCopyBtn'].Visibility     = 'Collapsed'
    $script:lapsPassword = $null

    if ($script:currentAadDeviceId) {
        try {
            $lapsResp = Invoke-MgGraphRequest `
                -Uri "https://graph.microsoft.com/beta/directory/deviceLocalCredentials/$($script:currentAadDeviceId)?`$select=credentials" `
                -Method GET -ErrorAction Stop
            if ($lapsResp.credentials -and $lapsResp.credentials.Count -gt 0) {
                $latest = $lapsResp.credentials | Sort-Object backupDateTime -Descending | Select-Object -First 1
                $script:lapsPassword = [System.Text.Encoding]::UTF8.GetString(
                    [System.Convert]::FromBase64String($latest.passwordBase64))
                $ui['LapsRevealBtn'].Visibility = 'Visible'
                $ui['LapsCopyBtn'].Visibility   = 'Visible'
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

    # ── BitLocker encryption status (from device record) ─────
    if ($script:deviceIsEncrypted -eq $true) {
        $ui['BitLockerStatusText'].Text        = 'Encrypted'
        $ui['BitLockerStatusBadge'].Background = $bc.ConvertFrom('#DCFCE7')
        $ui['BitLockerStatusText'].Foreground  = $bc.ConvertFrom('#16A34A')
    } elseif ($script:deviceIsEncrypted -eq $false) {
        $ui['BitLockerStatusText'].Text        = 'Not Encrypted'
        $ui['BitLockerStatusBadge'].Background = $bc.ConvertFrom('#FEE2E2')
        $ui['BitLockerStatusText'].Foreground  = $bc.ConvertFrom('#DC2626')
    } else {
        $ui['BitLockerStatusText'].Text        = 'Unknown'
        $ui['BitLockerStatusBadge'].Background = $bc.ConvertFrom('#F3F4F6')
        $ui['BitLockerStatusText'].Foreground  = $bc.ConvertFrom('#9CA3AF')
    }

    # ── BitLocker recovery keys ───────────────────────────────
    if (-not $script:currentAadDeviceId) {
        $ui['BitLockerKeysStatus'].Text = 'No Azure AD Device ID — cannot query recovery keys'
        $script:securityLoaded = $true
        return
    }

    try {
        $bkUri  = "https://graph.microsoft.com/v1.0/informationProtection/bitlocker/recoveryKeys" +
                  "?`$filter=deviceId eq '$($script:currentAadDeviceId)'&`$select=id,createdDateTime,volumeType"
        $bkResp = Invoke-MgGraphRequest -Uri $bkUri -Method GET -ErrorAction Stop

        if ($bkResp -and $bkResp.value -and $bkResp.value.Count -gt 0) {
            $ui['BitLockerKeysStatus'].Text = "$($bkResp.value.Count) recovery key(s) found"

            foreach ($key in ($bkResp.value | Sort-Object createdDateTime -Descending)) {
                $keyDate = if ($key.createdDateTime) {
                    ([datetime]$key.createdDateTime).ToString('yyyy-MM-dd HH:mm')
                } else { 'Unknown date' }
                $volType = if ($key.volumeType) { $key.volumeType } else { 'Unknown' }

                $keyCard = [System.Windows.Controls.Border]::new()
                $keyCard.Background      = $bc.ConvertFrom('#FFFFFF')
                $keyCard.BorderBrush     = $bc.ConvertFrom('#E5E7EB')
                $keyCard.BorderThickness = [System.Windows.Thickness]::new(1)
                $keyCard.CornerRadius    = [System.Windows.CornerRadius]::new(6)
                $keyCard.Padding         = [System.Windows.Thickness]::new(10,8,10,8)
                $keyCard.Margin          = [System.Windows.Thickness]::new(0,0,0,6)

                $cardStack = [System.Windows.Controls.StackPanel]::new()

                # ── Top row: meta info ──────────────────────────────────
                $keyIdTb = [System.Windows.Controls.TextBlock]::new()
                $keyIdTb.Text         = "ID: $($key.id)"
                $keyIdTb.FontSize     = 11
                $keyIdTb.FontFamily   = [System.Windows.Media.FontFamily]::new('Consolas')
                $keyIdTb.Foreground   = $bc.ConvertFrom('#374151')
                $keyIdTb.TextWrapping = [System.Windows.TextWrapping]::Wrap

                $keyMetaTb = [System.Windows.Controls.TextBlock]::new()
                $keyMetaTb.Text       = "$volType  ·  $keyDate"
                $keyMetaTb.FontSize   = 11
                $keyMetaTb.Foreground = $bc.ConvertFrom('#9CA3AF')
                $keyMetaTb.Margin     = [System.Windows.Thickness]::new(0,2,0,4)

                $cardStack.Children.Add($keyIdTb)   | Out-Null
                $cardStack.Children.Add($keyMetaTb) | Out-Null

                # ── Key value row (hidden until Reveal) ─────────────────
                $keyValueTb = [System.Windows.Controls.TextBlock]::new()
                $keyValueTb.Text         = '••••••••••••••••  (click Reveal)'
                $keyValueTb.FontSize     = 12
                $keyValueTb.FontFamily   = [System.Windows.Media.FontFamily]::new('Consolas')
                $keyValueTb.Foreground   = $bc.ConvertFrom('#9CA3AF')
                $keyValueTb.TextWrapping = [System.Windows.TextWrapping]::Wrap
                $keyValueTb.Margin       = [System.Windows.Thickness]::new(0,0,0,6)

                $cardStack.Children.Add($keyValueTb) | Out-Null

                # ── Buttons row ─────────────────────────────────────────
                $btnRow = [System.Windows.Controls.StackPanel]::new()
                $btnRow.Orientation = [System.Windows.Controls.Orientation]::Horizontal

                # Reveal/Hide button
                $revealBtnTb = [System.Windows.Controls.TextBlock]::new()
                $revealBtnTb.Text     = 'Reveal'
                $revealBtnTb.FontSize = 11
                $revealBtnTb.Foreground = $bc.ConvertFrom('#5B5FC7')

                $revealBtn = [System.Windows.Controls.Border]::new()
                $revealBtn.Background        = $bc.ConvertFrom('#EEF2FF')
                $revealBtn.CornerRadius      = [System.Windows.CornerRadius]::new(4)
                $revealBtn.Padding           = [System.Windows.Thickness]::new(8,3,8,3)
                $revealBtn.Margin            = [System.Windows.Thickness]::new(0,0,6,0)
                $revealBtn.Cursor            = [System.Windows.Input.Cursors]::Hand
                $revealBtn.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
                $revealBtn.Child             = $revealBtnTb
                # Store the key ID in Tag; the fetched key value will go into Tag once loaded
                $revealBtn.Tag = [PSCustomObject]@{ KeyId = $key.id; KeyValue = $null; ValueTb = $keyValueTb; LabelTb = $revealBtnTb }

                # Copy button
                $copyBtnTb = [System.Windows.Controls.TextBlock]::new()
                $copyBtnTb.Text       = 'Copy'
                $copyBtnTb.FontSize   = 11
                $copyBtnTb.Foreground = $bc.ConvertFrom('#16A34A')

                $copyBtn = [System.Windows.Controls.Border]::new()
                $copyBtn.Background        = $bc.ConvertFrom('#F0FDF4')
                $copyBtn.CornerRadius      = [System.Windows.CornerRadius]::new(4)
                $copyBtn.Padding           = [System.Windows.Thickness]::new(8,3,8,3)
                $copyBtn.Cursor            = [System.Windows.Input.Cursors]::Hand
                $copyBtn.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
                $copyBtn.Child             = $copyBtnTb
                # Checkmark indicator
                $copyCheckTb = [System.Windows.Controls.TextBlock]::new()
                $copyCheckTb.Text              = [string][char]0x2713
                $copyCheckTb.FontSize          = 13
                $copyCheckTb.FontWeight        = [System.Windows.FontWeights]::Bold
                $copyCheckTb.Foreground        = $bc.ConvertFrom('#16A34A')
                $copyCheckTb.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
                $copyCheckTb.Margin            = [System.Windows.Thickness]::new(6,0,0,0)
                $copyCheckTb.Opacity           = 0

                $copyBtn.Tag               = [PSCustomObject]@{ RevealBtn = $revealBtn; CheckTb = $copyCheckTb }

                $revealBtn.Add_MouseLeftButtonDown({
                    param($sender, $e)
                    $state = $sender.Tag
                    $bc2 = [System.Windows.Media.BrushConverter]::new()
                    if ($state.LabelTb.Text -eq 'Reveal') {
                        # Fetch if not yet loaded
                        if (-not $state.KeyValue) {
                            try {
                                $resp = Invoke-MgGraphRequest `
                                    -Uri "https://graph.microsoft.com/v1.0/informationProtection/bitlocker/recoveryKeys/$($state.KeyId)?`$select=key" `
                                    -Method GET -ErrorAction Stop
                                $state.KeyValue = $resp.key
                            } catch {
                                $errMsg = $_.Exception.Message
                                if ($errMsg -match '403|Forbidden') {
                                    $state.ValueTb.Text      = 'Permission denied (BitlockerKey.Read.All required)'
                                    $state.ValueTb.Foreground = $bc2.ConvertFrom('#DC2626')
                                } else {
                                    $state.ValueTb.Text      = "Error: $errMsg"
                                    $state.ValueTb.Foreground = $bc2.ConvertFrom('#DC2626')
                                }
                                return
                            }
                        }
                        $state.ValueTb.Text       = $state.KeyValue
                        $state.ValueTb.Foreground = $bc2.ConvertFrom('#1A1B25')
                        $state.LabelTb.Text       = 'Hide'
                    } else {
                        $state.ValueTb.Text       = '••••••••••••••••  (click Reveal)'
                        $state.ValueTb.Foreground = $bc2.ConvertFrom('#9CA3AF')
                        $state.LabelTb.Text       = 'Reveal'
                    }
                })

                $copyBtn.Add_MouseLeftButtonDown({
                    param($sender, $e)
                    $state = $sender.Tag.RevealBtn.Tag   # revealBtn.Tag = state object
                    $checkTb = $sender.Tag.CheckTb
                    if ($state.KeyValue) {
                        [System.Windows.Clipboard]::SetText($state.KeyValue)
                        Show-CopyConfirmation $checkTb
                    } else {
                        # Fetch silently then copy
                        try {
                            $resp = Invoke-MgGraphRequest `
                                -Uri "https://graph.microsoft.com/v1.0/informationProtection/bitlocker/recoveryKeys/$($state.KeyId)?`$select=key" `
                                -Method GET -ErrorAction Stop
                            $state.KeyValue = $resp.key
                            [System.Windows.Clipboard]::SetText($state.KeyValue)
                            Show-CopyConfirmation $checkTb
                        } catch {
                            $errMsg = $_.Exception.Message
                            if ($errMsg -match '403|Forbidden') {
                                [System.Windows.MessageBox]::Show(
                                    "Permission denied.`nRequires BitlockerKey.Read.All scope.",
                                    "Access Denied", "OK", "Warning") | Out-Null
                            } else {
                                [System.Windows.MessageBox]::Show(
                                    "Could not retrieve key:`n$errMsg",
                                    "Error", "OK", "Error") | Out-Null
                            }
                        }
                    }
                })

                $btnRow.Children.Add($revealBtn) | Out-Null
                $btnRow.Children.Add($copyBtn)   | Out-Null
                $btnRow.Children.Add($copyCheckTb) | Out-Null
                $cardStack.Children.Add($btnRow) | Out-Null
                $keyCard.Child = $cardStack
                $ui['BitLockerKeysList'].Children.Add($keyCard) | Out-Null
            }
        } else {
            $ui['BitLockerKeysStatus'].Text = 'No recovery keys found'
        }
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match '403|Forbidden|Authorization') {
            $ui['BitLockerKeysStatus'].Text = 'No permission (BitlockerKey.ReadBasic.All required)'
        } else {
            $ui['BitLockerKeysStatus'].Text = "Could not load recovery keys: $($_.Exception.Message)"
        }
    }

    $script:securityLoaded = $true
}

# ═══════════════════════════════════════════════════════════
#  LOAD GROUPS
# ═══════════════════════════════════════════════════════════
function New-GroupCard {
    param($group)
    $bc = [System.Windows.Media.BrushConverter]::new()

    $isM365 = $group.groupTypes -and ($group.groupTypes -contains 'Unified')
    $typeText  = if ($isM365) { 'Microsoft 365' } else { 'Security' }
    $badgeR    = if ($isM365) { [byte]37  } else { [byte]107 }
    $badgeG    = if ($isM365) { [byte]99  } else { [byte]114 }
    $badgeB    = if ($isM365) { [byte]235 } else { [byte]128 }
    $typeFg    = [System.Windows.Media.SolidColorBrush]::new(
                     [System.Windows.Media.Color]::FromArgb(255, $badgeR, $badgeG, $badgeB))
    $badgeBg   = [System.Windows.Media.SolidColorBrush]::new(
                     [System.Windows.Media.Color]::FromArgb(26, $badgeR, $badgeG, $badgeB))

    $card = [System.Windows.Controls.Border]::new()
    $card.CornerRadius    = [System.Windows.CornerRadius]::new(6)
    $card.BorderThickness = [System.Windows.Thickness]::new(1)
    $card.Margin          = [System.Windows.Thickness]::new(0,0,0,6)
    $card.Padding         = [System.Windows.Thickness]::new(12,8,12,8)
    $card.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'ThBorder')
    $card.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty,  'ThSurface')

    $row = [System.Windows.Controls.Grid]::new()
    $c1  = [System.Windows.Controls.ColumnDefinition]::new(); $c1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    $c2  = [System.Windows.Controls.ColumnDefinition]::new(); $c2.Width = [System.Windows.GridLength]::Auto
    $row.ColumnDefinitions.Add($c1)
    $row.ColumnDefinitions.Add($c2)

    # Name + description
    $sp = [System.Windows.Controls.StackPanel]::new()

    $nameTb = [System.Windows.Controls.TextBlock]::new()
    $nameTb.Text       = if ($group.displayName) { $group.displayName } else { '—' }
    $nameTb.FontSize   = 13
    $nameTb.FontWeight = [System.Windows.FontWeights]::Medium
    $nameTb.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThTxt1')
    $sp.Children.Add($nameTb) | Out-Null

    if ($group.description) {
        $descTb = [System.Windows.Controls.TextBlock]::new()
        $descTb.Text         = $group.description
        $descTb.FontSize     = 11
        $descTb.TextWrapping = [System.Windows.TextWrapping]::Wrap
        $descTb.Margin       = [System.Windows.Thickness]::new(0,2,0,0)
        $descTb.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThTxt3')
        $sp.Children.Add($descTb) | Out-Null
    }

    [System.Windows.Controls.Grid]::SetColumn($sp, 0)
    $row.Children.Add($sp) | Out-Null

    # Type badge
    $badge = [System.Windows.Controls.Border]::new()
    $badge.CornerRadius      = [System.Windows.CornerRadius]::new(4)
    $badge.Background        = $badgeBg
    $badge.BorderThickness   = [System.Windows.Thickness]::new(0)
    $badge.Padding           = [System.Windows.Thickness]::new(8,3,8,3)
    $badge.Margin            = [System.Windows.Thickness]::new(10,0,0,0)
    $badge.VerticalAlignment = [System.Windows.VerticalAlignment]::Center

    $typeTb = [System.Windows.Controls.TextBlock]::new()
    $typeTb.Text       = $typeText
    $typeTb.FontSize   = 10
    $typeTb.FontWeight = [System.Windows.FontWeights]::SemiBold
    $typeTb.Foreground = $typeFg
    $badge.Child = $typeTb

    [System.Windows.Controls.Grid]::SetColumn($badge, 1)
    $row.Children.Add($badge) | Out-Null

    $card.Child = $row
    return $card
}

function Load-DeviceGroups {
    param([string]$deviceId)

    $ui['GroupsDeviceList'].Children.Clear()
    $ui['GroupsUserList'].Children.Clear()
    $ui['GroupsDeviceStatus'].Text  = 'Loading...'
    $ui['GroupsUserStatus'].Text    = 'Loading...'
    $ui['GroupsActiveStatus'].Text  = 'Loading...'
    Push-UI

    # ── Device groups ──
    $script:allDeviceGroups = @()
    try {
        $aadId = $script:currentAadDeviceId
        if ($aadId) {
            $devResp = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/devices?`$filter=deviceId eq '$aadId'&`$select=id,displayName" -Method GET -ErrorAction Stop
            if ($devResp -and $devResp.value -and $devResp.value.Count -gt 0) {
                $objId   = $devResp.value[0].id
                $memResp = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/devices/$objId/memberOf?`$select=displayName,description,groupTypes&`$top=200" -Method GET -ErrorAction Stop
                $script:allDeviceGroups = if ($memResp -and $memResp.value) { @($memResp.value | Sort-Object displayName) } else { @() }
            } else {
                $ui['GroupsDeviceStatus'].Text = 'Device not found in Azure AD'
            }
        } else {
            $ui['GroupsDeviceStatus'].Text = 'No Azure AD device ID available'
        }
    } catch {
        $ui['GroupsDeviceStatus'].Text = "Error: $($_.Exception.Message.Split([char]10)[0])"
    }

    # ── User groups ──
    $script:allUserGroups = @()
    try {
        $upn = $script:currentPrimaryUpn
        if ($upn) {
            $safeUpn = [Uri]::EscapeDataString($upn)
            $memResp = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/users/$safeUpn/memberOf?`$select=displayName,description,groupTypes&`$top=200" -Method GET -ErrorAction Stop
            $script:allUserGroups = if ($memResp -and $memResp.value) { @($memResp.value | Sort-Object displayName) } else { @() }
        } else {
            $ui['GroupsUserStatus'].Text = 'No primary user assigned to this device'
        }
    } catch {
        $ui['GroupsUserStatus'].Text = "Error: $($_.Exception.Message.Split([char]10)[0])"
    }

    $script:groupsLoaded = $true
    Filter-GroupsList
}

$script:allDeviceGroups = @()
$script:allUserGroups   = @()
$script:groupsSubTab    = 'Device'

function Filter-GroupsList {
    $query    = $ui['GroupsSearchBox'].Text.Trim()
    $isDevice = ($script:groupsSubTab -eq 'Device')
    $source   = if ($isDevice) { $script:allDeviceGroups } else { $script:allUserGroups }
    $listCtrl = if ($isDevice) { $ui['GroupsDeviceList']  } else { $ui['GroupsUserList']  }
    $statCtrl = if ($isDevice) { $ui['GroupsDeviceStatus']} else { $ui['GroupsUserStatus']}

    $listCtrl.Children.Clear()
    $filtered = if ([string]::IsNullOrEmpty($query)) {
        $source
    } else {
        $source | Where-Object { $_.displayName -like "*$query*" -or $_.description -like "*$query*" }
    }
    if ($null -eq $filtered) { $filtered = @() }
    $count = @($filtered).Count
    if ($count -eq 0) {
        $statCtrl.Text = if ([string]::IsNullOrEmpty($query)) { 'No group memberships found' } else { "No results for ""$query""" }
    } else {
        $total = @($source).Count
        $statCtrl.Text = if ([string]::IsNullOrEmpty($query)) { "$total group(s)" } else { "$count of $total group(s)" }
    }
    $ui['GroupsActiveStatus'].Text = ''
    foreach ($g in @($filtered)) { $listCtrl.Children.Add((New-GroupCard $g)) | Out-Null }
}

function Switch-GroupsSubTab ([string]$sub) {
    $script:groupsSubTab = $sub
    $bc       = [System.Windows.Media.BrushConverter]::new()
    $activeBg = $bc.ConvertFrom('#3B82F6')
    $activeFg = [System.Windows.Media.Brushes]::White
    $inactBg  = $window.Resources['ThBorder']
    $inactFg  = $window.Resources['ThTxt2']

    if ($sub -eq 'Device') {
        $ui['GrpSubDevicePanel'].Visibility = 'Visible'
        $ui['GrpSubUserPanel'].Visibility   = 'Collapsed'
        $ui['GrpSubDeviceBtn'].Background   = $activeBg
        ($ui['GrpSubDeviceBtn'].Child).Foreground = $activeFg
        $ui['GrpSubUserBtn'].Background     = $inactBg
        ($ui['GrpSubUserBtn'].Child).Foreground   = $inactFg
    } else {
        $ui['GrpSubDevicePanel'].Visibility = 'Collapsed'
        $ui['GrpSubUserPanel'].Visibility   = 'Visible'
        $ui['GrpSubDeviceBtn'].Background   = $inactBg
        ($ui['GrpSubDeviceBtn'].Child).Foreground = $inactFg
        $ui['GrpSubUserBtn'].Background     = $activeBg
        ($ui['GrpSubUserBtn'].Child).Foreground   = $activeFg
    }
    Filter-GroupsList
}

# Initialize Groups sub-tab appearance (Device active by default)
Switch-GroupsSubTab 'Device'

$ui['GrpSubDeviceBtn'].Add_MouseLeftButtonDown({ Switch-GroupsSubTab 'Device' })
$ui['GrpSubUserBtn'].Add_MouseLeftButtonDown({ Switch-GroupsSubTab 'User' })

$ui['GroupsSearchBox'].Add_TextChanged({ if ($script:groupsLoaded) { Filter-GroupsList } })

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
        # Run Connect-MgGraph in a background runspace so the browser
        # popup can complete without the WPF UI thread blocking it.
        # Disconnect inside the same runspace so the token cache is cleared
        # before the new Connect — avoids reuse of a cached session with missing scopes.
        # Clear the entire .IdentityService folder so stale PIM/role tokens are not
        # silently reused after a privilege change (e.g. PIM activation).
        $identityServiceDir = Join-Path $env:LOCALAPPDATA '.IdentityService'
        if (Test-Path $identityServiceDir) {
            try { Remove-Item $identityServiceDir -Recurse -Force -ErrorAction SilentlyContinue } catch {}
        }

        $rs = [runspacefactory]::CreateRunspace()
        $rs.Open()
        $ps = [powershell]::Create().AddScript({
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
            try { Disconnect-MgGraph -ErrorAction Stop } catch {}
            Connect-MgGraph -TenantId 'organizations' -Scopes @(
                'DeviceManagementManagedDevices.ReadWrite.All',
                'DeviceManagementManagedDevices.PrivilegedOperations.All',
                'User.Read.All',
                'DeviceManagementServiceConfig.ReadWrite.All',
                'DeviceLocalCredential.Read.All',
                'BitlockerKey.ReadBasic.All',
                'BitlockerKey.Read.All',
                'Organization.Read.All',
                'GroupMember.Read.All',
                'Device.Read.All'
            ) -NoWelcome -ErrorAction Stop
        })
        $ps.Runspace = $rs
        $handle = $ps.BeginInvoke()

            # Keep the WPF dispatcher alive while waiting
            while (-not $handle.IsCompleted) {
                Push-UI
                Start-Sleep -Milliseconds 100
            }

            # Collect result / errors — skip errors caused only by Disconnect having no session,
            # and skip informational scope-warning records emitted by Connect-MgGraph on success
            $ps.EndInvoke($handle)
            $realErrors = $ps.Streams.Error | Where-Object {
                $msg = if ($_.Exception.Message) { $_.Exception.Message } else { $_.ToString() }
                $msg -notmatch 'No application|sign out|not connected|no active|no connection'
            }
            if ($realErrors) {
                $errParts = foreach ($e in $realErrors) {
                    if ($e.Exception.Message)                    { $e.Exception.Message }
                    elseif ($e.Exception.InnerException.Message) { $e.Exception.InnerException.Message }
                    elseif ($e.ToString())                       { $e.ToString() }
                    else                                         { $e.CategoryInfo.ToString() }
                }
                $errMsg = ($errParts | Where-Object { $_ }) -join "`n"
                if (-not $errMsg) { $errMsg = 'Unknown sign-in error' }
                throw $errMsg
            }
            $ps.Dispose()
            $rs.Close()

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

        # Verify critical scopes are present in the token
        $required = @(
            'DeviceManagementManagedDevices.ReadWrite.All',
            'DeviceManagementManagedDevices.PrivilegedOperations.All'
        )
        $ctx = Get-MgContext
        $grantedScopes = $ctx.Scopes
        $missing = $required | Where-Object { $_ -notin $grantedScopes }
        if ($missing) {
            [System.Windows.MessageBox]::Show(
                "Signed in, but the following permissions were not consented:`n`n$($missing -join "`n")`n`nThis can happen when an admin has not granted tenant-wide consent, or the token cache held an old session.`n`nThe cache has been cleared — please sign out and sign in again to re-consent.",
                'Missing Permissions',
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            ) | Out-Null
        }

        $ui['StatusText'].Text      = $tenantName
        $ui['LoginButton'].Content  = 'Sign Out'
        $ui['SearchBox'].IsEnabled  = $true
        $ui['SearchButton'].IsEnabled = $true
        $ui['SearchHint'].Text      = 'Search by name, email, or device name'
    } catch {
        $msg = $_.Exception.Message
        # Silently ignore user-cancelled authentication — not an error
        if ($msg -notmatch 'cancel|abort|user canceled|User canceled') {
            [System.Windows.MessageBox]::Show(
                "Sign-in failed:`n$msg", "Authentication Error", "OK", "Error") | Out-Null
        }
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
           "&`$select=id,deviceName,operatingSystem,osVersion,serialNumber,model,manufacturer,lastSyncDateTime,complianceState" +
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

    $bcDev = [System.Windows.Media.BrushConverter]::new()
    foreach ($d in $sorted) {
        $lastSync = if ($d.lastSyncDateTime) {
            'Last sync: ' + ([datetime]$d.lastSyncDateTime).ToString('yyyy-MM-dd HH:mm')
        } else { 'Last sync: unknown' }
        $compState = if ($d.complianceState) { $d.complianceState } else { 'unknown' }
        $compColor = switch ($compState) {
            'compliant'      { '#16A34A' }  # green
            'noncompliant'   { '#DC2626' }  # red
            'inGracePeriod'  { '#EA580C' }  # orange
            'configManager'  { '#5B5FC7' }  # indigo
            default          { '#9CA3AF' }  # grey (unknown / notApplicable)
        }
        $compLabel = switch ($compState) {
            'compliant'      { 'Compliant' }
            'noncompliant'   { 'Not compliant' }
            'inGracePeriod'  { 'In grace period' }
            'configManager'  { 'Managed by Config Manager' }
            'notApplicable'  { 'Not applicable' }
            default          { 'Unknown' }
        }
        $ui['DeviceList'].Items.Add([PSCustomObject]@{
            DeviceName       = $d.deviceName
            OS               = "$($d.operatingSystem) $($d.osVersion)"
            DeviceId         = $d.id
            Serial           = $d.serialNumber
            LastSync         = $lastSync
            ComplianceColor  = $bcDev.ConvertFrom($compColor)
            ComplianceLabel  = $compLabel
        }) | Out-Null
    }
})

# ═══════════════════════════════════════════════════════════
#  LOAD DEVICE DETAILS (reusable)
# ═══════════════════════════════════════════════════════════
function Load-DeviceDetails ([string]$deviceId) {
    $ui['EmptyState'].Visibility   = 'Collapsed'
    $ui['DetailPanel'].Visibility  = 'Collapsed'
    $ui['LoadingState'].Visibility = 'Visible'
    Push-UI

    $id = $deviceId

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

    if ($loggedOn) {
        # Ensure it's an array (single-item collections may unwrap)
        $loggedOnList = @($loggedOn)
        if ($loggedOnList.Count -gt 0) {
            $sorted = $loggedOnList | Sort-Object { [datetime]$_.lastLogOnDateTime } -Descending
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
    }

    # Fallback: if usersLoggedOn was empty, use the primary user from the device record
    if ($lastLogon -eq 'N/A' -and $primaryUser -ne 'N/A') {
        $lastLogon    = $primaryUser
        $allUsersText = $primaryUser
    }

    # ── Enrolled date ──
    $enrolled = 'N/A'
    if ($dev.enrolledDateTime) {
        $enrolled = ([datetime]$dev.enrolledDateTime).ToString('yyyy-MM-dd HH:mm')
    }

    # ── Installed Enrollment profile (recorded at enrollment time) ──
    $enrollProfile = if ($dev.enrollmentProfileName) { $dev.enrollmentProfileName } else { 'N/A' }

    # ── Assigned Enrollment Profile (current Autopilot deployment profile) ──
    # Fetched after the Autopilot lookup below; initialise here.
    $assignedEnrollProfile = 'N/A'

    # ── Autopilot Group Tag ──
    # $filter is not supported on this endpoint — page through all identities and match client-side.
    # Use $top=500 and $select to minimise round-trips and payload size.
    $groupTag = 'N/A'
    $script:currentAutopilotId        = $null
    $script:autopilotAssignmentStatus = $null
    $script:autopilotProfileName      = $null

    if ($dev.serialNumber) {
        try {
            $serial  = $dev.serialNumber.Trim()
            $apMatch = $null
            $apUri   = "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?`$top=100"

            while ($apUri -and -not $apMatch) {
                $apPage  = Invoke-MgGraphRequest -Uri $apUri -Method GET -ErrorAction Stop
                $apMatch = @($apPage.value) | Where-Object {
                    $sn = if ($_['serialNumber']) { $_['serialNumber'].Trim() } else { '' }
                    $sn -eq $serial
                } | Select-Object -First 1
                $apUri = if (-not $apMatch) { $apPage.'@odata.nextLink' } else { $null }
            }

            if ($apMatch -and $apMatch['id']) {
                $script:currentAutopilotId        = $apMatch['id']
                $groupTag                         = if ($apMatch['groupTag']) { $apMatch['groupTag'] } else { '(none)' }
                $script:autopilotAssignmentStatus = $apMatch['deploymentProfileAssignmentStatus']
                $script:autopilotProfileName      = if ($apMatch['deploymentProfileAssignedDateTime']) {
                    'Assigned ' + ([datetime]$apMatch['deploymentProfileAssignedDateTime']).ToString('yyyy-MM-dd')
                } else { $null }
            }
            # Fetch the currently assigned deployment profile name via the same
            # Autopilot identity endpoint using $expand=deploymentProfile.
            if ($script:currentAutopilotId) {
                $apId = $script:currentAutopilotId
                $assignedEnrollProfile = 'Not assigned'
                try {
                    $apDetail = Invoke-MgGraphRequest `
                        -Uri "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities/$apId`?`$expand=deploymentProfile(`$select=displayName)" `
                        -Method GET -ErrorAction Stop
                    $profileName = if ($apDetail['deploymentProfile'] -and $apDetail['deploymentProfile']['displayName']) {
                        $apDetail['deploymentProfile']['displayName']
                    } else { $null }
                    if ($profileName) { $assignedEnrollProfile = $profileName }
                } catch { <# leave Not assigned if profile lookup fails #> }
            }
            # $apMatch null → device not Autopilot-registered → leave N/A
        } catch {
            $errLine  = $_.Exception.Message.Split([char]10)[0]
            $groupTag = "Error: $errLine"
            $script:currentAutopilotId        = $null
            $script:autopilotAssignmentStatus = $null
        }
    }

    # ── Last password change (primary user's Azure AD password) ──
    $lastPwdChange = 'N/A'
    if ($dev.userPrincipalName) {
        try {
            $safeUpnPwd = [Uri]::EscapeDataString($dev.userPrincipalName)
            $uInfo = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/users/${safeUpnPwd}?`$select=lastPasswordChangeDateTime" -Method GET -ErrorAction Stop
            if ($uInfo -and $uInfo.lastPasswordChangeDateTime) {
                $lastPwdChange = ([datetime]$uInfo.lastPasswordChangeDateTime).ToString('yyyy-MM-dd HH:mm')
            }
        } catch { <# permission not granted or user not found — leave N/A #> }
    }

    # ── BIOS / UEFI Version (hardwareInformation via beta, Windows only) ──
    $biosVersion = 'N/A'
    if ($dev.operatingSystem -match '(?i)windows') {
        try {
            $hwInfo = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices/${id}?`$select=hardwareInformation" -Method GET -ErrorAction Stop
            # Use bracket notation — Invoke-MgGraphRequest returns nested objects as Hashtable
            $hwBios = if ($hwInfo -and $hwInfo['hardwareInformation']) { $hwInfo['hardwareInformation']['systemManagementBIOSVersion'] } else { $null }
            if ($hwBios) { $biosVersion = $hwBios }
        } catch { <# beta endpoint or permission not available — leave N/A #> }
    }

    # ── Defender / AV (windowsProtectionState via beta, Windows only) ──
    $wps            = $null
    $defenderStatus = 'N/A'
    if ($dev.operatingSystem -match '(?i)windows') {
        try {
            $wpsResp = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices/${id}/windowsProtectionState" -Method GET -ErrorAction Stop
            if ($wpsResp) {
                $wps = $wpsResp

                # Safely extract booleans — Graph returns $null for unreported fields, not $false
                $rtpProp = $wps.realTimeProtectionEnabled
                $avProp  = $wps.antivirusEnabled          # may be $null on modern Defender reporting
                $malProp = $wps.malwareProtectionEnabled  # more reliable fallback
                $sigOld  = $wps.signatureUpdateOverdue -eq $true

                # Treat both key protections as "not reported" if all are null
                $hasData = ($null -ne $rtpProp) -or ($null -ne $avProp) -or ($null -ne $malProp)
                $protOn  = ($rtpProp -eq $true)
                $avOn    = ($avProp  -eq $true) -or ($malProp -eq $true)

                $lastScan = if ($wps.lastReportedDateTime) {
                    ' · Last scan: ' + ([datetime]$wps.lastReportedDateTime).ToString('yyyy-MM-dd HH:mm')
                } else { '' }
                $sigDate = if ($wps.antivirusSignatureLastUpdateDateTime) {
                    ' · Sigs: ' + ([datetime]$wps.antivirusSignatureLastUpdateDateTime).ToString('yyyy-MM-dd')
                } else { '' }

                if (-not $hasData) {
                    $defenderStatus = "Not reporting$lastScan"
                } elseif ($protOn -and $avOn -and -not $sigOld) {
                    $defenderStatus = "Protected$sigDate$lastScan"
                } elseif ($sigOld) {
                    $defenderStatus = "Signatures outdated$sigDate$lastScan"
                } elseif (-not $protOn -and $null -ne $rtpProp) {
                    $defenderStatus = "Real-time protection disabled$lastScan"
                } elseif (-not $avOn -and ($null -ne $avProp -or $null -ne $malProp)) {
                    $defenderStatus = "Antivirus disabled$lastScan"
                } else {
                    $defenderStatus = "Needs attention$lastScan"
                }
            }
        } catch { <# DeviceManagementManagedDevices.Read.All required; or endpoint not available — leave N/A #> }
    }

    # ── Store device ID for lazy-loaded tabs ──
    $script:currentDeviceId    = $id
    $script:currentAadDeviceId = if ($dev.azureADDeviceId) { $dev.azureADDeviceId } else { $null }
    $script:currentPrimaryUpn  = if ($dev.userPrincipalName) { $dev.userPrincipalName } else { $null }
    $script:deviceIsEncrypted  = $dev.isEncrypted
    $script:deviceOS           = if ($dev.operatingSystem) { $dev.operatingSystem.ToLower() } else { '' }
    $script:deviceComplianceState   = if ($dev.complianceState) { $dev.complianceState } else { 'unknown' }
    $script:deviceLastSync          = if ($dev.lastSyncDateTime) { [datetime]$dev.lastSyncDateTime } else { $null }
    $script:deviceGracePeriodExpiry = if ($dev.complianceGraceperiodExpirationDateTime) { [datetime]$dev.complianceGraceperiodExpirationDateTime } else { $null }
    $script:lapsPassword            = $null
    $script:appsLoaded         = $false
    $script:complianceLoaded   = $false
    $script:securityLoaded     = $false
    $script:groupsLoaded       = $false

    # ── Populate UI ──
    $ui['DetailTitle'].Text    = $dev.deviceName
    $ui['DetailSubtitle'].Text = "$($dev.manufacturer) $($dev.model) — S/N $($dev.serialNumber)"

    $ui['ValDeviceName'].Text   = $dev.deviceName
    $ui['ValOSVersion'].Text    = $osFriendly
    $ui['ValOSBuild'].Text      = $osBuild
    $ui['ValLastLogon'].Text     = $lastLogon
    $ui['ValPrimaryUser'].Text  = $primaryUser
    $ui['ValAllUsers'].Text      = $allUsersText
    $ui['ValEnrollProfile'].Text         = $enrollProfile
    $ui['ValAssignedEnrollProfile'].Text = $assignedEnrollProfile
    $ui['ValGroupTag'].Text              = $groupTag

    # Show Autopilot assignment status badge
    $bc = [System.Windows.Media.BrushConverter]::new()
    if ($script:currentAutopilotId) {
        $rawStatus = if ($script:autopilotAssignmentStatus) { $script:autopilotAssignmentStatus } else { 'unknown' }
        $statusLabel = switch ($rawStatus) {
            'assigned'                { 'Assigned' }
            'assignedUnkownSyncState' { 'Assigned' }
            'assignedInSync'          { 'Assigned (In Sync)' }
            'assignedOutOfSync'       { 'Assigned (Out of Sync)' }
            'notAssigned'             { 'Not Assigned' }
            'pending'                 { 'Pending' }
            'failed'                  { 'Failed' }
            'unknown'                 { 'Unknown' }
            default                   { $rawStatus }
        }
        $statusColor = switch ($rawStatus) {
            'assigned'                { @{ Bg='#F0FDF4'; Fg='#16A34A' } }
            'assignedUnkownSyncState' { @{ Bg='#F0FDF4'; Fg='#16A34A' } }
            'assignedInSync'          { @{ Bg='#F0FDF4'; Fg='#16A34A' } }
            'assignedOutOfSync'       { @{ Bg='#FEF3C7'; Fg='#D97706' } }
            'pending'                 { @{ Bg='#FEF3C7'; Fg='#D97706' } }
            'failed'                  { @{ Bg='#FEF2F2'; Fg='#DC2626' } }
            'notAssigned'             { @{ Bg='#F3F4F6'; Fg='#6B7280' } }
            default                   { @{ Bg='#F3F4F6'; Fg='#9CA3AF' } }
        }
        $statusTooltip = switch ($rawStatus) {
            'assigned'                { "Profile assigned — device is ready for Autopilot enrollment with this group tag.$(if ($script:autopilotProfileName) { "`n$($script:autopilotProfileName)" })" }
            'assignedUnkownSyncState' { "Profile assigned — waiting for device to confirm receipt.$(if ($script:autopilotProfileName) { "`n$($script:autopilotProfileName)" })" }
            'assignedInSync'          { "Profile assigned and in sync — device is fully ready for Autopilot.$(if ($script:autopilotProfileName) { "`n$($script:autopilotProfileName)" })" }
            'assignedOutOfSync'       { "Profile assigned but out of sync — sync the device to update.$(if ($script:autopilotProfileName) { "`n$($script:autopilotProfileName)" })" }
            'notAssigned'             { 'No Autopilot profile is assigned. Verify the group tag matches a deployment profile.' }
            'pending'                 { 'Intune is processing the group tag change. Sync the device and refresh to check again.' }
            'failed'                  { 'Profile assignment failed. Verify the group tag matches an Autopilot deployment profile.' }
            default                   { 'Assignment status could not be determined. Try refreshing the device details.' }
        }
        $ui['GroupTagStatus'].Background          = $bc.ConvertFrom($statusColor.Bg)
        $ui['GroupTagStatusText'].Text            = $statusLabel
        $ui['GroupTagStatusText'].Foreground      = $bc.ConvertFrom($statusColor.Fg)
        $ui['GroupTagStatus'].ToolTip             = $statusTooltip
        $ui['GroupTagStatus'].Visibility          = 'Visible'
    } else {
        $ui['GroupTagStatus'].Visibility = 'Collapsed'
    }
    $ui['ValLastEnrolled'].Text        = $enrolled
    $ui['ValLastPasswordChange'].Text  = $lastPwdChange
    $ui['ValBiosVersion'].Text         = $biosVersion
    $ui['ValDefenderStatus'].Text      = $defenderStatus
    Update-HealthCard $dev $wps

    # Reset to Details tab
    Switch-Tab 'Info'
    $ui['AppsList'].Items.Clear()
    $ui['AppsSearchBox'].Text = ''
    $script:allFilteredApps = @()
    $ui['ComplianceList'].Children.Clear()
    $ui['ComplianceStatus'].Text = ''
    $ui['ComplianceSyncInfo'].Text = ''
    $ui['ComplianceGraceInfo'].Visibility = 'Collapsed'
    $ui['ComplianceMismatchCard'].Visibility = 'Collapsed'
    $ui['CompSubOverviewPanel'].Visibility = 'Visible'
    $ui['CompSubPoliciesPanel'].Visibility = 'Collapsed'
    $bc2 = [System.Windows.Media.BrushConverter]::new()
    $ui['CompSubOverviewBtn'].Background = $bc2.ConvertFrom('#3B82F6')
    ($ui['CompSubOverviewBtn'].Child).Foreground = [System.Windows.Media.Brushes]::White
    $ui['CompSubPoliciesBtn'].Background = $window.Resources['ThBorder']
    ($ui['CompSubPoliciesBtn'].Child).Foreground = $window.Resources['ThTxt2']
    $ui['DeviceComplianceBadge'].Background = $bc2.ConvertFrom('#F3F4F6')
    $ui['DeviceComplianceText'].Text       = '—'
    $ui['DeviceComplianceText'].Foreground = $bc2.ConvertFrom('#9CA3AF')
    $ui['BitLockerKeysList'].Children.Clear()
    $ui['SecurityStatus'].Text = ''
    $ui['GroupsDeviceList'].Children.Clear()
    $ui['GroupsUserList'].Children.Clear()
    $ui['GroupsDeviceStatus'].Text  = ''
    $ui['GroupsUserStatus'].Text    = ''
    $ui['GroupsActiveStatus'].Text  = ''
    $ui['GroupsSearchBox'].Text     = ''
    $script:allDeviceGroups = @()
    $script:allUserGroups   = @()
    $script:groupsSubTab    = 'Device'
    Switch-GroupsSubTab 'Device'

    $ui['LoadingState'].Visibility = 'Collapsed'
    $ui['DetailPanel'].Visibility  = 'Visible'
}

# ═══════════════════════════════════════════════════════════
#  DEVICE SELECTED → LOAD DETAILS
# ═══════════════════════════════════════════════════════════
$ui['DeviceList'].Add_SelectionChanged({
    $sel = $ui['DeviceList'].SelectedItem
    if (-not $sel) { return }
    Load-DeviceDetails $sel.DeviceId
})

# ═══════════════════════════════════════════════════════════
#  REFRESH BUTTON
# ═══════════════════════════════════════════════════════════
$ui['RefreshBtn'].Add_MouseLeftButtonDown({
    if (-not $script:connected -or -not $script:currentDeviceId) { return }
    Load-DeviceDetails $script:currentDeviceId
})

# ═══════════════════════════════════════════════════════════
#  SYNC DEVICE BUTTON
# ═══════════════════════════════════════════════════════════
$ui['SyncDeviceBtn'].Add_MouseLeftButtonDown({
    if (-not $script:connected -or -not $script:currentDeviceId) { return }
    $id = $script:currentDeviceId

    $sp      = $ui['SyncDeviceBtn'].Child
    $labelTb = $sp.Children[1]

    $labelTb.Text = 'Syncing…'
    $ui['SyncDeviceBtn'].IsEnabled = $false
    Push-UI

    $syncOk = $false
    $errDetail = ''
    try {
        # Use PSObject output so 204 No Content doesn't cause deserialisation errors,
        # but exceptions (4xx/5xx) are still thrown with full response body.
        Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$id/syncDevice" `
            -Method POST -ContentType 'application/json' -ErrorAction Stop | Out-Null
        $syncOk = $true
    } catch {
        # Try to extract the structured Graph error message from the response body
        $errDetail = $_.Exception.Message
        $rawBody = ''
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $rawBody = $_.ErrorDetails.Message }
        elseif ($_.Exception.Response) {
            try {
                $stream = $_.Exception.Response.Content.ReadAsStringAsync().Result
                $rawBody = $stream
            } catch {}
        }
        if ($rawBody) {
            try {
                $parsed = $rawBody | ConvertFrom-Json -ErrorAction SilentlyContinue
                if ($parsed.error.message) { $errDetail = $parsed.error.message }
                elseif ($parsed.error.code) { $errDetail = "$($parsed.error.code): $errDetail" }
            } catch {}
        }

        # Append diagnostics for authentication / authorisation failures
        if ($errDetail -match '401|Unauthorized') {
            $errDetail += "`n`nYour session may have expired. Please sign out and sign back in."
        } elseif ($errDetail -match '403|Forbidden|Authorization') {
            try {
                $ctx = Get-MgContext -ErrorAction SilentlyContinue
                $relevantScopes = $ctx.Scopes | Where-Object { $_ -match 'DeviceManagement' }
                $errDetail += "`n`nToken scopes (DeviceManagement):`n" + ($relevantScopes -join "`n")
                $errDetail += "`n`nIf 'PrivilegedOperations.All' is missing, sign out and sign back in.`nIf it is present, an Intune RBAC role with 'Remote tasks' permission is required."
            } catch {}
        }
    }

    if ($syncOk) {
        $labelTb.Text = 'Sent ✓'
        $ui['SyncDeviceBtn'].IsEnabled = $true
        $t = [System.Windows.Threading.DispatcherTimer]::new()
        $t.Interval = [TimeSpan]::FromSeconds(3)
        $t.Tag = $labelTb
        $t.Add_Tick({ param($s,$e) $s.Tag.Text = 'Sync'; $s.Stop() })
        $t.Start()
    } else {
        $labelTb.Text = 'Error'
        $ui['SyncDeviceBtn'].IsEnabled = $true
        [System.Windows.MessageBox]::Show(
            $errDetail,
            'Sync Failed',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        ) | Out-Null
        $t = [System.Windows.Threading.DispatcherTimer]::new()
        $t.Interval = [TimeSpan]::FromSeconds(4)
        $t.Tag = $labelTb
        $t.Add_Tick({ param($s,$e) $s.Tag.Text = 'Sync'; $s.Stop() })
        $t.Start()
    }
})

# ═══════════════════════════════════════════════════════════
#  CHECK PERMISSIONS
# ═══════════════════════════════════════════════════════════
$ui['CheckPermsButton'].Add_Click({
    # Build a fresh WPF window for the permissions report
    Add-Type -AssemblyName PresentationFramework

    # ── Required permissions with descriptions ─────────────────────
    $requiredScopes = @(
        [PSCustomObject]@{ Scope = 'DeviceManagementManagedDevices.ReadWrite.All';        Purpose = 'Read device data, apps, compliance' }
        [PSCustomObject]@{ Scope = 'DeviceManagementManagedDevices.PrivilegedOperations.All'; Purpose = 'Trigger remote actions (Sync, Fresh Start, etc.)' }
        [PSCustomObject]@{ Scope = 'User.Read.All';                                       Purpose = 'Search Entra ID users by name/UPN/email' }
        [PSCustomObject]@{ Scope = 'DeviceManagementServiceConfig.ReadWrite.All';         Purpose = 'Read/write enrollment config, Autopilot group tags' }
        [PSCustomObject]@{ Scope = 'DeviceLocalCredential.Read.All';                      Purpose = 'Retrieve LAPS passwords for devices' }
        [PSCustomObject]@{ Scope = 'BitlockerKey.ReadBasic.All';                          Purpose = 'List BitLocker recovery key IDs' }
        [PSCustomObject]@{ Scope = 'BitlockerKey.Read.All';                               Purpose = 'Reveal full BitLocker recovery key values' }
        [PSCustomObject]@{ Scope = 'Organization.Read.All';                               Purpose = 'Display tenant name in the header' }
        [PSCustomObject]@{ Scope = 'GroupMember.Read.All';                                Purpose = 'Read device and user group memberships (Groups tab)' }
        [PSCustomObject]@{ Scope = 'Device.Read.All';                                     Purpose = 'Resolve Azure AD device object for group lookups' }
    )

    # ── Determine which scopes are currently consented ───────────
    $grantedScopes = @()
    if ($script:connected) {
        try {
            $ctx = Get-MgContext -ErrorAction SilentlyContinue
            if ($ctx -and $ctx.Scopes) {
                $grantedScopes = $ctx.Scopes
            }
        } catch { }
    }

    # ── Build the permission-check window ────────────────────
    $popupWindow = [System.Windows.Window]::new()
    $popupWindow.Title                  = 'Graph Permission Check'
    $popupWindow.Width                  = 640
    $popupWindow.Height                 = 520
    $popupWindow.MinWidth               = 500
    $popupWindow.MinHeight              = 380
    $popupWindow.WindowStartupLocation  = [System.Windows.WindowStartupLocation]::CenterOwner
    $popupWindow.Owner                  = $window
    $popupWindow.Background             = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#F0F2F5')
    $popupWindow.FontFamily             = [System.Windows.Media.FontFamily]::new('Segoe UI')
    $popupWindow.ResizeMode             = [System.Windows.ResizeMode]::CanResize

    $bc = [System.Windows.Media.BrushConverter]::new()

    $outerStack = [System.Windows.Controls.StackPanel]::new()
    $outerStack.Margin = [System.Windows.Thickness]::new(0)

    # ── Header strip ───────────────────────────────────────
    $headerBorder = [System.Windows.Controls.Border]::new()
    $headerBorder.Background      = $bc.ConvertFrom('#FFFFFF')
    $headerBorder.BorderBrush     = $bc.ConvertFrom('#E5E7EB')
    $headerBorder.BorderThickness = [System.Windows.Thickness]::new(0,0,0,1)
    $headerBorder.Padding         = [System.Windows.Thickness]::new(24,16,24,16)

    $headerInner = [System.Windows.Controls.StackPanel]::new()

    $titleTb = [System.Windows.Controls.TextBlock]::new()
    $titleTb.Text       = 'Microsoft Graph Permissions'
    $titleTb.FontSize   = 16
    $titleTb.FontWeight = [System.Windows.FontWeights]::SemiBold
    $titleTb.Foreground = $bc.ConvertFrom('#1A1B25')

    $subTb = [System.Windows.Controls.TextBlock]::new()
    if ($script:connected) {
        $granted = ($grantedScopes | Measure-Object).Count
        $required = $requiredScopes.Count
        $matched  = ($requiredScopes | Where-Object { $grantedScopes -contains $_.Scope }).Count
        $subTb.Text = "Signed in · $matched of $required required permissions granted"
    } else {
        $subTb.Text = 'Not signed in · sign in to verify actual granted permissions'
    }
    $subTb.FontSize   = 12
    $subTb.Foreground = $bc.ConvertFrom('#9CA3AF')
    $subTb.Margin     = [System.Windows.Thickness]::new(0,4,0,0)

    $headerInner.Children.Add($titleTb) | Out-Null
    $headerInner.Children.Add($subTb)   | Out-Null
    $headerBorder.Child = $headerInner

    # ── Scrollable list ──────────────────────────────────────
    $grid = [System.Windows.Controls.Grid]::new()
    $rDef0 = [System.Windows.Controls.RowDefinition]::new()
    $rDef0.Height = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    $rDef1 = [System.Windows.Controls.RowDefinition]::new()
    $rDef1.Height = [System.Windows.GridLength]::Auto
    $grid.RowDefinitions.Add($rDef0)
    $grid.RowDefinitions.Add($rDef1)

    $scroll = [System.Windows.Controls.ScrollViewer]::new()
    $scroll.VerticalScrollBarVisibility = [System.Windows.Controls.ScrollBarVisibility]::Auto
    $scroll.Margin = [System.Windows.Thickness]::new(20,16,20,0)
    [System.Windows.Controls.Grid]::SetRow($scroll, 0)

    $listPanel = [System.Windows.Controls.StackPanel]::new()

    foreach ($perm in $requiredScopes) {
        $granted = $grantedScopes -contains $perm.Scope

        if ($script:connected) {
            $icon      = if ($granted) { [string][char]0x2713 } else { [string][char]0x2717 }
            $iconColor = if ($granted) { '#16A34A' } else { '#DC2626' }
            $badgeBg   = if ($granted) { '#DCFCE7' } else { '#FEE2E2' }
            $badgeText = if ($granted) { 'Granted' } else { 'Missing' }
            $badgeFg   = $iconColor
        } else {
            $icon      = [string][char]0x2014   # —
            $iconColor = '#9CA3AF'
            $badgeBg   = '#F3F4F6'
            $badgeText = 'Not verified'
            $badgeFg   = '#9CA3AF'
        }

        $row = [System.Windows.Controls.Border]::new()
        $row.Background      = $bc.ConvertFrom('#FFFFFF')
        $row.BorderBrush     = $bc.ConvertFrom('#E5E7EB')
        $row.BorderThickness = [System.Windows.Thickness]::new(1)
        $row.CornerRadius    = [System.Windows.CornerRadius]::new(8)
        $row.Padding         = [System.Windows.Thickness]::new(16,12,16,12)
        $row.Margin          = [System.Windows.Thickness]::new(0,0,0,8)

        $rowGrid = [System.Windows.Controls.Grid]::new()
        $c0 = [System.Windows.Controls.ColumnDefinition]::new()
        $c0.Width = [System.Windows.GridLength]::new(28)
        $c1 = [System.Windows.Controls.ColumnDefinition]::new()
        $c1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $c2 = [System.Windows.Controls.ColumnDefinition]::new()
        $c2.Width = [System.Windows.GridLength]::Auto
        $rowGrid.ColumnDefinitions.Add($c0)
        $rowGrid.ColumnDefinitions.Add($c1)
        $rowGrid.ColumnDefinitions.Add($c2)

        # Icon
        $iconTb = [System.Windows.Controls.TextBlock]::new()
        $iconTb.Text              = $icon
        $iconTb.FontSize          = 15
        $iconTb.FontWeight        = [System.Windows.FontWeights]::Bold
        $iconTb.Foreground        = $bc.ConvertFrom($iconColor)
        $iconTb.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        [System.Windows.Controls.Grid]::SetColumn($iconTb, 0)

        # Scope name + purpose
        $textStack = [System.Windows.Controls.StackPanel]::new()
        $scopeTb = [System.Windows.Controls.TextBlock]::new()
        $scopeTb.Text         = $perm.Scope
        $scopeTb.FontSize     = 13
        $scopeTb.FontWeight   = [System.Windows.FontWeights]::SemiBold
        $scopeTb.Foreground   = $bc.ConvertFrom('#1A1B25')
        $scopeTb.TextWrapping = [System.Windows.TextWrapping]::Wrap

        $purposeTb = [System.Windows.Controls.TextBlock]::new()
        $purposeTb.Text         = $perm.Purpose
        $purposeTb.FontSize     = 11
        $purposeTb.Foreground   = $bc.ConvertFrom('#9CA3AF')
        $purposeTb.TextWrapping = [System.Windows.TextWrapping]::Wrap
        $purposeTb.Margin       = [System.Windows.Thickness]::new(0,2,0,0)

        $textStack.Children.Add($scopeTb)   | Out-Null
        $textStack.Children.Add($purposeTb) | Out-Null
        [System.Windows.Controls.Grid]::SetColumn($textStack, 1)

        # Badge
        $badgeTb = [System.Windows.Controls.TextBlock]::new()
        $badgeTb.Text       = $badgeText
        $badgeTb.FontSize   = 11
        $badgeTb.FontWeight = [System.Windows.FontWeights]::Medium
        $badgeTb.Foreground = $bc.ConvertFrom($badgeFg)

        $badge = [System.Windows.Controls.Border]::new()
        $badge.Background        = $bc.ConvertFrom($badgeBg)
        $badge.CornerRadius      = [System.Windows.CornerRadius]::new(4)
        $badge.Padding           = [System.Windows.Thickness]::new(8,3,8,3)
        $badge.Margin            = [System.Windows.Thickness]::new(12,0,0,0)
        $badge.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        $badge.Child             = $badgeTb
        [System.Windows.Controls.Grid]::SetColumn($badge, 2)

        $rowGrid.Children.Add($iconTb)    | Out-Null
        $rowGrid.Children.Add($textStack) | Out-Null
        $rowGrid.Children.Add($badge)     | Out-Null
        $row.Child = $rowGrid
        $listPanel.Children.Add($row) | Out-Null
    }

    $scroll.Content = $listPanel
    $grid.Children.Add($scroll) | Out-Null

    # ── Footer with note ───────────────────────────────────────
    $footerBorder = [System.Windows.Controls.Border]::new()
    $footerBorder.Background      = $bc.ConvertFrom('#FFFFFF')
    $footerBorder.BorderBrush     = $bc.ConvertFrom('#E5E7EB')
    $footerBorder.BorderThickness = [System.Windows.Thickness]::new(0,1,0,0)
    $footerBorder.Padding         = [System.Windows.Thickness]::new(20,12,20,12)
    $footerBorder.Margin          = [System.Windows.Thickness]::new(0,10,0,0)
    [System.Windows.Controls.Grid]::SetRow($footerBorder, 1)

    $noteTb = [System.Windows.Controls.TextBlock]::new()
    $noteTb.Text         = "These are Graph API (OAuth) scopes only — all green means your token has the right permissions consented. Note: Intune remote actions (Sync, Fresh Start, etc.) also require the appropriate Intune RBAC role with 'Remote tasks' enabled, which is a separate check not shown here."
    $noteTb.FontSize     = 11
    $noteTb.Foreground   = $bc.ConvertFrom('#9CA3AF')
    $noteTb.TextWrapping = [System.Windows.TextWrapping]::Wrap
    $footerBorder.Child  = $noteTb
    $grid.Children.Add($footerBorder) | Out-Null

    # Make the outer stack stretch to fill the window
    $outerGrid = [System.Windows.Controls.Grid]::new()
    $outerGridRow0 = [System.Windows.Controls.RowDefinition]::new()
    $outerGridRow0.Height = [System.Windows.GridLength]::Auto
    $outerGridRow1 = [System.Windows.Controls.RowDefinition]::new()
    $outerGridRow1.Height = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    $outerGrid.RowDefinitions.Add($outerGridRow0)
    $outerGrid.RowDefinitions.Add($outerGridRow1)

    [System.Windows.Controls.Grid]::SetRow($headerBorder, 0)
    $outerGrid.Children.Add($headerBorder) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($grid, 1)
    $outerGrid.Children.Add($grid) | Out-Null

    $popupWindow.Content = $outerGrid
    $popupWindow.ShowDialog() | Out-Null
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

$ui['LapsCopyBtn'].Add_Click({
    if ($script:lapsPassword) {
        [System.Windows.Clipboard]::SetText($script:lapsPassword)
        Show-CopyConfirmation $ui['LapsCopyCheck']
    }
})

# ═══════════════════════════════════════════════════════════
#  DARK MODE TOGGLE
# ═══════════════════════════════════════════════════════════
$ui['DarkModeToggle'].Add_MouseLeftButtonDown({
    $script:isDarkMode = -not $script:isDarkMode
    Set-Theme -dark $script:isDarkMode
})

# ═══════════════════════════════════════════════════════════
#  ACTION BUTTON HOVER / PRESS FEEDBACK
# ═══════════════════════════════════════════════════════════
function Add-ButtonFeedback ([System.Windows.UIElement]$btn) {
    $btn.Add_MouseEnter({
        $bc = [System.Windows.Media.BrushConverter]::new()
        $this.Background = if ($script:isDarkMode) { $bc.ConvertFrom('#3A3B4E') } else { $bc.ConvertFrom('#E8E9F0') }
    })
    $btn.Add_MouseLeave({
        $this.Background = [System.Windows.Media.Brushes]::Transparent
    })
    $btn.Add_MouseLeftButtonDown({
        $bc = [System.Windows.Media.BrushConverter]::new()
        $this.Background = if ($script:isDarkMode) { $bc.ConvertFrom('#4A4B60') } else { $bc.ConvertFrom('#D5D7E0') }
    }.GetNewClosure())
    $btn.Add_MouseLeftButtonUp({
        $bc = [System.Windows.Media.BrushConverter]::new()
        $this.Background = if ($script:isDarkMode) { $bc.ConvertFrom('#3A3B4E') } else { $bc.ConvertFrom('#E8E9F0') }
    }.GetNewClosure())
}

foreach ($btnName in @('RefreshBtn','SyncDeviceBtn','FreshStartBtn','ExportPdfBtn')) {
    Add-ButtonFeedback $ui[$btnName]
}

# ═══════════════════════════════════════════════════════════
#  FRESH START (cleanWindowsDevice)
# ═══════════════════════════════════════════════════════════
$ui['FreshStartBtn'].Add_MouseLeftButtonDown({
    if (-not $script:connected -or -not $script:currentDeviceId) { return }
    $id = $script:currentDeviceId
    $deviceName = $ui['ValDeviceName'].Text

    # Confirm with user — custom dialog so we can label buttons Keep / Remove / Cancel
    $dlgResult  = $null
    $confirmDlg = [System.Windows.Window]::new()
    $confirmDlg.Title               = 'Confirm Fresh Start'
    $confirmDlg.Width               = 420
    $confirmDlg.SizeToContent       = [System.Windows.SizeToContent]::Height
    $confirmDlg.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner
    $confirmDlg.Owner               = $window
    $confirmDlg.ResizeMode          = [System.Windows.ResizeMode]::NoResize
    $confirmDlg.Background          = $window.Resources['ThSurface']
    $confirmDlg.FontFamily          = [System.Windows.Media.FontFamily]::new('Segoe UI')

    $bc = [System.Windows.Media.BrushConverter]::new()
    $dlgStack = [System.Windows.Controls.StackPanel]::new()
    $dlgStack.Margin = [System.Windows.Thickness]::new(24, 20, 24, 20)

    $dlgIcon = [System.Windows.Controls.TextBlock]::new()
    $dlgIcon.Text = '⚠'
    $dlgIcon.FontSize = 28
    $dlgIcon.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
    $dlgIcon.Margin = [System.Windows.Thickness]::new(0, 0, 0, 10)
    $dlgStack.Children.Add($dlgIcon) | Out-Null

    $dlgTitle = [System.Windows.Controls.TextBlock]::new()
    $dlgTitle.Text = "Fresh Start — $deviceName"
    $dlgTitle.FontSize = 14
    $dlgTitle.FontWeight = [System.Windows.FontWeights]::SemiBold
    $dlgTitle.TextWrapping = [System.Windows.TextWrapping]::Wrap
    $dlgTitle.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
    $dlgTitle.Margin = [System.Windows.Thickness]::new(0, 0, 0, 10)
    $dlgTitle.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThTxt1')
    $dlgStack.Children.Add($dlgTitle) | Out-Null

    $dlgBody = [System.Windows.Controls.TextBlock]::new()
    $dlgBody.Text = "Fresh Start reinstalls Windows and removes OEM pre-installed apps. The device stays enrolled in Entra ID / Intune.`n`nWhat should happen to the user's personal files?"
    $dlgBody.FontSize = 12
    $dlgBody.TextWrapping = [System.Windows.TextWrapping]::Wrap
    $dlgBody.Margin = [System.Windows.Thickness]::new(0, 0, 0, 20)
    $dlgBody.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ThTxt2')
    $dlgStack.Children.Add($dlgBody) | Out-Null

    $btnPanel = [System.Windows.Controls.StackPanel]::new()
    $btnPanel.Orientation = [System.Windows.Controls.Orientation]::Horizontal
    $btnPanel.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center

    $keepBtn = [System.Windows.Controls.Button]::new()
    $keepBtn.Content = 'Keep files'
    $keepBtn.Width = 110
    $keepBtn.Padding = [System.Windows.Thickness]::new(0, 8, 0, 8)
    $keepBtn.Margin = [System.Windows.Thickness]::new(0, 0, 8, 0)
    $keepBtn.Background = $bc.ConvertFrom('#5B5FC7')
    $keepBtn.Foreground = [System.Windows.Media.Brushes]::White
    $keepBtn.BorderThickness = [System.Windows.Thickness]::new(0)
    $keepBtn.FontSize = 13
    $keepBtn.FontWeight = [System.Windows.FontWeights]::SemiBold
    $keepBtn.Cursor = [System.Windows.Input.Cursors]::Hand
    $keepBtn.Add_Click({ $script:dlgResult = 'Keep'; $confirmDlg.Close() })
    $btnPanel.Children.Add($keepBtn) | Out-Null

    $removeBtn = [System.Windows.Controls.Button]::new()
    $removeBtn.Content = 'Remove files'
    $removeBtn.Width = 110
    $removeBtn.Padding = [System.Windows.Thickness]::new(0, 8, 0, 8)
    $removeBtn.Margin = [System.Windows.Thickness]::new(0, 0, 8, 0)
    $removeBtn.Background = $bc.ConvertFrom('#DC2626')
    $removeBtn.Foreground = [System.Windows.Media.Brushes]::White
    $removeBtn.BorderThickness = [System.Windows.Thickness]::new(0)
    $removeBtn.FontSize = 13
    $removeBtn.FontWeight = [System.Windows.FontWeights]::SemiBold
    $removeBtn.Cursor = [System.Windows.Input.Cursors]::Hand
    $removeBtn.Add_Click({ $script:dlgResult = 'Remove'; $confirmDlg.Close() })
    $btnPanel.Children.Add($removeBtn) | Out-Null

    $cancelBtn = [System.Windows.Controls.Button]::new()
    $cancelBtn.Content = 'Cancel'
    $cancelBtn.Width = 80
    $cancelBtn.Padding = [System.Windows.Thickness]::new(0, 8, 0, 8)
    $cancelBtn.BorderThickness = [System.Windows.Thickness]::new(1)
    $cancelBtn.FontSize = 13
    $cancelBtn.Cursor = [System.Windows.Input.Cursors]::Hand
    $cancelBtn.Add_Click({ $script:dlgResult = 'Cancel'; $confirmDlg.Close() })
    $btnPanel.Children.Add($cancelBtn) | Out-Null

    $dlgStack.Children.Add($btnPanel) | Out-Null
    $confirmDlg.Content = $dlgStack

    $script:dlgResult = 'Cancel'
    $confirmDlg.ShowDialog() | Out-Null
    if ($script:dlgResult -eq 'Cancel') { return }
    $keepUserData = ($script:dlgResult -eq 'Keep')

    $sp      = $ui['FreshStartBtn'].Child
    $labelTb = $sp.Children[1]

    $labelTb.Text = 'Sending…'
    $ui['FreshStartBtn'].IsEnabled = $false
    Push-UI

    $ok = $false
    $errDetail = ''
    try {
        $body = @{ keepUserData = $keepUserData } | ConvertTo-Json
        Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$id/cleanWindowsDevice" `
            -Method POST -Body $body -ContentType 'application/json' -ErrorAction Stop | Out-Null
        $ok = $true
    } catch {
        $errDetail = $_.Exception.Message
        $rawBody = ''
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $rawBody = $_.ErrorDetails.Message }
        elseif ($_.Exception.Response) {
            try {
                $stream = $_.Exception.Response.Content.ReadAsStringAsync().Result
                $rawBody = $stream
            } catch {}
        }
        if ($rawBody) {
            try {
                $parsed = $rawBody | ConvertFrom-Json -ErrorAction SilentlyContinue
                if ($parsed.error.message) { $errDetail = $parsed.error.message }
                elseif ($parsed.error.code) { $errDetail = "$($parsed.error.code): $errDetail" }
            } catch {}
        }
        # Append diagnostics for authentication / authorisation failures
        if ($errDetail -match '401|Unauthorized') {
            $errDetail += "`n`nYour session may have expired. Please sign out and sign back in."
        } elseif ($errDetail -match '403|Forbidden|Authorization') {
            try {
                $ctx = Get-MgContext -ErrorAction SilentlyContinue
                $relevantScopes = $ctx.Scopes | Where-Object { $_ -match 'DeviceManagement' }
                $errDetail += "`n`nToken scopes (DeviceManagement):`n" + ($relevantScopes -join "`n")
                $errDetail += "`n`nIf 'PrivilegedOperations.All' is missing, sign out and sign back in.`nIf it is present, an Intune RBAC role with 'Remote tasks' permission is required."
            } catch {}
        }
    }

    if ($ok) {
        $labelTb.Text = 'Sent ✓'
        $ui['FreshStartBtn'].IsEnabled = $true
        $t = [System.Windows.Threading.DispatcherTimer]::new()
        $t.Interval = [TimeSpan]::FromSeconds(3)
        $t.Tag = $labelTb
        $t.Add_Tick({ param($s,$e) $s.Tag.Text = 'Fresh Start'; $s.Stop() })
        $t.Start()
    } else {
        $labelTb.Text = 'Error'
        $ui['FreshStartBtn'].IsEnabled = $true
        [System.Windows.MessageBox]::Show(
            $errDetail,
            'Fresh Start Failed',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        ) | Out-Null
        $t = [System.Windows.Threading.DispatcherTimer]::new()
        $t.Interval = [TimeSpan]::FromSeconds(4)
        $t.Tag = $labelTb
        $t.Add_Tick({ param($s,$e) $s.Tag.Text = 'Fresh Start'; $s.Stop() })
        $t.Start()
    }
})

# ═══════════════════════════════════════════════════════════
#  CHANGE AUTOPILOT GROUP TAG
# ═══════════════════════════════════════════════════════════
$ui['ChangeGroupTagBtn'].Add_MouseLeftButtonDown({
    if (-not $script:connected -or -not $script:currentAutopilotId) {
        [System.Windows.MessageBox]::Show(
            "No Autopilot identity found for this device.`nThe device may not be Autopilot-registered.",
            'Cannot Change Group Tag', 'OK', 'Warning') | Out-Null
        return
    }

    $currentTag = $ui['ValGroupTag'].Text

    # ── Fetch all distinct group tags from the tenant ──
    Show-Overlay 'Loading group tags…'
    $allTags = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    try {
        $uri = "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?`$top=100"
        while ($uri) {
            $result = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop
            foreach ($item in $result.value) {
                if ($item.groupTag -and $item.groupTag.Trim().Length -gt 0) {
                    [void]$allTags.Add($item.groupTag.Trim())
                }
            }
            $uri = $result.'@odata.nextLink'
        }
    } catch {
        Hide-Overlay
        [System.Windows.MessageBox]::Show(
            "Failed to retrieve group tags:`n$($_.Exception.Message)",
            'Error', 'OK', 'Error') | Out-Null
        return
    }
    Hide-Overlay

    if ($allTags.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            'No group tags found in the tenant.',
            'No Group Tags', 'OK', 'Information') | Out-Null
        return
    }

    $sortedTags = $allTags | Sort-Object

    # ── Build selection dialog ──
    $dlg = [System.Windows.Window]::new()
    $dlg.Title  = 'Select Group Tag'
    $dlg.Width  = 380
    $dlg.Height = 480
    $dlg.WindowStartupLocation = 'CenterOwner'
    $dlg.Owner  = $window
    $dlg.ResizeMode = 'NoResize'
    $dlg.Background = $window.Resources['ThSurface']

    $stack = [System.Windows.Controls.StackPanel]::new()
    $stack.Margin = [System.Windows.Thickness]::new(20)

    $header = [System.Windows.Controls.TextBlock]::new()
    $header.Text = "Current tag: $currentTag"
    $header.FontSize = 13
    $header.Foreground = $window.Resources['ThTxt3']
    $header.Margin = [System.Windows.Thickness]::new(0, 0, 0, 12)
    $stack.Children.Add($header) | Out-Null

    $label = [System.Windows.Controls.TextBlock]::new()
    $label.Text = 'Select a new group tag:'
    $label.FontSize = 13
    $label.FontWeight = [System.Windows.FontWeights]::Medium
    $label.Foreground = $window.Resources['ThTxt1']
    $label.Margin = [System.Windows.Thickness]::new(0, 0, 0, 8)
    $stack.Children.Add($label) | Out-Null

    $listBox = [System.Windows.Controls.ListBox]::new()
    $listBox.Height = 270
    $listBox.FontSize = 13
    $listBox.BorderBrush = $window.Resources['ThBorderSub']
    $listBox.Background = $window.Resources['ThInputBg']
    foreach ($tag in $sortedTags) {
        $listBox.Items.Add($tag) | Out-Null
    }
    # Pre-select the current tag if it exists
    $idx = [System.Array]::IndexOf($sortedTags, $currentTag)
    if ($idx -ge 0) { $listBox.SelectedIndex = $idx }
    $stack.Children.Add($listBox) | Out-Null

    $btnPanel = [System.Windows.Controls.StackPanel]::new()
    $btnPanel.Orientation = 'Horizontal'
    $btnPanel.HorizontalAlignment = 'Right'
    $btnPanel.Margin = [System.Windows.Thickness]::new(0, 14, 0, 0)

    $cancelBtn = [System.Windows.Controls.Button]::new()
    $cancelBtn.Content = 'Cancel'
    $cancelBtn.Width = 80
    $cancelBtn.Padding = [System.Windows.Thickness]::new(0, 6, 0, 6)
    $cancelBtn.Margin = [System.Windows.Thickness]::new(0, 0, 8, 0)
    $cancelBtn.Add_Click({ $dlg.DialogResult = $false; $dlg.Close() })
    $btnPanel.Children.Add($cancelBtn) | Out-Null

    $okBtn = [System.Windows.Controls.Button]::new()
    $okBtn.Content = 'OK'
    $okBtn.Width = 80
    $okBtn.Padding = [System.Windows.Thickness]::new(0, 6, 0, 6)
    $okBtn.IsDefault = $true
    $okBtn.Add_Click({ $dlg.DialogResult = $true; $dlg.Close() })
    $btnPanel.Children.Add($okBtn) | Out-Null

    $stack.Children.Add($btnPanel) | Out-Null
    $dlg.Content = $stack

    $dlgResult = $dlg.ShowDialog()
    if (-not $dlgResult -or $null -eq $listBox.SelectedItem) { return }

    $newTag = $listBox.SelectedItem.ToString()
    if ($newTag -eq $currentTag) { return }

    # ── Confirmation dialog ──
    $confirm = [System.Windows.MessageBox]::Show(
        "Change group tag from '$currentTag' to '$newTag'?`n`nDevice: $($ui['ValDeviceName'].Text)",
        'Confirm Group Tag Change',
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    if ($confirm -ne 'Yes') { return }

    # ── Apply the change via Graph API ──
    Show-Overlay 'Updating group tag…'
    $apId = $script:currentAutopilotId
    try {
        $body = @{ groupTag = $newTag } | ConvertTo-Json
        Invoke-MgGraphRequest `
            -Uri "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities/$apId/updateDeviceProperties" `
            -Method POST -Body $body -ContentType 'application/json' -OutputType HttpResponseMessage -ErrorAction Stop | Out-Null
        Hide-Overlay
        $ui['ValGroupTag'].Text = $newTag
        # Show "Pending sync" badge
        $bcTag = [System.Windows.Media.BrushConverter]::new()
        $ui['GroupTagStatus'].Background      = $bcTag.ConvertFrom('#FEF3C7')
        $ui['GroupTagStatusText'].Text        = 'Pending sync'
        $ui['GroupTagStatusText'].Foreground  = $bcTag.ConvertFrom('#D97706')
        $ui['GroupTagStatus'].Visibility      = 'Visible'
        [System.Windows.MessageBox]::Show(
            "Group tag updated to '$newTag'.`nThe device may need to sync for the change to fully apply.",
            'Group Tag Updated', 'OK', 'Information') | Out-Null
    } catch {
        Hide-Overlay
        $errMsg = $_.Exception.Message
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            try {
                $parsed = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
                if ($parsed.error.message) { $errMsg = $parsed.error.message }
            } catch {}
        }
        [System.Windows.MessageBox]::Show(
            "Failed to update group tag:`n$errMsg",
            'Update Failed', 'OK', 'Error') | Out-Null
    }
})

# ═══════════════════════════════════════════════════════════
#  EXPORT PDF
# ═══════════════════════════════════════════════════════════
function Export-DevicePdf {
    if (-not $script:currentDeviceId) { return }

    $deviceName = $ui['ValDeviceName'].Text
    $safeName   = $deviceName -replace '[\\/:*?"<>|]', '_'
    $desktop    = [Environment]::GetFolderPath('Desktop')
    $timestamp  = (Get-Date).ToString('yyyy-MM-dd_HHmm')
    $pdfPath    = Join-Path $desktop "${safeName}_Report_${timestamp}.pdf"
    $htmlPath   = Join-Path ([System.IO.Path]::GetTempPath()) "intune_report_${timestamp}.html"

    # Gather health card data from UI
    $healthScore = $ui['HealthScoreText'].Text
    $healthLabel = $ui['HealthScoreLabel'].Text
    $healthBg    = ($ui['HealthScoreBadge'].Background).ToString() -replace '^#FF','#'

    # Map icon colours → friendly CSS colour (WPF returns #AARRGGBB, so strip alpha)
    function Get-HealthColor($color) {
        $c = "$color" -replace '^#FF','#'
        switch -Regex ($c) {
            '#16A34A|#10B981|#059669|green'  { return @{ css='#16A34A'; bg='#ECFDF5' } }
            '#DC2626|#EF4444|red'            { return @{ css='#DC2626'; bg='#FEF2F2' } }
            '#EA580C|#F59E0B|#D97706|orange' { return @{ css='#EA580C'; bg='#FFF7ED' } }
            default                          { return @{ css='#9CA3AF'; bg='#F3F4F6' } }
        }
    }
    # Category-specific SVG icons (each 16x16, stroke-based)
    $categoryIcons = @{
        'Compliance'  = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><polyline points="9 12 11 14 15 10"/></svg>'
        'Encryption'  = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="11" width="18" height="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0110 0v4"/><circle cx="12" cy="16" r="1"/></svg>'
        'Last Sync'   = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/><path d="M3.51 9a9 9 0 0114.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0020.49 15"/></svg>'
        'Management'  = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-4 0v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83-2.83l.06-.06A1.65 1.65 0 004.68 15a1.65 1.65 0 00-1.51-1H3a2 2 0 010-4h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 012.83-2.83l.06.06A1.65 1.65 0 009 4.68a1.65 1.65 0 001-1.51V3a2 2 0 014 0v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 2.83l-.06.06A1.65 1.65 0 0019.4 9a1.65 1.65 0 001.51 1H21a2 2 0 010 4h-.09a1.65 1.65 0 00-1.51 1z"/></svg>'
        'Defender/AV' = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>'
    }
    $hthItems = @(
        @{ Label='Compliance';  Value=$ui['HthCompliantVal'].Text; Color=$ui['HthCompliantIcon'].Foreground.ToString() }
        @{ Label='Encryption';  Value=$ui['HthEncryptedVal'].Text; Color=$ui['HthEncryptedIcon'].Foreground.ToString() }
        @{ Label='Last Sync';   Value=$ui['HthSyncVal'].Text;      Color=$ui['HthSyncIcon'].Foreground.ToString() }
        @{ Label='Management';  Value=$ui['HthMgmtVal'].Text;      Color=$ui['HthMgmtIcon'].Foreground.ToString() }
        @{ Label='Defender/AV'; Value=$ui['HthDefenderVal'].Text;   Color=$ui['HthDefenderIcon'].Foreground.ToString() }
    )
    $healthItemsHtml = ($hthItems | ForEach-Object {
        $hc = Get-HealthColor $_.Color
        $iconSvg = $categoryIcons[$_.Label]
        @"
<div class="hth-item">
  <div class="hth-icon-circle" style="background:$($hc.bg);color:$($hc.css)">$iconSvg</div>
  <div class="hth-info">
    <div class="hth-label">$($_.Label)</div>
    <div class="hth-value" style="color:$($hc.css)">$([System.Net.WebUtility]::HtmlEncode($_.Value))</div>
  </div>
</div>
"@
    }) -join "`n"

    # Device properties
    $props = @(
        @{ Label='Device Name';          Value=$ui['ValDeviceName'].Text }
        @{ Label='OS Version';           Value=$ui['ValOSVersion'].Text }
        @{ Label='OS Build';             Value=$ui['ValOSBuild'].Text }
        @{ Label='Last Logged On User';  Value=$ui['ValLastLogon'].Text }
        @{ Label='Primary User';         Value=$ui['ValPrimaryUser'].Text }
        @{ Label='All Logged-In Users';  Value=$ui['ValAllUsers'].Text }
        @{ Label='Installed Enrollment Profile'; Value=$ui['ValEnrollProfile'].Text }
        @{ Label='Assigned Enrollment Profile';  Value=$ui['ValAssignedEnrollProfile'].Text }
        @{ Label='Autopilot Group Tag';          Value=$ui['ValGroupTag'].Text }
        @{ Label='Last Enrolled';        Value=$ui['ValLastEnrolled'].Text }
        @{ Label='Last Password Change'; Value=$ui['ValLastPasswordChange'].Text }
        @{ Label='BIOS / UEFI Version';  Value=$ui['ValBiosVersion'].Text }
        @{ Label='Defender / AV';        Value=$ui['ValDefenderStatus'].Text }
    )
    $propRows = ($props | ForEach-Object {
        "<tr><td class='prop-label'>$($_.Label)</td><td class='prop-value'>$([System.Net.WebUtility]::HtmlEncode($_.Value))</td></tr>"
    }) -join "`n"

    # Applications (if loaded)
    $appsHtml = ''
    if ($script:appsLoaded -and $script:allFilteredApps.Count -gt 0) {
        $appRows = ($script:allFilteredApps | ForEach-Object {
            "<tr><td>$([System.Net.WebUtility]::HtmlEncode($_.Name))</td>" +
            "<td style='color:#6B7280'>$([System.Net.WebUtility]::HtmlEncode($_.Publisher))</td>" +
            "<td style='color:#6B7280;text-align:right'>$([System.Net.WebUtility]::HtmlEncode($_.Version))</td></tr>"
        }) -join "`n"
        $appsHtml = @"
<div class="section-title">Applications ($($script:allFilteredApps.Count))</div>
<table class="apps-table">
<thead><tr><th>Name</th><th>Publisher</th><th style="text-align:right">Version</th></tr></thead>
<tbody>
$appRows
</tbody>
</table>
"@
    }

    # Compliance (if loaded)
    $compHtml = ''
    if ($script:complianceLoaded) {
        $compState = $ui['DeviceComplianceText'].Text
        $compHtml = "<div class=`"section-title`">Compliance — $([System.Net.WebUtility]::HtmlEncode($compState))</div>`n"
        $compHtml += "<p style='font-size:11px;color:#6B7280;margin-bottom:10px'>$([System.Net.WebUtility]::HtmlEncode($ui['ComplianceStatus'].Text))</p>"
    }

    $subtitle   = $ui['DetailSubtitle'].Text
    $exportDate = (Get-Date).ToString('yyyy-MM-dd HH:mm')

    $html = @"
<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>Device Report - $([System.Net.WebUtility]::HtmlEncode($deviceName))</title>
<style>
  @page { size: A4; margin: 10mm 12mm; }
  * { margin:0; padding:0; box-sizing:border-box; }
  body { font-family:'Segoe UI',system-ui,-apple-system,sans-serif; color:#1A1B25; font-size:11px; line-height:1.4; background:#fff; }

  /* ── Header ── */
  .header { display:flex; align-items:center; gap:16px; margin-bottom:16px; padding-bottom:12px; border-bottom:2px solid #E5E7EB; }
  .header-left { flex:1; }
  .header-left h1 { font-size:22px; font-weight:700; color:#111827; margin-bottom:2px; letter-spacing:-0.5px; }
  .header-left .subtitle { font-size:11px; color:#6B7280; }
  .header-left .export-date { font-size:9px; color:#9CA3AF; margin-top:4px; }

  /* ── Health Card ── */
  .health-card {
    background: linear-gradient(135deg, #FAFBFC 0%, #F3F4F6 100%);
    border:1px solid #E5E7EB; border-radius:14px;
    padding:14px 18px; width:100%; margin-bottom:16px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    page-break-inside: avoid;
  }
  .hc-top { display:flex; align-items:center; margin-bottom:12px; }
  .hc-title-block { flex:1; }
  .hc-title { font-size:11px; font-weight:700; color:#6B7280; letter-spacing:1px; text-transform:uppercase; }
  .hc-subtitle { font-size:10px; color:#9CA3AF; margin-top:2px; }

  /* Circular score badge */
  .score-ring {
    width:60px; height:60px; border-radius:50%;
    background: ${healthBg};
    display:flex; flex-direction:column; align-items:center; justify-content:center;
    color:white; box-shadow: 0 4px 12px rgba(0,0,0,0.12);
  }
  .score-ring .score { font-size:20px; font-weight:800; line-height:1; }
  .score-ring .label { font-size:8px; font-weight:700; text-transform:uppercase; letter-spacing:0.5px; opacity:0.95; margin-top:2px; }

  /* Health items grid */
  .hth-grid { display:grid; grid-template-columns:1fr 1fr 1fr; gap:8px 16px; }
  .hth-item {
    display:flex; align-items:center; gap:8px;
    background:#fff; border:1px solid #E5E7EB; border-radius:8px;
    padding:6px 10px;
  }
  .hth-icon-circle {
    width:26px; height:26px; border-radius:50%;
    display:flex; align-items:center; justify-content:center;
    flex-shrink:0;
  }
  .hth-info { display:flex; flex-direction:column; min-width:0; }
  .hth-label { font-size:9px; font-weight:600; color:#6B7280; text-transform:uppercase; letter-spacing:0.3px; }
  .hth-value { font-size:11px; font-weight:600; margin-top:1px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }

  /* ── Sections ── */
  .section-title {
    font-size:13px; font-weight:700; color:#111827;
    margin:16px 0 8px 0; padding-bottom:6px;
    border-bottom:2px solid #E5E7EB;
    letter-spacing:-0.2px;
  }

  /* ── Properties table ── */
  .props-table { width:100%; border-collapse:collapse; }
  .props-table tr:nth-child(even) { background:#FAFBFC; }
  .prop-label { padding:6px 12px 6px 8px; font-weight:600; color:#6B7280; width:180px; vertical-align:top; font-size:10px; text-transform:uppercase; letter-spacing:0.3px; }
  .prop-value { padding:6px 8px 6px 0; color:#111827; font-size:11px; }

  /* ── Apps table ── */
  .apps-table { width:100%; border-collapse:collapse; border:1px solid #E5E7EB; border-radius:8px; overflow:hidden; }
  .apps-table th { background:#F9FAFB; text-align:left; padding:10px 14px; color:#6B7280; font-weight:700; font-size:10px; text-transform:uppercase; letter-spacing:0.5px; border-bottom:2px solid #E5E7EB; }
  .apps-table td { padding:8px 14px; border-bottom:1px solid #F3F4F6; font-size:11px; }
  .apps-table tr:nth-child(even) { background:#FAFBFC; }
  .apps-table tr:hover { background:#F3F4F6; }

  /* ── Footer ── */
  .footer { margin-top:20px; padding-top:10px; border-top:2px solid #E5E7EB; text-align:center; font-size:9px; color:#9CA3AF; }
</style></head><body>

<!-- Header -->
<div class="header">
  <div class="header-left">
    <h1>$([System.Net.WebUtility]::HtmlEncode($deviceName))</h1>
    <div class="subtitle">$([System.Net.WebUtility]::HtmlEncode($subtitle))</div>
    <div class="export-date">Exported $exportDate</div>
  </div>
</div>

<!-- Health Card -->
<div class="health-card">
  <div class="hc-top">
    <div class="hc-title-block">
      <div class="hc-title">Device Health</div>
      <div class="hc-subtitle">Real-time status overview</div>
    </div>
    <div class="score-ring">
      <div class="score">$healthScore</div>
      <div class="label">$healthLabel</div>
    </div>
  </div>
  <div class="hth-grid">
    $healthItemsHtml
  </div>
</div>

<!-- Device Details -->
<div class="section-title">Device Details</div>
<table class="props-table">$propRows</table>

$compHtml

$appsHtml

<div class="footer">Intune Device Lookup &middot; Jacob D&#252;ring Bakkeli &middot; $([datetime]::Now.Year)</div>

</body></html>
"@

    [System.IO.File]::WriteAllText($htmlPath, $html, [System.Text.Encoding]::UTF8)

    # Find Edge or Chrome for headless PDF conversion
    $edgePaths = @(
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
    )
    $chromePaths = @(
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
    )
    $browser = ($edgePaths + $chromePaths) | Where-Object { Test-Path $_ } | Select-Object -First 1

    if (-not $browser) {
        [System.Windows.MessageBox]::Show(
            "Could not find Edge or Chrome to generate PDF.`nHTML report saved to:`n$htmlPath",
            "PDF Export", "OK", "Warning") | Out-Null
        Start-Process $htmlPath
        return
    }

    Show-PdfProgress 'Preparing report data...' 10
    Sleep-UI 300

    try {
        Show-PdfProgress 'Building HTML template...' 30
        Sleep-UI 300

        Show-PdfProgress 'Launching browser renderer...' 50
        Sleep-UI 200
        $fileUri = ([System.Uri]::new($htmlPath)).AbsoluteUri
        $argString = "--headless=new --disable-gpu --no-sandbox --disable-extensions --run-all-compositor-stages-before-draw `"--print-to-pdf=$pdfPath`" --no-pdf-header-footer `"$fileUri`""
        $psi = [System.Diagnostics.ProcessStartInfo]::new()
        $psi.FileName  = $browser
        $psi.Arguments = $argString
        $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow  = $true

        Show-PdfProgress 'Converting to PDF...' 65
        $proc = [System.Diagnostics.Process]::Start($psi)

        # Poll instead of blocking WaitForExit so WPF stays responsive
        $timeout = [datetime]::UtcNow.AddSeconds(30)
        while (-not $proc.HasExited -and [datetime]::UtcNow -lt $timeout) {
            Push-UI
            Start-Sleep -Milliseconds 80
        }
        if (-not $proc.HasExited) { $proc.Kill() }

        # Brief wait for file system flush
        Sleep-UI 600

        Show-PdfProgress 'Finalising...' 90
        Sleep-UI 400

        # Retry check — file may appear after a short delay
        $pdfExists = $false
        for ($i = 0; $i -lt 5; $i++) {
            if (Test-Path $pdfPath) { $pdfExists = $true; break }
            Sleep-UI 300
        }

        if ($pdfExists) {
            Show-PdfProgress 'Done!' 100
            Sleep-UI 400
            Hide-Overlay
            [System.Windows.MessageBox]::Show(
                "PDF saved to desktop:`n$([System.IO.Path]::GetFileName($pdfPath))",
                "Export Complete", "OK", "Information") | Out-Null
            Start-Process $pdfPath
        } else {
            Hide-Overlay
            [System.Windows.MessageBox]::Show(
                "PDF generation may have failed.`nHTML report saved to:`n$htmlPath",
                "PDF Export", "OK", "Warning") | Out-Null
        }
    } catch {
        Hide-Overlay
        [System.Windows.MessageBox]::Show(
            "Error generating PDF: $($_.Exception.Message)",
            "PDF Export Error", "OK", "Error") | Out-Null
    } finally {
        Remove-Item $htmlPath -ErrorAction SilentlyContinue
    }
}

$ui['ExportPdfBtn'].Add_MouseLeftButtonDown({ Export-DevicePdf })

# ═══════════════════════════════════════════════════════════
#  LAUNCH
# ═══════════════════════════════════════════════════════════
$window.ShowDialog() | Out-Null

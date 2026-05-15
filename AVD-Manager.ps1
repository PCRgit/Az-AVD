#Requires -Version 5.1
<#
.SYNOPSIS
    Azure Virtual Desktop Manager — Universal Deployment & Management Console

.DESCRIPTION
    Complete AVD lifecycle tool supporting all deployment types:
      * Pooled Desktop (BreadthFirst / DepthFirst)
      * Personal Desktop (Automatic / Direct assignment)
      * RemoteApp (published application streaming)

    Features:
      * Connect to any Azure tenant / subscription
      * License assessment via Microsoft Graph (detects AVD entitlement)
      * Step-by-step deployment wizard for all host pool types
      * Domain join wizard (Entra ID Join or Hybrid AD Join)
      * Session host management (drain, heal, scale)
      * FSLogix profile configuration
      * Auto-scaling plan deployment
      * App group / RemoteApp publishing
      * Monitoring and Log Analytics integration
      * RDP property editor with security presets
      * Cost estimator

.NOTES
    Author  : IT Infrastructure
    Version : 1.0
    Run as  : powershell.exe -STA -File AVD-Manager.ps1
    Requires: Az PowerShell 11+, Windows OS (WPF)
#>

# -- STA guard ----------------------------------------------------------------
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne "STA") {
    $s = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
    Start-Process powershell.exe -ArgumentList "-STA -NoProfile -ExecutionPolicy Bypass -WorkingDirectory `"$(Split-Path $s -Parent)`" -File `"$s`""
    return
}
Set-StrictMode -Off
$ErrorActionPreference = "Continue"

# -- Assemblies ---------------------------------------------------------------
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName Microsoft.VisualBasic, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# -- License SKU Map ----------------------------------------------------------
$Script:AvdSkuMap = [ordered]@{
    "SPE_E3"                          = @{Name="Microsoft 365 E3";                    AVD=$true;  LicType="Desktop+RemoteApp"; Notes="Windows multi-session included"}
    "SPE_E5"                          = @{Name="Microsoft 365 E5";                    AVD=$true;  LicType="Desktop+RemoteApp"; Notes="Windows multi-session included"}
    "O365_BUSINESS_PREMIUM"           = @{Name="Microsoft 365 Business Premium";      AVD=$true;  LicType="Desktop+RemoteApp"; Notes="Up to 300 users"}
    "SPE_F1"                          = @{Name="Microsoft 365 F1";                    AVD=$true;  LicType="RemoteApp only";    Notes="No full desktop streaming"}
    "SPE_F3"                          = @{Name="Microsoft 365 F3";                    AVD=$true;  LicType="Desktop+RemoteApp"; Notes="Windows multi-session included"}
    "WIN10_VDA_E3"                    = @{Name="Windows 10/11 Enterprise E3";         AVD=$true;  LicType="Desktop+RemoteApp"; Notes="Per-user, includes RDS SAL"}
    "WIN10_VDA_E5"                    = @{Name="Windows 10/11 Enterprise E5";         AVD=$true;  LicType="Desktop+RemoteApp"; Notes="Per-user, includes RDS SAL"}
    "Microsoft_Azure_Virtual_Desktop" = @{Name="Azure Virtual Desktop";               AVD=$true;  LicType="Desktop+RemoteApp"; Notes="Standalone AVD entitlement"}
    "M365EDU_A3_FACULTY"              = @{Name="Microsoft 365 A3 Faculty";            AVD=$true;  LicType="Desktop+RemoteApp"; Notes="Education"}
    "M365EDU_A5_FACULTY"              = @{Name="Microsoft 365 A5 Faculty";            AVD=$true;  LicType="Desktop+RemoteApp"; Notes="Education"}
    "AAD_PREMIUM"                     = @{Name="Entra ID P1";                         AVD=$false; LicType="Supporting";        Notes="Required: Conditional Access, MFA"}
    "AAD_PREMIUM_P2"                  = @{Name="Entra ID P2";                         AVD=$false; LicType="Supporting";        Notes="PIM, Identity Protection"}
    "INTUNE_A"                        = @{Name="Microsoft Intune Plan 1";             AVD=$false; LicType="Supporting";        Notes="Required for Entra ID Join mgmt"}
    "EMS"                             = @{Name="Enterprise Mobility + Security E3";   AVD=$false; LicType="Supporting";        Notes="Entra P1 + Intune bundle"}
    "EMSPREMIUM"                      = @{Name="Enterprise Mobility + Security E5";   AVD=$false; LicType="Supporting";        Notes="Entra P2 + Intune bundle"}
}

# -- VM Sizes -----------------------------------------------------------------
$Script:VmSizes = @(
    [PSCustomObject]@{Series="D-Series";   Size="Standard_D2s_v5";          vCPU=2;  RAM="8 GB";   MaxSessions=2;  UseCase="Light/Test"}
    [PSCustomObject]@{Series="D-Series";   Size="Standard_D4s_v5";          vCPU=4;  RAM="16 GB";  MaxSessions=6;  UseCase="Office, Teams light"}
    [PSCustomObject]@{Series="D-Series";   Size="Standard_D8s_v5";          vCPU=8;  RAM="32 GB";  MaxSessions=12; UseCase="Office, Teams moderate"}
    [PSCustomObject]@{Series="D-Series";   Size="Standard_D16s_v5";         vCPU=16; RAM="64 GB";  MaxSessions=20; UseCase="Power users"}
    [PSCustomObject]@{Series="D-Series";   Size="Standard_D4ds_v5";         vCPU=4;  RAM="16 GB";  MaxSessions=6;  UseCase="Office + local temp disk"}
    [PSCustomObject]@{Series="D-Series";   Size="Standard_D8ds_v5";         vCPU=8;  RAM="32 GB";  MaxSessions=12; UseCase="Moderate + local temp disk"}
    [PSCustomObject]@{Series="E-Series";   Size="Standard_E4s_v5";          vCPU=4;  RAM="32 GB";  MaxSessions=8;  UseCase="Memory-intensive apps"}
    [PSCustomObject]@{Series="E-Series";   Size="Standard_E8s_v5";          vCPU=8;  RAM="64 GB";  MaxSessions=15; UseCase="CAD, large datasets"}
    [PSCustomObject]@{Series="E-Series";   Size="Standard_E16s_v5";         vCPU=16; RAM="128 GB"; MaxSessions=25; UseCase="Heavy memory workloads"}
    [PSCustomObject]@{Series="NV-Series";  Size="Standard_NV6ads_A10_v5";   vCPU=6;  RAM="55 GB";  MaxSessions=2;  UseCase="GPU / 3D graphics"}
    [PSCustomObject]@{Series="NV-Series";  Size="Standard_NV12ads_A10_v5";  vCPU=12; RAM="110 GB"; MaxSessions=4;  UseCase="GPU moderate"}
    [PSCustomObject]@{Series="NV-Series";  Size="Standard_NV18ads_A10_v5";  vCPU=18; RAM="220 GB"; MaxSessions=6;  UseCase="GPU heavy"}
    [PSCustomObject]@{Series="B-Series";   Size="Standard_B4ms";            vCPU=4;  RAM="16 GB";  MaxSessions=4;  UseCase="Dev/test, low cost"}
    [PSCustomObject]@{Series="B-Series";   Size="Standard_B8ms";            vCPU=8;  RAM="32 GB";  MaxSessions=8;  UseCase="Dev/test, moderate"}
)

# -- Marketplace Images -------------------------------------------------------
$Script:MktImages = @(
    "Windows 11 Enterprise multi-session + Microsoft 365 Apps (latest) [RECOMMENDED]"
    "Windows 11 Enterprise multi-session (latest)"
    "Windows 10 Enterprise multi-session + Microsoft 365 Apps (latest)"
    "Windows 10 Enterprise multi-session (latest)"
    "Windows Server 2022 Datacenter Azure Edition"
    "Windows Server 2019 Datacenter Azure Edition"
)

# -- Azure Regions ------------------------------------------------------------
$Script:AzureRegions = @(
    "eastus","eastus2","centralus","northcentralus","southcentralus","westcentralus",
    "westus","westus2","westus3","canadacentral","canadaeast",
    "northeurope","westeurope","uksouth","ukwest","francecentral","germanywestcentral",
    "switzerlandnorth","norwayeast","swedencentral",
    "australiaeast","australiasoutheast","japaneast","japanwest",
    "southeastasia","eastasia","koreacentral","centralindia","southindia"
)

# -- Global State -------------------------------------------------------------
$Global:AppVersion   = "1.0"
$Global:ConfigFile   = ".\AVD-Manager-Config.json"
$Global:LogFile      = ".\AVD-Manager-$(Get-Date -f 'yyyyMMdd').log"
$Global:IsConnected  = $false
$Global:Subscription = $null
$Global:TenantId     = $null
$Global:RefreshSecs  = 60

$Global:Sync = [hashtable]::Synchronized(@{
    LogQueue    = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    IsDeploying = $false
    Progress    = 0
    CancelToken = $false
    StatusMsg   = "Ready"
})

$Global:Cfg = @{ LastSubscription=""; LastTenant=""; LastLocation="eastus"; RefreshSecs=60 }
if (Test-Path $Global:ConfigFile) {
    try {
        $lc = Get-Content $Global:ConfigFile -Raw | ConvertFrom-Json -EA Stop
        "LastSubscription","LastTenant","LastLocation","RefreshSecs" | ForEach-Object {
            if ($lc.$_) { $Global:Cfg[$_] = $lc.$_ }
        }
    } catch {}
}

# -- Logging ------------------------------------------------------------------
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = Get-Date -f "HH:mm:ss"
    $Global:Sync.LogQueue.Enqueue(@{Timestamp=$ts; Level=$Level; Message=$Message})
    Add-Content $Global:LogFile "[$ts][$Level] $Message" -EA SilentlyContinue
}

function Flush-UILog {
    $item = $null
    while ($Global:Sync.LogQueue.TryDequeue([ref]$item)) {
        try {
            $lb = $Global:LogBox
            if (-not ($lb -and $lb.Document)) { continue }
            $color = switch ($item.Level) {
                "OK"     { [System.Windows.Media.Brushes]::LightGreen      }
                "ERROR"  { [System.Windows.Media.Brushes]::IndianRed       }
                "WARN"   { [System.Windows.Media.Brushes]::Gold            }
                "DEPLOY" { [System.Windows.Media.Brushes]::CornflowerBlue  }
                "STEP"   { [System.Windows.Media.Brushes]::MediumAquamarine}
                default  { [System.Windows.Media.Brushes]::LightSteelBlue  }
            }
            $run = [System.Windows.Documents.Run]::new("[$($item.Timestamp)][$($item.Level)] $($item.Message)`n")
            $run.Foreground = $color
            $para = [System.Windows.Documents.Paragraph]::new($run)
            $para.Margin = [System.Windows.Thickness]::new(0,0,0,0)
            $lb.Document.Blocks.Add($para)
            $lb.ScrollToEnd()
        } catch {}
    }
}

function Get-Brush { param([string]$Hex)
    [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($Hex)) }

# ============================================================================
# XAML DEFINITION
# ============================================================================
[xml]$XAML = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Azure Virtual Desktop Manager v1.0"
    Height="900" Width="1440" MinHeight="700" MinWidth="1080"
    WindowStartupLocation="CenterScreen"
    Background="#080F1E" Foreground="#F1F5F9" FontFamily="Segoe UI"
    UseLayoutRounding="True"
    TextOptions.TextFormattingMode="Display"
    TextOptions.TextRenderingMode="ClearType"
    RenderOptions.BitmapScalingMode="HighQuality">
  <Window.Resources>

    <!-- ── Base text ─────────────────────────────────────────────────────── -->
    <Style TargetType="TextBlock">
      <Setter Property="Foreground" Value="#F1F5F9"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
      <Setter Property="TextOptions.TextRenderingMode" Value="ClearType"/>
    </Style>

    <!-- ── ToolTip ───────────────────────────────────────────────────────── -->
    <Style TargetType="ToolTip">
      <Setter Property="Background" Value="#0C1829"/>
      <Setter Property="Foreground" Value="#CBD5E1"/>
      <Setter Property="BorderBrush" Value="#1B324E"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="10,6"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="HasDropShadow" Value="True"/>
    </Style>

    <!-- ── Separator ─────────────────────────────────────────────────────── -->
    <Style TargetType="Separator">
      <Setter Property="Background" Value="#1B324E"/>
      <Setter Property="Height" Value="1"/>
      <Setter Property="Margin" Value="0,8"/>
    </Style>

    <!-- ── Label ─────────────────────────────────────────────────────────── -->
    <Style TargetType="Label">
      <Setter Property="Foreground" Value="#7A9DC4"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Padding" Value="0,6,0,2"/>
    </Style>

    <!-- ── Scrollbar — thin and subtle ──────────────────────────────────── -->
    <Style TargetType="ScrollBar">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Width" Value="5"/>
      <Setter Property="MinWidth" Value="5"/>
      <Setter Property="Opacity" Value="0.55"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ScrollBar">
            <Grid>
              <Track x:Name="PART_Track" IsDirectionReversed="True">
                <Track.Thumb>
                  <Thumb>
                    <Thumb.Template>
                      <ControlTemplate TargetType="Thumb">
                        <Border Background="#3B82F6" CornerRadius="3" Margin="1"/>
                      </ControlTemplate>
                    </Thumb.Template>
                  </Thumb>
                </Track.Thumb>
              </Track>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ── Nav button (inactive) ─────────────────────────────────────────── -->
    <Style x:Key="NavBtn" TargetType="Button">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="#64839E"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="14,10"/>
      <Setter Property="HorizontalContentAlignment" Value="Left"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FontSize" Value="12.5"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" Background="Transparent" CornerRadius="7" Margin="6,2" Padding="{TemplateBinding Padding}">
              <Border.RenderTransform><ScaleTransform ScaleX="1" ScaleY="1" CenterX="0.5" CenterY="0.5"/></Border.RenderTransform>
              <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <EventTrigger RoutedEvent="MouseEnter">
                <BeginStoryboard>
                  <Storyboard>
                    <ColorAnimation Storyboard.TargetName="Bd"
                      Storyboard.TargetProperty="(Border.Background).(SolidColorBrush.Color)"
                      To="#0F2238" Duration="0:0:0.12" FillBehavior="HoldEnd"/>
                  </Storyboard>
                </BeginStoryboard>
              </EventTrigger>
              <EventTrigger RoutedEvent="MouseLeave">
                <BeginStoryboard>
                  <Storyboard>
                    <ColorAnimation Storyboard.TargetName="Bd"
                      Storyboard.TargetProperty="(Border.Background).(SolidColorBrush.Color)"
                      To="#00000000" Duration="0:0:0.18" FillBehavior="HoldEnd"/>
                  </Storyboard>
                </BeginStoryboard>
              </EventTrigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Foreground" Value="#C0D8F0"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ── Nav button (active) with left accent bar ──────────────────────── -->
    <Style x:Key="NavActive" TargetType="Button">
      <Setter Property="Foreground" Value="#60B4FF"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="14,10"/>
      <Setter Property="HorizontalContentAlignment" Value="Left"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FontSize" Value="12.5"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Grid Margin="6,2">
              <Rectangle Width="3" HorizontalAlignment="Left" VerticalAlignment="Stretch" Margin="0,4,0,4" RadiusX="2" RadiusY="2">
                <Rectangle.Fill>
                  <LinearGradientBrush StartPoint="0,0" EndPoint="0,1">
                    <GradientStop Color="#60B4FF" Offset="0"/>
                    <GradientStop Color="#3B82F6" Offset="1"/>
                  </LinearGradientBrush>
                </Rectangle.Fill>
              </Rectangle>
              <Border x:Name="Bd" Background="#0E2040" Margin="3,0,0,0" CornerRadius="0,7,7,0" Padding="{TemplateBinding Padding}">
                <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" VerticalAlignment="Center"/>
              </Border>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ── Primary button ────────────────────────────────────────────────── -->
    <Style x:Key="BtnPrimary" TargetType="Button">
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="16,8"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" CornerRadius="7" Padding="{TemplateBinding Padding}">
              <Border.Background>
                <LinearGradientBrush StartPoint="0,0" EndPoint="0,1">
                  <GradientStop Color="#3B82F6" Offset="0"/>
                  <GradientStop Color="#2563EB" Offset="1"/>
                </LinearGradientBrush>
              </Border.Background>
              <Border.RenderTransform><ScaleTransform ScaleX="1" ScaleY="1" CenterX="0.5" CenterY="0.5"/></Border.RenderTransform>
              <Border.RenderTransformOrigin>0.5,0.5</Border.RenderTransformOrigin>
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <EventTrigger RoutedEvent="MouseEnter">
                <BeginStoryboard>
                  <Storyboard>
                    <DoubleAnimation Storyboard.TargetName="Bd"
                      Storyboard.TargetProperty="(Border.RenderTransform).(ScaleTransform.ScaleX)"
                      To="1.02" Duration="0:0:0.10"/>
                    <DoubleAnimation Storyboard.TargetName="Bd"
                      Storyboard.TargetProperty="(Border.RenderTransform).(ScaleTransform.ScaleY)"
                      To="1.02" Duration="0:0:0.10"/>
                  </Storyboard>
                </BeginStoryboard>
              </EventTrigger>
              <EventTrigger RoutedEvent="MouseLeave">
                <BeginStoryboard>
                  <Storyboard>
                    <DoubleAnimation Storyboard.TargetName="Bd"
                      Storyboard.TargetProperty="(Border.RenderTransform).(ScaleTransform.ScaleX)"
                      To="1.0" Duration="0:0:0.12"/>
                    <DoubleAnimation Storyboard.TargetName="Bd"
                      Storyboard.TargetProperty="(Border.RenderTransform).(ScaleTransform.ScaleY)"
                      To="1.0" Duration="0:0:0.12"/>
                  </Storyboard>
                </BeginStoryboard>
              </EventTrigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="Bd" Property="RenderTransform">
                  <Setter.Value><ScaleTransform ScaleX="0.96" ScaleY="0.96"/></Setter.Value>
                </Setter>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="Bd" Property="Opacity" Value="0.38"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ── Secondary button ──────────────────────────────────────────────── -->
    <Style x:Key="BtnSec" TargetType="Button" BasedOn="{StaticResource BtnPrimary}">
      <Setter Property="Foreground" Value="#A0BCD8"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" Background="#0F2038" CornerRadius="7"
                    BorderBrush="#1B324E" BorderThickness="1"
                    Padding="{TemplateBinding Padding}">
              <Border.RenderTransform><ScaleTransform ScaleX="1" ScaleY="1"/></Border.RenderTransform>
              <Border.RenderTransformOrigin>0.5,0.5</Border.RenderTransformOrigin>
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <EventTrigger RoutedEvent="MouseEnter">
                <BeginStoryboard>
                  <Storyboard>
                    <ColorAnimation Storyboard.TargetName="Bd"
                      Storyboard.TargetProperty="(Border.Background).(SolidColorBrush.Color)"
                      To="#162D4E" Duration="0:0:0.12"/>
                  </Storyboard>
                </BeginStoryboard>
              </EventTrigger>
              <EventTrigger RoutedEvent="MouseLeave">
                <BeginStoryboard>
                  <Storyboard>
                    <ColorAnimation Storyboard.TargetName="Bd"
                      Storyboard.TargetProperty="(Border.Background).(SolidColorBrush.Color)"
                      To="#0F2038" Duration="0:0:0.15"/>
                  </Storyboard>
                </BeginStoryboard>
              </EventTrigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="Bd" Property="RenderTransform">
                  <Setter.Value><ScaleTransform ScaleX="0.96" ScaleY="0.96"/></Setter.Value>
                </Setter>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="Bd" Property="Opacity" Value="0.35"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ── Green / Red action buttons ───────────────────────────────────── -->
    <Style x:Key="BtnGreen" TargetType="Button" BasedOn="{StaticResource BtnPrimary}">
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" CornerRadius="7" Padding="{TemplateBinding Padding}">
              <Border.Background><LinearGradientBrush StartPoint="0,0" EndPoint="0,1">
                <GradientStop Color="#16A34A" Offset="0"/><GradientStop Color="#15803D" Offset="1"/>
              </LinearGradientBrush></Border.Background>
              <Border.RenderTransform><ScaleTransform ScaleX="1" ScaleY="1"/></Border.RenderTransform>
              <Border.RenderTransformOrigin>0.5,0.5</Border.RenderTransformOrigin>
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Opacity" Value="0.85"/></Trigger>
              <Trigger Property="IsPressed" Value="True"><Setter TargetName="Bd" Property="RenderTransform"><Setter.Value><ScaleTransform ScaleX="0.96" ScaleY="0.96"/></Setter.Value></Setter></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="Bd" Property="Opacity" Value="0.35"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="BtnRed" TargetType="Button" BasedOn="{StaticResource BtnPrimary}">
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" CornerRadius="7" Padding="{TemplateBinding Padding}">
              <Border.Background><LinearGradientBrush StartPoint="0,0" EndPoint="0,1">
                <GradientStop Color="#DC2626" Offset="0"/><GradientStop Color="#B91C1C" Offset="1"/>
              </LinearGradientBrush></Border.Background>
              <Border.RenderTransform><ScaleTransform ScaleX="1" ScaleY="1"/></Border.RenderTransform>
              <Border.RenderTransformOrigin>0.5,0.5</Border.RenderTransformOrigin>
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Opacity" Value="0.85"/></Trigger>
              <Trigger Property="IsPressed" Value="True"><Setter TargetName="Bd" Property="RenderTransform"><Setter.Value><ScaleTransform ScaleX="0.96" ScaleY="0.96"/></Setter.Value></Setter></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="Bd" Property="Opacity" Value="0.35"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ── Card — elevated with shadow ───────────────────────────────────── -->
    <Style x:Key="Card" TargetType="Border">
      <Setter Property="Background" Value="#0C1829"/>
      <Setter Property="CornerRadius" Value="12"/>
      <Setter Property="Padding" Value="18"/>
      <Setter Property="BorderBrush" Value="#1B324E"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Effect">
        <Setter.Value>
          <DropShadowEffect Color="#000000" BlurRadius="28" ShadowDepth="3" Opacity="0.40" Direction="270"/>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="MetricCard" TargetType="Border" BasedOn="{StaticResource Card}">
      <Setter Property="Padding" Value="20"/>
    </Style>

    <!-- ── TextBox with focus glow ────────────────────────────────────────── -->
    <Style TargetType="TextBox">
      <Setter Property="Background" Value="#08152A"/>
      <Setter Property="Foreground" Value="#E2E8F0"/>
      <Setter Property="BorderBrush" Value="#1B324E"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="9,7"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="CaretBrush" Value="#3B82F6"/>
      <Setter Property="SelectionBrush" Value="#1D4ED8"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border x:Name="Bd" Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="7" Padding="{TemplateBinding Padding}">
              <ScrollViewer x:Name="PART_ContentHost" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsFocused" Value="True">
                <Setter TargetName="Bd" Property="BorderBrush" Value="#3B82F6"/>
                <Setter TargetName="Bd" Property="Effect">
                  <Setter.Value>
                    <DropShadowEffect Color="#3B82F6" BlurRadius="8" ShadowDepth="0" Opacity="0.30"/>
                  </Setter.Value>
                </Setter>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="BorderBrush" Value="#2D5480"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ── ComboBox Item ─────────────────────────────────────────────────── -->
    <Style TargetType="ComboBoxItem">
      <Setter Property="Background" Value="#08152A"/>
      <Setter Property="Foreground" Value="#E2E8F0"/>
      <Setter Property="Padding" Value="10,7"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBoxItem">
            <Border x:Name="Bd" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
              <ContentPresenter VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsHighlighted" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#0F2240"/>
                <Setter Property="Foreground" Value="White"/>
              </Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#1D3D6B"/>
                <Setter Property="Foreground" Value="#60B4FF"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ── ComboBox with dark popup ──────────────────────────────────────── -->
    <Style TargetType="ComboBox">
      <Setter Property="Background" Value="#08152A"/>
      <Setter Property="Foreground" Value="#E2E8F0"/>
      <Setter Property="BorderBrush" Value="#1B324E"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="9,7"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="MaxDropDownHeight" Value="320"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBox">
            <Grid>
              <ToggleButton x:Name="TogBtn" Focusable="False" ClickMode="Press"
                IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}">
                <ToggleButton.Template>
                  <ControlTemplate TargetType="ToggleButton">
                    <Border x:Name="TgBd" Background="{TemplateBinding Background}"
                            BorderBrush="{TemplateBinding BorderBrush}"
                            BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="7">
                      <Grid>
                        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="24"/></Grid.ColumnDefinitions>
                        <Path Grid.Column="1" Data="M 0 0 L 4 5 L 8 0 Z" Fill="#5080A0"
                              HorizontalAlignment="Center" VerticalAlignment="Center" Margin="0,1,0,0"/>
                      </Grid>
                    </Border>
                    <ControlTemplate.Triggers>
                      <Trigger Property="IsMouseOver" Value="True">
                        <Setter TargetName="TgBd" Property="BorderBrush" Value="#2D5480"/>
                        <Setter TargetName="TgBd" Property="Background" Value="#0F1E35"/>
                      </Trigger>
                    </ControlTemplate.Triggers>
                  </ControlTemplate>
                </ToggleButton.Template>
              </ToggleButton>
              <ContentPresenter x:Name="ContentSite" IsHitTestVisible="False"
                Margin="9,0,28,0" VerticalAlignment="Center"
                Content="{TemplateBinding SelectionBoxItem}"
                ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"/>
              <TextBox x:Name="PART_EditableTextBox" Visibility="Hidden" IsReadOnly="{TemplateBinding IsReadOnly}"
                Background="Transparent" Foreground="#E2E8F0" BorderThickness="0"
                Margin="9,0,28,0" VerticalAlignment="Center" Focusable="True" CaretBrush="#3B82F6"/>
              <Popup x:Name="Popup" Placement="Bottom"
                IsOpen="{TemplateBinding IsDropDownOpen}"
                AllowsTransparency="True" Focusable="False" PopupAnimation="Fade">
                <Grid SnapsToDevicePixels="True"
                  MinWidth="{TemplateBinding ActualWidth}"
                  MaxHeight="{TemplateBinding MaxDropDownHeight}">
                  <Border Background="#08152A" BorderBrush="#2D5480" BorderThickness="1" CornerRadius="0,0,7,7">
                    <Border.Effect>
                      <DropShadowEffect Color="#000000" BlurRadius="20" ShadowDepth="4" Opacity="0.50"/>
                    </Border.Effect>
                    <ScrollViewer SnapsToDevicePixels="True" Background="#08152A">
                      <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained"/>
                    </ScrollViewer>
                  </Border>
                </Grid>
              </Popup>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsGrouping" Value="True">
                <Setter Property="ScrollViewer.CanContentScroll" Value="False"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ── DataGrid ───────────────────────────────────────────────────────── -->
    <Style TargetType="DataGrid">
      <Setter Property="Background" Value="#08152A"/>
      <Setter Property="Foreground" Value="#E2E8F0"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="GridLinesVisibility" Value="Horizontal"/>
      <Setter Property="HorizontalGridLinesBrush" Value="#0F2038"/>
      <Setter Property="RowBackground" Value="#08152A"/>
      <Setter Property="AlternatingRowBackground" Value="#0C1829"/>
      <Setter Property="HeadersVisibility" Value="Column"/>
      <Setter Property="AutoGenerateColumns" Value="False"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="CanUserResizeRows" Value="False"/>
      <Setter Property="SelectionMode" Value="Single"/>
      <Setter Property="VirtualizingPanel.ScrollUnit" Value="Pixel"/>
      <Setter Property="VirtualizingPanel.IsVirtualizing" Value="True"/>
      <Setter Property="EnableRowVirtualization" Value="True"/>
      <Setter Property="ScrollViewer.CanContentScroll" Value="True"/>
    </Style>
    <Style TargetType="DataGridColumnHeader">
      <Setter Property="Background" Value="#0B1E36"/>
      <Setter Property="Foreground" Value="#7A9DC4"/>
      <Setter Property="Padding" Value="12,9"/>
      <Setter Property="BorderThickness" Value="0,0,0,1"/>
      <Setter Property="BorderBrush" Value="#1B324E"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style TargetType="DataGridRow">
      <Setter Property="Foreground" Value="#CBD5E1"/>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#0F2238"/>
          <Setter Property="Foreground" Value="#F1F5F9"/>
        </Trigger>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="#1A3A60"/>
          <Setter Property="Foreground" Value="White"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style TargetType="DataGridCell">
      <Setter Property="Padding" Value="12,8"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="DataGridCell">
            <Border Padding="{TemplateBinding Padding}" Background="Transparent">
              <ContentPresenter VerticalAlignment="Center"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ── ProgressBar — gradient fill ──────────────────────────────────── -->
    <Style TargetType="ProgressBar">
      <Setter Property="Background" Value="#0F2038"/>
      <Setter Property="Foreground" Value="#3B82F6"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Height" Value="5"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ProgressBar">
            <Border Background="{TemplateBinding Background}" CornerRadius="3">
              <Border x:Name="PART_Track">
                <Border x:Name="PART_Indicator" HorizontalAlignment="Left" CornerRadius="3">
                  <Border.Background>
                    <LinearGradientBrush StartPoint="0,0" EndPoint="1,0">
                      <GradientStop Color="#3B82F6" Offset="0"/>
                      <GradientStop Color="#60B4FF" Offset="1"/>
                    </LinearGradientBrush>
                  </Border.Background>
                </Border>
              </Border>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ── CheckBox ──────────────────────────────────────────────────────── -->
    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="#CBD5E1"/>
      <Setter Property="FontSize" Value="12"/>
    </Style>

    <!-- ── RadioButton ───────────────────────────────────────────────────── -->
    <Style TargetType="RadioButton">
      <Setter Property="Foreground" Value="#CBD5E1"/>
      <Setter Property="FontSize" Value="12"/>
    </Style>

    <!-- ── Slider ────────────────────────────────────────────────────────── -->
    <Style TargetType="Slider">
      <Setter Property="Foreground" Value="#3B82F6"/>
    </Style>

    <!-- ── TabControl / TabItem ──────────────────────────────────────────── -->
    <Style TargetType="TabControl">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderThickness" Value="0"/>
    </Style>
    <Style TargetType="TabItem">
      <Setter Property="Background" Value="#08152A"/>
      <Setter Property="Foreground" Value="#64839E"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="14,8"/>
      <Setter Property="FontSize" Value="12"/>
    </Style>

  </Window.Resources>
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="54"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="30"/>
    </Grid.RowDefinitions>
    <!-- HEADER -->
    <Border Grid.Row="0" BorderBrush="#1B324E" BorderThickness="0,0,0,1">
      <Border.Background><LinearGradientBrush StartPoint="0,0" EndPoint="1,0"><GradientStop Color="#060D1A" Offset="0"/><GradientStop Color="#080F1E" Offset="0.5"/><GradientStop Color="#060D1A" Offset="1"/></LinearGradientBrush></Border.Background>
      <Grid Margin="16,0">
        <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <Border Background="#0078D4" CornerRadius="8" Width="30" Height="30" Margin="0,0,10,0">
            <TextBlock Text="AVD" FontSize="9" FontWeight="Bold" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <StackPanel VerticalAlignment="Center">
            <TextBlock Text="Azure Virtual Desktop Manager" FontSize="14" FontWeight="SemiBold" Foreground="#F1F5F9"/>
            <TextBlock Text="Universal Deployment and Management Console" FontSize="10" Foreground="#334155"/>
          </StackPanel>
        </StackPanel>
        <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Center">
          <Border Background="#0B1E35" CornerRadius="20" Padding="12,3" Margin="4,0" BorderBrush="#1B324E" BorderThickness="1">
            <StackPanel Orientation="Horizontal">
              <Ellipse x:Name="DotConn" Width="7" Height="7" Fill="#475569" Margin="0,0,6,0" VerticalAlignment="Center"/>
              <TextBlock x:Name="TxtSubName" Text="Not Connected" FontSize="11" Foreground="#94A3B8" VerticalAlignment="Center"/>
            </StackPanel>
          </Border>
          <Border Background="#0D2547" CornerRadius="20" Padding="6,2" Margin="4,0">
            <ComboBox x:Name="CboSubSwitcher" Width="240" FontSize="11" Background="#0D2547" Foreground="#50ABF1" BorderThickness="0" ToolTip="Switch subscription"/>
          </Border>
          <Border Background="#1A1200" CornerRadius="20" Padding="12,3" Margin="4,0" BorderBrush="#2D1E00" BorderThickness="1">
            <TextBlock x:Name="TxtLicBadge" Text="Licenses: unknown" FontSize="11" Foreground="#F59E0B"/>
          </Border>
        </StackPanel>
        <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
          <Button x:Name="BtnRefreshAll" Content="Refresh" Style="{StaticResource BtnSec}" Margin="4,0" Padding="12,5" FontSize="11"/>
          <Button x:Name="BtnConnect" Content="Connect to Azure" Style="{StaticResource BtnPrimary}" Margin="4,0" Padding="14,6" FontSize="12"/>
        </StackPanel>
      </Grid>
    </Border>
    <!-- BODY -->
    <Grid Grid.Row="1" Background="#080F1E">
      <Grid.ColumnDefinitions><ColumnDefinition Width="200"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
      <!-- SIDEBAR -->
      <Border BorderBrush="#1B324E" BorderThickness="0,0,1,0">
        <Border.Background><LinearGradientBrush StartPoint="0,0" EndPoint="0,1"><GradientStop Color="#060C18" Offset="0"/><GradientStop Color="#07101E" Offset="1"/></LinearGradientBrush></Border.Background>
        <DockPanel>
          <StackPanel DockPanel.Dock="Top" Margin="0,8,0,0">
            <TextBlock Text="OVERVIEW" FontSize="9" Foreground="#334155" FontWeight="Bold" Margin="20,6,0,4"/>
            <Button x:Name="NavDash"    Content="Dashboard"         Tag="Dash"   Style="{StaticResource NavActive}"/>
            <Button x:Name="NavLicense" Content="License Check"     Tag="Lic"    Style="{StaticResource NavBtn}"/>
            <Separator Margin="14,8"/>
            <TextBlock Text="DEPLOY" FontSize="9" Foreground="#334155" FontWeight="Bold" Margin="20,4,0,4"/>
            <Button x:Name="NavWizard"  Content="New Deployment"    Tag="Wiz"    Style="{StaticResource NavBtn}"/>
            <Button x:Name="NavHP"      Content="Host Pools"        Tag="HP"     Style="{StaticResource NavBtn}"/>
            <Button x:Name="NavSess"    Content="Session Hosts"     Tag="Sess"   Style="{StaticResource NavBtn}"/>
            <Button x:Name="NavAG"      Content="App Groups"        Tag="AG"     Style="{StaticResource NavBtn}"/>
            <Separator Margin="14,8"/>
            <TextBlock Text="MANAGE" FontSize="9" Foreground="#334155" FontWeight="Bold" Margin="20,4,0,4"/>
            <Button x:Name="NavFSL"     Content="FSLogix Profiles"  Tag="FSL"    Style="{StaticResource NavBtn}"/>
            <Button x:Name="NavScale"   Content="Auto-Scaling"      Tag="Scale"  Style="{StaticResource NavBtn}"/>
            <Button x:Name="NavMon"     Content="Monitoring"        Tag="Mon"    Style="{StaticResource NavBtn}"/>
            <Button x:Name="NavRDP"     Content="RDP and Security"  Tag="RDP"    Style="{StaticResource NavBtn}"/>
            <Separator Margin="14,8"/>
            <Button x:Name="NavBP"      Content="Best Practices"    Tag="BP"     Style="{StaticResource NavBtn}"/>
            <Button x:Name="NavLog"     Content="Activity Log"      Tag="Log"    Style="{StaticResource NavBtn}"/>
            <Button x:Name="NavSet"     Content="Settings"          Tag="Set"    Style="{StaticResource NavBtn}"/>
          </StackPanel>
          <StackPanel DockPanel.Dock="Bottom" Margin="16,0,16,10">
            <Separator/>
            <TextBlock Text="AVD Manager v1.0" FontSize="10" Foreground="#2D5480" FontWeight="SemiBold"/>
          </StackPanel>
        </DockPanel>
      </Border>
      <!-- CONTENT -->
      <Grid Grid.Column="1">
        <!-- Loading overlay — shown during data fetches -->
        <Border x:Name="LoadingOverlay" Visibility="Collapsed" Panel.ZIndex="99"
                Background="#CC080F1E">
          <StackPanel HorizontalAlignment="Center" VerticalAlignment="Center">
            <Border Width="60" Height="60" CornerRadius="30" Margin="0,0,0,14" Background="#0F2238"
                    BorderBrush="#1B324E" BorderThickness="1">
              <Border.Effect><DropShadowEffect Color="#3B82F6" BlurRadius="20" ShadowDepth="0" Opacity="0.4"/></Border.Effect>
              <TextBlock Text="..." FontSize="22" Foreground="#3B82F6"
                         HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <TextBlock x:Name="LoadingText" Text="Loading..." FontSize="13"
                       Foreground="#7A9DC4" HorizontalAlignment="Center"/>
          </StackPanel>
        </Border>
        <ScrollViewer x:Name="PanelDash" Visibility="Visible" Padding="20">
          <StackPanel>
            <Grid Margin="0,0,0,14">
              <StackPanel>
                <TextBlock Text="Dashboard" FontSize="22" FontWeight="SemiBold"/>
                <TextBlock x:Name="TxtDashSub" Text="Connect to Azure to view your AVD environment" FontSize="12" Foreground="#5A7A99"/>
              </StackPanel>
              <Button x:Name="BtnNewDeploy" Content="+ New Deployment" Style="{StaticResource BtnPrimary}" HorizontalAlignment="Right" VerticalAlignment="Bottom" Padding="14,8"/>
            </Grid>
            <Grid Margin="0,0,0,12">
              <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition/><ColumnDefinition/><ColumnDefinition/><ColumnDefinition/></Grid.ColumnDefinitions>
              <Border Grid.Column="0" Style="{StaticResource Card}" Margin="0,0,6,0">
                <StackPanel>
                  <TextBlock Text="HOST POOLS" FontSize="9" Foreground="#64748B" FontWeight="SemiBold"/>
                  <TextBlock x:Name="MetHP" Text="-" FontSize="30" FontWeight="Bold" Foreground="#50ABF1" Margin="0,4,0,2"/>
                  <TextBlock x:Name="MetHPSub" Text="total" FontSize="11" Foreground="#475569"/>
                </StackPanel>
              </Border>
              <Border Grid.Column="1" Style="{StaticResource Card}" Margin="3,0">
                <StackPanel>
                  <TextBlock Text="SESSION HOSTS" FontSize="9" Foreground="#64748B" FontWeight="SemiBold"/>
                  <TextBlock x:Name="MetHosts" Text="-" FontSize="30" FontWeight="Bold" Foreground="#10B981" Margin="0,4,0,2"/>
                  <TextBlock x:Name="MetHostsSub" Text="avail / total" FontSize="11" Foreground="#475569"/>
                </StackPanel>
              </Border>
              <Border Grid.Column="2" Style="{StaticResource Card}" Margin="3,0">
                <StackPanel>
                  <TextBlock Text="ACTIVE SESSIONS" FontSize="9" Foreground="#64748B" FontWeight="SemiBold"/>
                  <TextBlock x:Name="MetSess" Text="-" FontSize="30" FontWeight="Bold" Foreground="#FBB040" Margin="0,4,0,2"/>
                  <TextBlock Text="across all pools" FontSize="11" Foreground="#475569"/>
                </StackPanel>
              </Border>
              <Border Grid.Column="3" Style="{StaticResource Card}" Margin="3,0">
                <StackPanel>
                  <TextBlock Text="AVD LICENSES" FontSize="9" Foreground="#64748B" FontWeight="SemiBold"/>
                  <TextBlock x:Name="MetLic" Text="-" FontSize="30" FontWeight="Bold" Foreground="#A78BFA" Margin="0,4,0,2"/>
                  <TextBlock Text="entitled users" FontSize="11" Foreground="#475569"/>
                </StackPanel>
              </Border>
              <Border Grid.Column="4" Style="{StaticResource Card}" Margin="6,0,0,0">
                <StackPanel>
                  <TextBlock Text="WORKSPACES" FontSize="9" Foreground="#64748B" FontWeight="SemiBold"/>
                  <TextBlock x:Name="MetWS" Text="-" FontSize="30" FontWeight="Bold" Foreground="#34D399" Margin="0,4,0,2"/>
                  <TextBlock Text="published" FontSize="11" Foreground="#475569"/>
                </StackPanel>
              </Border>
            </Grid>
            <Border Style="{StaticResource Card}">
              <DockPanel>
                <Grid DockPanel.Dock="Top" Margin="0,0,0,10">
                  <TextBlock Text="Host Pools" FontSize="14" FontWeight="SemiBold"/>
                  <Button x:Name="BtnDashRefresh" Content="Refresh" Style="{StaticResource BtnSec}" HorizontalAlignment="Right" Padding="10,4" FontSize="11"/>
                </Grid>
                <DataGrid x:Name="GridDash" Height="280" CanUserSortColumns="True">
                  <DataGrid.Columns>
                    <DataGridTextColumn Header="Name"         Binding="{Binding Name}"     Width="200"/>
                    <DataGridTemplateColumn Header="Type" Width="100">
                      <DataGridTemplateColumn.CellTemplate><DataTemplate>
                        <Border CornerRadius="4" Padding="6,2" Background="{Binding TypeBg}" HorizontalAlignment="Left">
                          <TextBlock Text="{Binding Type}" FontSize="10" FontWeight="SemiBold" Foreground="{Binding TypeFg}"/>
                        </Border>
                      </DataTemplate></DataGridTemplateColumn.CellTemplate>
                    </DataGridTemplateColumn>
                    <DataGridTextColumn Header="Load Balancing" Binding="{Binding LB}"     Width="120"/>
                    <DataGridTextColumn Header="Max Sessions"   Binding="{Binding MaxSess}" Width="100"/>
                    <DataGridTextColumn Header="Hosts (Avail)"  Binding="{Binding Hosts}"  Width="100"/>
                    <DataGridTextColumn Header="Sessions"        Binding="{Binding Sessions}" Width="80"/>
                    <DataGridTemplateColumn Header="Status" Width="100">
                      <DataGridTemplateColumn.CellTemplate><DataTemplate>
                        <Border CornerRadius="4" Padding="6,2" Background="{Binding StatusBg}" HorizontalAlignment="Left">
                          <TextBlock Text="{Binding Status}" FontSize="10" FontWeight="SemiBold" Foreground="{Binding StatusFg}"/>
                        </Border>
                      </DataTemplate></DataGridTemplateColumn.CellTemplate>
                    </DataGridTemplateColumn>
                    <DataGridTextColumn Header="Region"        Binding="{Binding Location}" Width="110"/>
                    <DataGridTextColumn Header="Resource Group" Binding="{Binding RG}"     Width="*"/>
                  </DataGrid.Columns>
                </DataGrid>
              </DockPanel>
            </Border>
          </StackPanel>
        </ScrollViewer>

        <!-- ===== LICENSE CHECK ===== -->
        <ScrollViewer x:Name="PanelLic" Visibility="Collapsed" Padding="20">
          <StackPanel>
            <Grid Margin="0,0,0,14">
              <StackPanel>
                <TextBlock Text="License Assessment" FontSize="22" FontWeight="SemiBold"/>
                <TextBlock Text="Scans Microsoft Graph API to detect AVD entitlement and supporting licenses" FontSize="12" Foreground="#64748B"/>
              </StackPanel>
              <Button x:Name="BtnScanLic" Content="Scan Licenses" Style="{StaticResource BtnPrimary}" HorizontalAlignment="Right" VerticalAlignment="Bottom" Padding="16,8"/>
            </Grid>
            <Grid Margin="0,0,0,12">
              <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition/><ColumnDefinition/></Grid.ColumnDefinitions>
              <Border Grid.Column="0" Style="{StaticResource Card}" Margin="0,0,6,0">
                <StackPanel>
                  <TextBlock Text="AVD-ENTITLED USERS" FontSize="9" Foreground="#64748B" FontWeight="SemiBold"/>
                  <TextBlock x:Name="LicTotalUsers" Text="-" FontSize="30" FontWeight="Bold" Foreground="#10B981" Margin="0,4,0,2"/>
                  <TextBlock x:Name="LicTotalSub" Text="assigned seats" FontSize="11" Foreground="#475569"/>
                </StackPanel>
              </Border>
              <Border Grid.Column="1" Style="{StaticResource Card}" Margin="3,0">
                <StackPanel>
                  <TextBlock Text="ENTITLEMENT LEVEL" FontSize="9" Foreground="#64748B" FontWeight="SemiBold"/>
                  <TextBlock x:Name="LicType" Text="-" FontSize="16" FontWeight="Bold" Foreground="#50ABF1" Margin="0,4,0,2" TextWrapping="Wrap"/>
                  <TextBlock x:Name="LicTypeSub" Text="highest level found" FontSize="11" Foreground="#475569"/>
                </StackPanel>
              </Border>
              <Border Grid.Column="2" Style="{StaticResource Card}" Margin="6,0,0,0">
                <StackPanel>
                  <TextBlock Text="SUPPORTING LICENSES" FontSize="9" Foreground="#64748B" FontWeight="SemiBold"/>
                  <TextBlock x:Name="LicSupport" Text="-" FontSize="30" FontWeight="Bold" Foreground="#A78BFA" Margin="0,4,0,2"/>
                  <TextBlock Text="Intune, Entra ID P1/P2" FontSize="11" Foreground="#475569"/>
                </StackPanel>
              </Border>
            </Grid>
            <Grid>
              <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="300"/></Grid.ColumnDefinitions>
              <Border Grid.Column="0" Style="{StaticResource Card}" Margin="0,0,8,0">
                <DockPanel>
                  <TextBlock DockPanel.Dock="Top" Text="Detected License SKUs" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,10"/>
                  <DataGrid x:Name="GridLic" CanUserSortColumns="True">
                    <DataGrid.Columns>
                      <DataGridTextColumn Header="AVD"      Binding="{Binding AvdIcon}"   Width="50"/>
                      <DataGridTextColumn Header="Name"     Binding="{Binding LicName}"   Width="*"/>
                      <DataGridTextColumn Header="SKU"      Binding="{Binding SkuId}"     Width="200"/>
                      <DataGridTextColumn Header="Assigned" Binding="{Binding Assigned}"  Width="80"/>
                      <DataGridTextColumn Header="Avail"    Binding="{Binding Available}" Width="70"/>
                      <DataGridTextColumn Header="Type"     Binding="{Binding LicType}"   Width="130"/>
                    </DataGrid.Columns>
                  </DataGrid>
                </DockPanel>
              </Border>
              <Border Grid.Column="1" Style="{StaticResource Card}" Margin="8,0,0,0">
                <ScrollViewer>
                  <StackPanel>
                    <TextBlock Text="Requirements Check" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,10"/>
                    <DataGrid x:Name="GridLicReqs" HeadersVisibility="None" CanUserSortColumns="False">
                      <DataGrid.Columns>
                        <DataGridTextColumn Binding="{Binding Icon}" Width="36"/>
                        <DataGridTextColumn Binding="{Binding Req}"  Width="*"/>
                      </DataGrid.Columns>
                    </DataGrid>
                    <Separator Margin="0,10"/>
                    <TextBlock Text="Recommendations" FontSize="12" FontWeight="SemiBold" Foreground="#64748B" Margin="0,0,0,6"/>
                    <TextBlock x:Name="TxtLicRec" Text="Run a scan to see recommendations." FontSize="11" Foreground="#475569" TextWrapping="Wrap"/>
                  </StackPanel>
                </ScrollViewer>
              </Border>
            </Grid>
          </StackPanel>
        </ScrollViewer>

        <!-- ===== WIZARD ===== -->
        <Grid x:Name="PanelWiz" Visibility="Collapsed">
          <Grid.ColumnDefinitions><ColumnDefinition Width="250"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
          <!-- Steps sidebar -->
          <Border Background="#070F1C" BorderBrush="#1E3A5F" BorderThickness="0,0,1,0" Padding="14">
            <StackPanel>
              <TextBlock Text="New Deployment" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,2"/>
              <TextBlock x:Name="TxtWizType" Text="Select deployment type below" FontSize="10" Foreground="#64748B" Margin="0,0,0,12"/>
              <TextBlock Text="Deployment Type" FontSize="10" Foreground="#64748B" FontWeight="SemiBold" Margin="0,0,0,6"/>
              <RadioButton x:Name="RdoPooled"    Content="Pooled Desktop"    GroupName="HPType" IsChecked="True" Margin="0,3"/>
              <RadioButton x:Name="RdoPersonal"  Content="Personal Desktop"  GroupName="HPType" Margin="0,3"/>
              <RadioButton x:Name="RdoRemoteApp" Content="RemoteApp"         GroupName="HPType" Margin="0,3"/>
              <Separator Margin="0,10"/>
              <ItemsControl x:Name="WizStepList">
                <ItemsControl.ItemTemplate>
                  <DataTemplate>
                    <Grid Margin="0,3">
                      <Grid.ColumnDefinitions><ColumnDefinition Width="26"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                      <Border Width="20" Height="20" CornerRadius="10" Background="{Binding StepBg}" HorizontalAlignment="Center">
                        <TextBlock Text="{Binding Num}" FontSize="10" FontWeight="Bold" Foreground="{Binding NumFg}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                      </Border>
                      <StackPanel Grid.Column="1" VerticalAlignment="Center" Margin="6,0,0,0">
                        <TextBlock Text="{Binding Title}" FontSize="11" FontWeight="{Binding Weight}" Foreground="{Binding TitleFg}"/>
                        <TextBlock Text="{Binding Sub}" FontSize="9" Foreground="#334155"/>
                      </StackPanel>
                    </Grid>
                  </DataTemplate>
                </ItemsControl.ItemTemplate>
              </ItemsControl>
              <Separator Margin="0,10"/>
              <ProgressBar x:Name="WizProg" Maximum="100" Value="12" Height="5" Margin="0,0,0,4"/>
              <TextBlock x:Name="TxtWizStep" Text="Step 1 of 8" FontSize="10" Foreground="#64748B"/>
            </StackPanel>
          </Border>
          <!-- Step content -->
          <Grid Grid.Column="1">
            <!-- Step 1: Basics -->
            <ScrollViewer x:Name="WS1" Padding="22" Visibility="Visible">
              <StackPanel>
                <TextBlock Text="Basics" FontSize="20" FontWeight="SemiBold" Margin="0,0,0,2"/>
                <TextBlock Text="Name your host pool and choose the deployment scope" FontSize="12" Foreground="#64748B" Margin="0,0,0,16"/>
                <Border Style="{StaticResource Card}" Margin="0,0,0,10">
                  <StackPanel>
                    <TextBlock Text="Host Pool Identity" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,12"/>
                    <Label Content="Host Pool Name *"/>
                    <TextBox x:Name="WizHPName" Margin="0,2,0,10" ToolTip="alphanumeric and hyphens, max 64 chars"/>
                    <Label Content="Friendly Display Name"/>
                    <TextBox x:Name="WizHPFriendly" Margin="0,2,0,10"/>
                    <Label Content="Description"/>
                    <TextBox x:Name="WizHPDesc" Margin="0,2,0,10"/>
                    <Grid>
                      <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition/></Grid.ColumnDefinitions>
                      <StackPanel>
                        <Label Content="Subscription *"/>
                        <ComboBox x:Name="WizSub" Margin="0,2,8,10"/>
                      </StackPanel>
                      <StackPanel Grid.Column="1">
                        <Label Content="Resource Group *"/>
                        <Grid>
                          <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                          <ComboBox x:Name="WizRG" Margin="0,2,4,10"/>
                          <Button x:Name="BtnNewRG" Content="+ New" Grid.Column="1" Style="{StaticResource BtnSec}" Padding="8,6" Margin="0,2,0,10" FontSize="11"/>
                        </Grid>
                      </StackPanel>
                    </Grid>
                    <Grid>
                      <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition/></Grid.ColumnDefinitions>
                      <StackPanel>
                        <Label Content="Azure Region *"/>
                        <ComboBox x:Name="WizRegion" Margin="0,2,8,10"/>
                      </StackPanel>
                      <StackPanel Grid.Column="1">
                        <Label Content="Workspace"/>
                        <Grid>
                          <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                          <ComboBox x:Name="WizWS" Margin="0,2,4,10"/>
                          <Button x:Name="BtnNewWS" Content="+ New" Grid.Column="1" Style="{StaticResource BtnSec}" Padding="8,6" Margin="0,2,0,10" FontSize="11"/>
                        </Grid>
                      </StackPanel>
                    </Grid>
                  </StackPanel>
                </Border>
                <Border Style="{StaticResource Card}">
                  <StackPanel>
                    <TextBlock Text="Tags (optional)" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,10"/>
                    <Grid>
                      <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition Width="Auto"/><ColumnDefinition/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                      <TextBox x:Name="WizTagK" Margin="0,0,4,0" FontSize="12"/>
                      <TextBlock Grid.Column="1" Text="=" VerticalAlignment="Center" Margin="4,0"/>
                      <TextBox x:Name="WizTagV" Grid.Column="2" Margin="4,0,4,0" FontSize="12"/>
                      <Button x:Name="BtnAddTag" Grid.Column="3" Content="Add" Style="{StaticResource BtnSec}" Padding="10,6" FontSize="11"/>
                    </Grid>
                    <DataGrid x:Name="GridTags" Height="70" Margin="0,6,0,0" CanUserSortColumns="False">
                      <DataGrid.Columns>
                        <DataGridTextColumn Header="Key" Binding="{Binding TagKey}" Width="*"/>
                        <DataGridTextColumn Header="Value" Binding="{Binding TagVal}" Width="*"/>
                      </DataGrid.Columns>
                    </DataGrid>
                  </StackPanel>
                </Border>
              </StackPanel>
            </ScrollViewer>
            <!-- Step 2: Host Pool Config -->
            <ScrollViewer x:Name="WS2" Padding="22" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="Host Pool Configuration" FontSize="20" FontWeight="SemiBold" Margin="0,0,0,2"/>
                <TextBlock Text="Configure load balancing, session limits, and advanced options" FontSize="12" Foreground="#64748B" Margin="0,0,0,16"/>
                <Border Style="{StaticResource Card}" Margin="0,0,0,10">
                  <StackPanel>
                    <TextBlock Text="Load Balancing Algorithm (Pooled / RemoteApp)" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,10"/>
                    <RadioButton x:Name="RdoBreadth"  Content="BreadthFirst - Spread new sessions evenly across all available hosts (recommended for most scenarios)" GroupName="LB" IsChecked="True" Margin="0,4"/>
                    <RadioButton x:Name="RdoDepth"    Content="DepthFirst - Fill one host completely before moving to the next (reduces costs by minimizing active VMs)" GroupName="LB" Margin="0,4"/>
                    <Separator/>
                    <TextBlock Text="Maximum Sessions per Host" FontSize="11" Foreground="#64748B" Margin="0,8,0,4"/>
                    <StackPanel Orientation="Horizontal">
                      <Slider x:Name="SliderMaxSess" Minimum="1" Maximum="50" Value="8" Width="260" VerticalAlignment="Center"/>
                      <TextBlock x:Name="TxtMaxSess" Text="8" FontSize="18" FontWeight="Bold" Foreground="#50ABF1" Margin="12,0,0,0" VerticalAlignment="Center"/>
                      <TextBlock Text=" sessions per host" FontSize="12" Foreground="#64748B" VerticalAlignment="Center"/>
                    </StackPanel>
                  </StackPanel>
                </Border>
                <Border Style="{StaticResource Card}" Margin="0,0,0,10">
                  <StackPanel>
                    <TextBlock Text="Personal Desktop Assignment (Personal pools only)" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,10"/>
                    <RadioButton x:Name="RdoAutoAssign"   Content="Automatic - Assign the first available host when user connects (recommended)" GroupName="Assign" IsChecked="True" Margin="0,4"/>
                    <RadioButton x:Name="RdoDirectAssign" Content="Direct - Administrator pre-assigns specific users to specific VMs" GroupName="Assign" Margin="0,4"/>
                  </StackPanel>
                </Border>
                <Border Style="{StaticResource Card}">
                  <StackPanel>
                    <TextBlock Text="Advanced Options" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,10"/>
                    <CheckBox x:Name="ChkStartVM"       Content="Start VM on Connect - Automatically start deallocated hosts when users connect (reduces idle costs)" IsChecked="True" Margin="0,4"/>
                    <CheckBox x:Name="ChkValidation"    Content="Validation Environment - Use this pool to test AVD service updates before production rollout" Margin="0,4"/>
                    <Separator/>
                    <TextBlock Text="Registration Token Expiry" FontSize="11" Foreground="#64748B" Margin="0,8,0,4"/>
                    <StackPanel Orientation="Horizontal">
                      <Slider x:Name="SliderTokenHrs" Minimum="1" Maximum="720" Value="48" Width="220" VerticalAlignment="Center"/>
                      <TextBlock x:Name="TxtTokenHrs" Text="48 hours" FontSize="12" Foreground="#50ABF1" Margin="10,0,0,0" VerticalAlignment="Center"/>
                    </StackPanel>
                  </StackPanel>
                </Border>
              </StackPanel>
            </ScrollViewer>
            <!-- Step 3: Networking -->
            <ScrollViewer x:Name="WS3" Padding="22" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="Virtual Network and Connectivity" FontSize="20" FontWeight="SemiBold" Margin="0,0,0,2"/>
                <TextBlock Text="Session hosts must be in a VNet with connectivity to identity services and the internet" FontSize="12" Foreground="#64748B" Margin="0,0,0,16"/>
                <Border Style="{StaticResource Card}" Margin="0,0,0,10">
                  <StackPanel>
                    <TextBlock Text="Network Configuration" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,12"/>
                    <Grid>
                      <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition/></Grid.ColumnDefinitions>
                      <StackPanel>
                        <Label Content="VNet Resource Group"/>
                        <ComboBox x:Name="WizVNetRG" Margin="0,2,8,10"/>
                      </StackPanel>
                      <StackPanel Grid.Column="1">
                        <Label Content="Virtual Network *"/>
                        <Grid>
                          <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                          <ComboBox x:Name="WizVNet" Margin="0,2,4,10"/>
                          <Button x:Name="BtnLoadVNet" Content="Load" Grid.Column="1" Style="{StaticResource BtnSec}" Padding="8,6" Margin="0,2,0,10" FontSize="11"/>
                        </Grid>
                      </StackPanel>
                    </Grid>
                    <Label Content="Subnet *"/>
                    <ComboBox x:Name="WizSubnet" Margin="0,2,0,10"/>
                    <Separator/>
                    <TextBlock Text="Network Checklist" FontSize="11" FontWeight="SemiBold" Foreground="#64748B" Margin="0,8,0,6"/>
                    <StackPanel>
                      <TextBlock Text="[OK] Outbound TCP 443 / 80 to *.wvd.microsoft.com" FontSize="11" Foreground="#64748B" Margin="0,2"/>
                      <TextBlock Text="[OK] Outbound TCP 443 to *.microsoftonline.com" FontSize="11" Foreground="#64748B" Margin="0,2"/>
                      <TextBlock Text="[OK] DNS resolution for Azure internal FQDNs" FontSize="11" Foreground="#64748B" Margin="0,2"/>
                      <TextBlock Text="[!]  For Hybrid AD: VNet DNS pointing to domain controllers" FontSize="11" Foreground="#F59E0B" Margin="0,2"/>
                    </StackPanel>
                  </StackPanel>
                </Border>
                <Border Style="{StaticResource Card}">
                  <StackPanel>
                    <TextBlock Text="VM Admin Credentials" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,12"/>
                    <Label Content="Key Vault (for credential storage)"/>
                    <Grid>
                      <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                      <ComboBox x:Name="WizKV" Margin="0,2,4,10"/>
                      <Button x:Name="BtnNewKV" Content="+ New" Grid.Column="1" Style="{StaticResource BtnSec}" Padding="8,6" Margin="0,2,0,10" FontSize="11"/>
                    </Grid>
                    <Grid>
                      <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition/></Grid.ColumnDefinitions>
                      <StackPanel>
                        <Label Content="Local Admin Username *"/>
                        <TextBox x:Name="WizAdminUser" Text="avdadmin" Margin="0,2,8,10"/>
                      </StackPanel>
                      <StackPanel Grid.Column="1">
                        <Label Content="Local Admin Password *"/>
                        <PasswordBox x:Name="WizAdminPass" Background="#0D1F36" Foreground="#E2E8F0" BorderBrush="#1E3A5F" BorderThickness="1" Padding="8,6" FontSize="12" Margin="0,2,0,10"/>
                      </StackPanel>
                    </Grid>
                  </StackPanel>
                </Border>
              </StackPanel>
            </ScrollViewer>
            <!-- Step 4: Session Hosts -->
            <ScrollViewer x:Name="WS4" Padding="22" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="Session Host Virtual Machines" FontSize="20" FontWeight="SemiBold" Margin="0,0,0,2"/>
                <TextBlock Text="Configure the VMs that will serve user sessions" FontSize="12" Foreground="#64748B" Margin="0,0,0,16"/>
                <Border Style="{StaticResource Card}" Margin="0,0,0,10">
                  <StackPanel>
                    <TextBlock Text="VM Configuration" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,12"/>
                    <Grid>
                      <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition/></Grid.ColumnDefinitions>
                      <StackPanel>
                        <Label Content="Number of VMs *"/>
                        <StackPanel Orientation="Horizontal" Margin="0,2,8,10">
                          <Slider x:Name="SliderVMCnt" Minimum="1" Maximum="50" Value="2" Width="180" VerticalAlignment="Center"/>
                          <TextBlock x:Name="TxtVMCnt" Text="2" FontSize="18" FontWeight="Bold" Foreground="#50ABF1" Margin="10,0,0,0" VerticalAlignment="Center"/>
                        </StackPanel>
                      </StackPanel>
                      <StackPanel Grid.Column="1">
                        <Label Content="VM Name Prefix *"/>
                        <TextBox x:Name="WizVMPrefix" Text="avd-host" Margin="0,2,0,10" ToolTip="VMs will be named prefix-0, prefix-1, ..."/>
                      </StackPanel>
                    </Grid>
                    <Label Content="VM Size *"/>
                    <ComboBox x:Name="WizVMSize" Margin="0,2,0,4"/>
                    <TextBlock x:Name="TxtVMDetail" Text="" FontSize="11" Foreground="#64748B" Margin="0,0,0,10"/>
                    <Grid>
                      <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition/></Grid.ColumnDefinitions>
                      <StackPanel>
                        <Label Content="OS Disk Type"/>
                        <ComboBox x:Name="WizDiskType" Margin="0,2,8,10">
                          <ComboBoxItem Content="Premium SSD (recommended)" IsSelected="True"/>
                          <ComboBoxItem Content="Standard SSD"/>
                          <ComboBoxItem Content="Standard HDD"/>
                        </ComboBox>
                      </StackPanel>
                      <StackPanel Grid.Column="1">
                        <Label Content="Availability"/>
                        <ComboBox x:Name="WizAvail" Margin="0,2,0,10">
                          <ComboBoxItem Content="No infrastructure redundancy" IsSelected="True"/>
                          <ComboBoxItem Content="Availability Zone"/>
                          <ComboBoxItem Content="Availability Set"/>
                        </ComboBox>
                      </StackPanel>
                    </Grid>
                    <CheckBox x:Name="ChkHybridBenefit" Content="Use Azure Hybrid Benefit (Windows Server SA)" Margin="0,4"/>
                  </StackPanel>
                </Border>
                <Border Style="{StaticResource Card}">
                  <StackPanel>
                    <TextBlock Text="OS Image Source" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,10"/>
                    <RadioButton x:Name="RdoMarket" Content="Azure Marketplace" GroupName="Img" IsChecked="True" Margin="0,4"/>
                    <ComboBox x:Name="WizMktImg" Margin="4,4,0,10"/>
                    <RadioButton x:Name="RdoGallery" Content="Azure Compute Gallery (custom golden image)" GroupName="Img" Margin="0,4"/>
                    <Grid Margin="4,4,0,0">
                      <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition/></Grid.ColumnDefinitions>
                      <ComboBox x:Name="WizGallery" IsEnabled="False" Margin="0,0,8,0"/>
                      <ComboBox x:Name="WizGalleryImg" Grid.Column="1" IsEnabled="False"/>
                    </Grid>
                  </StackPanel>
                </Border>
              </StackPanel>
            </ScrollViewer>
            <!-- Step 5: Identity -->
            <ScrollViewer x:Name="WS5" Padding="22" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="Identity and Domain Join" FontSize="20" FontWeight="SemiBold" Margin="0,0,0,2"/>
                <TextBlock Text="Choose how session hosts are joined to a directory for user authentication" FontSize="12" Foreground="#64748B" Margin="0,0,0,16"/>
                <Border Style="{StaticResource Card}" Margin="0,0,0,10">
                  <StackPanel>
                    <TextBlock Text="Join Type" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,10"/>
                    <RadioButton x:Name="RdoEntra"  Content="Microsoft Entra ID Join (cloud-only, recommended for new deployments)" GroupName="Join" IsChecked="True" Margin="0,4"/>
                    <Border Background="#0D2547" CornerRadius="6" Padding="12,8" Margin="24,2,0,10">
                      <StackPanel>
                        <TextBlock Text="Requires: Entra ID P1 license, Intune for device management" FontSize="11" Foreground="#94A3B8"/>
                        <TextBlock Text="Users sign in with their Entra ID (Microsoft 365) credentials" FontSize="11" Foreground="#64748B" Margin="0,2"/>
                      </StackPanel>
                    </Border>
                    <RadioButton x:Name="RdoHybrid" Content="Hybrid Active Directory Join (on-premises AD required)" GroupName="Join" Margin="0,4"/>
                    <Border Background="#0D2547" CornerRadius="6" Padding="12,8" Margin="24,2,0,0">
                      <StackPanel>
                        <TextBlock Text="Requires: AD domain with line-of-sight from VNet, Azure AD Connect sync" FontSize="11" Foreground="#94A3B8"/>
                        <TextBlock Text="VMs join on-premises AD and are also registered in Entra ID" FontSize="11" Foreground="#64748B" Margin="0,2"/>
                      </StackPanel>
                    </Border>
                  </StackPanel>
                </Border>
                <Border x:Name="EntraPanel" Style="{StaticResource Card}" Margin="0,0,0,10">
                  <StackPanel>
                    <TextBlock Text="Entra ID Join Settings" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,10"/>
                    <CheckBox x:Name="ChkIntune" Content="Enroll VMs in Microsoft Intune (required for full device management and compliance policies)" IsChecked="True" Margin="0,4"/>
                    <Separator Margin="0,8"/>
                    <TextBlock Text="RBAC assignments required on the resource group:" FontSize="11" Foreground="#64748B" Margin="0,0,0,6"/>
                    <TextBlock Text="  - Virtual Machine User Login (AVD users)" FontSize="11" Foreground="#475569" Margin="0,2"/>
                    <TextBlock Text="  - Virtual Machine Administrator Login (AVD admins)" FontSize="11" Foreground="#475569" Margin="0,2"/>
                    <TextBlock Text="  - Desktop Virtualization Power On Contributor (AVD service)" FontSize="11" Foreground="#475569" Margin="0,2"/>
                  </StackPanel>
                </Border>
                <Border x:Name="HybridPanel" Style="{StaticResource Card}" Visibility="Collapsed">
                  <StackPanel>
                    <TextBlock Text="Hybrid AD Join Settings" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,12"/>
                    <Label Content="Domain FQDN *"/>
                    <TextBox x:Name="WizDomain" Margin="0,2,0,10" ToolTip="e.g. corp.contoso.com"/>
                    <Label Content="Organizational Unit (OU) - leave blank for default Computers container"/>
                    <TextBox x:Name="WizOU" Margin="0,2,0,10" ToolTip="e.g. OU=AVDHosts,OU=Servers,DC=corp,DC=contoso,DC=com"/>
                    <Grid>
                      <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition/></Grid.ColumnDefinitions>
                      <StackPanel>
                        <Label Content="Domain Join Account (UPN or domain\user) *"/>
                        <TextBox x:Name="WizDomainUser" Margin="0,2,8,10"/>
                      </StackPanel>
                      <StackPanel Grid.Column="1">
                        <Label Content="Domain Join Password *"/>
                        <PasswordBox x:Name="WizDomainPass" Background="#0D1F36" Foreground="#E2E8F0" BorderBrush="#1E3A5F" BorderThickness="1" Padding="8,6" FontSize="12" Margin="0,2,0,10"/>
                      </StackPanel>
                    </Grid>
                    <Label Content="Domain Controller IP(s) for VNet DNS (comma-separated)"/>
                    <TextBox x:Name="WizDCIPs" Margin="0,2,0,0" ToolTip="e.g. 10.0.0.4, 10.0.0.5"/>
                  </StackPanel>
                </Border>
              </StackPanel>
            </ScrollViewer>
            <!-- Step 6: FSLogix -->
            <ScrollViewer x:Name="WS6" Padding="22" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="User Profiles - FSLogix" FontSize="20" FontWeight="SemiBold" Margin="0,0,0,2"/>
                <TextBlock Text="Configure profile containers for persistent user data across pooled sessions" FontSize="12" Foreground="#64748B" Margin="0,0,0,16"/>
                <Border Style="{StaticResource Card}" Margin="0,0,0,10">
                  <StackPanel>
                    <CheckBox x:Name="ChkFSL" Content="Enable FSLogix Profile Containers" IsChecked="True" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,12"/>
                    <Label Content="Profile Storage Backend"/>
                    <ComboBox x:Name="WizFSLBackend" Margin="0,2,0,10">
                      <ComboBoxItem Content="Azure Files Premium (recommended for AVD - low latency)" IsSelected="True"/>
                      <ComboBoxItem Content="Azure Files Standard"/>
                      <ComboBoxItem Content="Azure NetApp Files (enterprise, high performance)"/>
                      <ComboBoxItem Content="Custom SMB share (on-premises or third-party)"/>
                    </ComboBox>
                    <Grid>
                      <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition/></Grid.ColumnDefinitions>
                      <StackPanel>
                        <Label Content="Storage Account"/>
                        <Grid>
                          <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                          <ComboBox x:Name="WizSA" Margin="0,2,4,10"/>
                          <Button x:Name="BtnNewSA" Content="+ New" Grid.Column="1" Style="{StaticResource BtnSec}" Padding="8,6" Margin="0,2,0,10" FontSize="11"/>
                        </Grid>
                      </StackPanel>
                      <StackPanel Grid.Column="1">
                        <Label Content="File Share Name"/>
                        <TextBox x:Name="WizShareName" Text="profiles" Margin="0,2,0,10"/>
                      </StackPanel>
                    </Grid>
                    <Grid>
                      <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition/></Grid.ColumnDefinitions>
                      <StackPanel>
                        <Label Content="Max Profile Size (GB per user)"/>
                        <StackPanel Orientation="Horizontal" Margin="0,2,8,10">
                          <Slider x:Name="SliderProfGB" Minimum="5" Maximum="100" Value="30" Width="160" VerticalAlignment="Center"/>
                          <TextBlock x:Name="TxtProfGB" Text="30 GB" FontSize="13" FontWeight="Bold" Foreground="#50ABF1" Margin="8,0,0,0" VerticalAlignment="Center"/>
                        </StackPanel>
                      </StackPanel>
                      <StackPanel Grid.Column="1">
                        <Label Content="Container Type"/>
                        <ComboBox x:Name="WizFSLType" Margin="0,2,0,10">
                          <ComboBoxItem Content="Profile Container only" IsSelected="True"/>
                          <ComboBoxItem Content="Office Container (ODFC) only"/>
                          <ComboBoxItem Content="Profile + Office Container (recommended)"/>
                        </ComboBox>
                      </StackPanel>
                    </Grid>
                    <CheckBox x:Name="ChkCloudCache" Content="Enable Cloud Cache (multi-region profile redundancy)" Margin="0,4"/>
                    <CheckBox x:Name="ChkKFM" Content="Enable OneDrive Known Folder Move (redirect Desktop, Documents, Pictures to OneDrive)" IsChecked="True" Margin="0,4"/>
                  </StackPanel>
                </Border>
              </StackPanel>
            </ScrollViewer>
            <!-- Step 7: RDP and Scaling -->
            <ScrollViewer x:Name="WS7" Padding="22" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="RDP Properties and Auto-Scaling" FontSize="20" FontWeight="SemiBold" Margin="0,0,0,2"/>
                <TextBlock Text="Configure session behavior and automated host pool capacity management" FontSize="12" Foreground="#64748B" Margin="0,0,0,16"/>
                <Border Style="{StaticResource Card}" Margin="0,0,0,10">
                  <StackPanel>
                    <TextBlock Text="Security Preset" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,10"/>
                    <ComboBox x:Name="WizRDPPreset" Margin="0,0,0,10">
                      <ComboBoxItem Content="Balanced (clipboard on, drives off, cameras on) - recommended for office workers" IsSelected="True"/>
                      <ComboBoxItem Content="Strict / PCI (clipboard off, screen capture blocked, no redirection)"/>
                      <ComboBoxItem Content="Kiosk / Frontline (minimal redirection, single monitor)"/>
                      <ComboBoxItem Content="Developer / Admin (full redirection, no restrictions)"/>
                    </ComboBox>
                    <Grid>
                      <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition/><ColumnDefinition/></Grid.ColumnDefinitions>
                      <CheckBox x:Name="WizChkClip"   Content="Clipboard"              IsChecked="True" Margin="0,3,8,3"/>
                      <CheckBox x:Name="WizChkDrives" Content="Drive redirection"      Grid.Column="1" Margin="0,3,8,3"/>
                      <CheckBox x:Name="WizChkCam"    Content="Camera/mic"             IsChecked="True" Grid.Column="2" Margin="0,3"/>
                    </Grid>
                    <Grid>
                      <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition/><ColumnDefinition/></Grid.ColumnDefinitions>
                      <CheckBox x:Name="WizChkUSB"    Content="USB redirection"        Margin="0,3,8,3"/>
                      <CheckBox x:Name="WizChkWM"     Content="Watermarking"           IsChecked="True" Grid.Column="1" Margin="0,3,8,3"/>
                      <CheckBox x:Name="WizChkSCP"    Content="Screen capture protect" IsChecked="True" Grid.Column="2" Margin="0,3"/>
                    </Grid>
                  </StackPanel>
                </Border>
                <Border Style="{StaticResource Card}" Margin="0,0,0,10">
                  <StackPanel>
                    <CheckBox x:Name="ChkScaling" Content="Enable Auto-Scaling Plan" IsChecked="True" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,12"/>
                    <Grid>
                      <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition/></Grid.ColumnDefinitions>
                      <StackPanel>
                        <Label Content="Ramp-Up Start (morning)"/>
                        <ComboBox x:Name="WizRampUp" Margin="0,2,8,10">
                          <ComboBoxItem Content="06:00 AM"/><ComboBoxItem Content="06:30 AM"/>
                          <ComboBoxItem Content="07:00 AM" IsSelected="True"/><ComboBoxItem Content="07:30 AM"/><ComboBoxItem Content="08:00 AM"/>
                        </ComboBox>
                        <Label Content="Peak Start"/>
                        <ComboBox x:Name="WizPeakStart" Margin="0,2,8,10">
                          <ComboBoxItem Content="08:00 AM"/>
                          <ComboBoxItem Content="09:00 AM" IsSelected="True"/><ComboBoxItem Content="09:30 AM"/>
                        </ComboBox>
                      </StackPanel>
                      <StackPanel Grid.Column="1">
                        <Label Content="Ramp-Down Start (evening)"/>
                        <ComboBox x:Name="WizRampDown" Margin="0,2,0,10">
                          <ComboBoxItem Content="05:00 PM"/><ComboBoxItem Content="05:30 PM"/>
                          <ComboBoxItem Content="06:00 PM" IsSelected="True"/><ComboBoxItem Content="07:00 PM"/>
                        </ComboBox>
                        <Label Content="Off-Peak (night)"/>
                        <ComboBox x:Name="WizOffPeak" Margin="0,2,0,10">
                          <ComboBoxItem Content="09:00 PM"/>
                          <ComboBoxItem Content="10:00 PM" IsSelected="True"/><ComboBoxItem Content="11:00 PM"/>
                        </ComboBox>
                      </StackPanel>
                    </Grid>
                    <Grid>
                      <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition/></Grid.ColumnDefinitions>
                      <StackPanel>
                        <Label Content="Capacity Threshold (%)"/>
                        <StackPanel Orientation="Horizontal" Margin="0,2,8,0">
                          <Slider x:Name="SliderCap" Minimum="50" Maximum="95" Value="80" Width="160" VerticalAlignment="Center"/>
                          <TextBlock x:Name="TxtCap" Text="80%" FontSize="13" FontWeight="Bold" Foreground="#50ABF1" Margin="8,0,0,0" VerticalAlignment="Center"/>
                        </StackPanel>
                      </StackPanel>
                      <StackPanel Grid.Column="1">
                        <Label Content="Minimum Hosts (off-peak)"/>
                        <StackPanel Orientation="Horizontal" Margin="0,2,0,0">
                          <Slider x:Name="SliderMinH" Minimum="0" Maximum="20" Value="1" Width="160" VerticalAlignment="Center"/>
                          <TextBlock x:Name="TxtMinH" Text="1" FontSize="13" FontWeight="Bold" Foreground="#50ABF1" Margin="8,0,0,0" VerticalAlignment="Center"/>
                        </StackPanel>
                      </StackPanel>
                    </Grid>
                  </StackPanel>
                </Border>
                <Border Style="{StaticResource Card}">
                  <StackPanel>
                    <TextBlock Text="Monitoring" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,10"/>
                    <CheckBox x:Name="ChkMonitor" Content="Enable Azure Monitor / AVD Insights (sends diagnostics to Log Analytics)" IsChecked="True" Margin="0,4"/>
                    <Label Content="Log Analytics Workspace (optional - recommended)"/>
                    <Grid>
                      <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                      <ComboBox x:Name="WizLAW" Margin="0,2,4,0"/>
                      <Button x:Name="BtnNewLAW" Content="+ New" Grid.Column="1" Style="{StaticResource BtnSec}" Padding="8,6" Margin="0,2,0,0" FontSize="11"/>
                    </Grid>
                  </StackPanel>
                </Border>
              </StackPanel>
            </ScrollViewer>
            <!-- Step 8: Review and Deploy -->
            <ScrollViewer x:Name="WS8" Padding="22" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="Review and Deploy" FontSize="20" FontWeight="SemiBold" Margin="0,0,0,2"/>
                <TextBlock Text="Review your configuration then click Deploy to create resources" FontSize="12" Foreground="#64748B" Margin="0,0,0,16"/>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="200"/></Grid.ColumnDefinitions>
                  <StackPanel Grid.Column="0" Margin="0,0,10,0">
                    <Border Style="{StaticResource Card}" Margin="0,0,0,10">
                      <DockPanel>
                        <TextBlock DockPanel.Dock="Top" Text="Configuration Summary" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,10"/>
                        <DataGrid x:Name="GridReview" Height="220" CanUserSortColumns="False" HeadersVisibility="None">
                          <DataGrid.Columns>
                            <DataGridTextColumn Binding="{Binding Section}" Width="140">
                              <DataGridTextColumn.ElementStyle><Style TargetType="TextBlock"><Setter Property="Foreground" Value="#64748B"/><Setter Property="FontWeight" Value="SemiBold"/></Style></DataGridTextColumn.ElementStyle>
                            </DataGridTextColumn>
                            <DataGridTextColumn Binding="{Binding Setting}" Width="180"/>
                            <DataGridTextColumn Binding="{Binding Value}"   Width="*"/>
                          </DataGrid.Columns>
                        </DataGrid>
                      </DockPanel>
                    </Border>
                    <Border Style="{StaticResource Card}">
                      <DockPanel>
                        <TextBlock DockPanel.Dock="Top" Text="Deployment Log" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,10"/>
                        <Border Background="#060E1A" CornerRadius="6" BorderBrush="#1E3A5F" BorderThickness="1" Padding="4">
                          <RichTextBox x:Name="WizLog" Height="160" Background="Transparent" Foreground="#94A3B8"
                                       BorderThickness="0" IsReadOnly="True" FontFamily="Cascadia Code, Consolas" FontSize="11"
                                       VerticalScrollBarVisibility="Auto"/>
                        </Border>
                      </DockPanel>
                    </Border>
                  </StackPanel>
                  <StackPanel Grid.Column="1">
                    <Border Style="{StaticResource Card}" Margin="0,0,0,8">
                      <StackPanel>
                        <TextBlock Text="Est. Monthly Cost" FontSize="10" Foreground="#64748B"/>
                        <TextBlock x:Name="TxtEstCost" Text="--" FontSize="22" FontWeight="Bold" Foreground="#50ABF1" Margin="0,4"/>
                        <TextBlock x:Name="TxtEstNote" Text="(compute only)" FontSize="10" Foreground="#334155"/>
                      </StackPanel>
                    </Border>
                    <ProgressBar x:Name="DeployProg" Maximum="100" Value="0" Height="5" Margin="0,0,0,4"/>
                    <TextBlock x:Name="TxtDeployPhase" Text="Ready to deploy" FontSize="11" Foreground="#64748B" Margin="0,0,0,8"/>
                    <Button x:Name="BtnDeploy" Content="Deploy" Style="{StaticResource BtnGreen}" Padding="0,12" Margin="0,0,0,6" FontSize="14"/>
                    <Button x:Name="BtnCancelDeploy" Content="Cancel Deploy" Style="{StaticResource BtnRed}" Padding="0,9" IsEnabled="False"/>
                  </StackPanel>
                </Grid>
              </StackPanel>
            </ScrollViewer>
            <!-- Wizard nav bar -->
            <Border VerticalAlignment="Bottom" Background="#07101E" BorderBrush="#1E3A5F" BorderThickness="0,1,0,0" Padding="22,10">
              <Grid>
                <Button x:Name="BtnWizBack" Content="Back" Style="{StaticResource BtnSec}" HorizontalAlignment="Left" Padding="20,8" IsEnabled="False"/>
                <StackPanel HorizontalAlignment="Right" Orientation="Horizontal">
                  <Button x:Name="BtnWizCancel" Content="Cancel" Style="{StaticResource BtnSec}" Padding="14,8" Margin="0,0,8,0"/>
                  <Button x:Name="BtnWizNext" Content="Next" Style="{StaticResource BtnPrimary}" Padding="24,8"/>
                </StackPanel>
              </Grid>
            </Border>
          </Grid>
        </Grid>

        <!-- ===== HOST POOLS ===== -->
        <ScrollViewer x:Name="PanelHP" Visibility="Collapsed" Padding="20">
          <StackPanel>
            <Grid Margin="0,0,0,14">
              <TextBlock Text="Host Pools" FontSize="22" FontWeight="SemiBold" VerticalAlignment="Bottom"/>
              <StackPanel HorizontalAlignment="Right" VerticalAlignment="Bottom" Orientation="Horizontal">
                <Button x:Name="BtnRefHP" Content="Refresh" Style="{StaticResource BtnSec}" Padding="12,7" Margin="4,0"/>
                <Button x:Name="BtnNewHP" Content="+ New Host Pool" Style="{StaticResource BtnPrimary}" Padding="14,7" Margin="4,0"/>
              </StackPanel>
            </Grid>
            <Grid Margin="0,0,0,8">
              <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
              <TextBox x:Name="FilterHP" Margin="0,0,8,0" FontSize="12" ToolTip="Filter host pools by name, type, or region..." />
              <Button x:Name="BtnExportHP" Grid.Column="1" Content="Export CSV" Style="{StaticResource BtnSec}" Padding="10,6" FontSize="11"/>
            </Grid>
            <Border Style="{StaticResource Card}">
              <DataGrid x:Name="GridHP" CanUserSortColumns="True" Height="560">
                <DataGrid.Columns>
                  <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="200"/>
                  <DataGridTemplateColumn Header="Type" Width="110">
                    <DataGridTemplateColumn.CellTemplate><DataTemplate>
                      <Border CornerRadius="4" Padding="6,2" Background="{Binding TypeBg}" HorizontalAlignment="Left">
                        <TextBlock Text="{Binding Type}" FontSize="10" FontWeight="SemiBold" Foreground="{Binding TypeFg}"/>
                      </Border>
                    </DataTemplate></DataGridTemplateColumn.CellTemplate>
                  </DataGridTemplateColumn>
                  <DataGridTextColumn Header="Load Balancing"  Binding="{Binding LB}"         Width="130"/>
                  <DataGridTextColumn Header="Max Sessions"    Binding="{Binding MaxSess}"     Width="100"/>
                  <DataGridTextColumn Header="Hosts (Avail)"  Binding="{Binding Hosts}"       Width="100"/>
                  <DataGridTextColumn Header="Sessions"        Binding="{Binding Sessions}"    Width="80"/>
                  <DataGridTextColumn Header="Assignment"      Binding="{Binding Assignment}"  Width="100"/>
                  <DataGridTemplateColumn Header="Status" Width="100">
                    <DataGridTemplateColumn.CellTemplate><DataTemplate>
                      <Border CornerRadius="4" Padding="6,2" Background="{Binding StatusBg}" HorizontalAlignment="Left">
                        <TextBlock Text="{Binding Status}" FontSize="10" FontWeight="SemiBold" Foreground="{Binding StatusFg}"/>
                      </Border>
                    </DataTemplate></DataGridTemplateColumn.CellTemplate>
                  </DataGridTemplateColumn>
                  <DataGridTextColumn Header="Region"         Binding="{Binding Location}"    Width="110"/>
                  <DataGridTextColumn Header="Resource Group" Binding="{Binding RG}"          Width="*"/>
                </DataGrid.Columns>
                <DataGrid.ContextMenu>
                  <ContextMenu Background="#112240" BorderBrush="#1E3A5F">
                    <MenuItem x:Name="CtxHPAddHost"  Header="Add Session Host"    Foreground="#50ABF1"/>
                    <MenuItem x:Name="CtxHPRDP"      Header="Edit RDP Properties" Foreground="#E2E8F0"/>
                    <Separator Background="#1E3A5F"/>
                    <MenuItem x:Name="CtxHPPortal"   Header="Open in Portal"      Foreground="#50ABF1"/>
                    <Separator Background="#1E3A5F"/>
                    <MenuItem x:Name="CtxHPDelete"   Header="Delete"              Foreground="#FCA5A5"/>
                  </ContextMenu>
                </DataGrid.ContextMenu>
              </DataGrid>
            </Border>
          </StackPanel>
        </ScrollViewer>

        <!-- ===== SESSION HOSTS ===== -->
        <ScrollViewer x:Name="PanelSess" Visibility="Collapsed" Padding="20">
          <StackPanel>
            <Grid Margin="0,0,0,14">
              <StackPanel>
                <TextBlock Text="Session Hosts" FontSize="22" FontWeight="SemiBold"/>
                <StackPanel Orientation="Horizontal" Margin="0,4,0,0">
                  <TextBlock Text="Filter:" FontSize="12" Foreground="#64748B" VerticalAlignment="Center" Margin="0,0,8,0"/>
                  <ComboBox x:Name="SessFilter" Width="260" FontSize="12"/>
                </StackPanel>
              </StackPanel>
              <StackPanel HorizontalAlignment="Right" VerticalAlignment="Bottom" Orientation="Horizontal">
                <TextBox x:Name="FilterSess" Width="240" Margin="0,0,8,0" FontSize="12" VerticalAlignment="Center" ToolTip="Filter by VM name, pool, or status..."/>
                <Button x:Name="BtnExportSess" Content="Export CSV" Style="{StaticResource BtnSec}" Padding="10,7" Margin="4,0"/>
                <Button x:Name="BtnHealAll"   Content="Heal Unhealthy"  Style="{StaticResource BtnSec}" Padding="12,7" Margin="4,0"/>
                <Button x:Name="BtnDrainAll"  Content="Drain All"       Style="{StaticResource BtnSec}" Padding="12,7" Margin="4,0"/>
                <Button x:Name="BtnRefSess"   Content="Refresh"         Style="{StaticResource BtnSec}" Padding="12,7" Margin="4,0"/>
              </StackPanel>
            </Grid>
            <Border Style="{StaticResource Card}">
              <DataGrid x:Name="GridSess" CanUserSortColumns="True" Height="520">
                <DataGrid.Columns>
                  <DataGridTextColumn Header="VM Name"      Binding="{Binding VMName}"   Width="200"/>
                  <DataGridTextColumn Header="Host Pool"    Binding="{Binding Pool}"     Width="180"/>
                  <DataGridTemplateColumn Header="Status" Width="120">
                    <DataGridTemplateColumn.CellTemplate><DataTemplate>
                      <Border CornerRadius="4" Padding="6,2" Background="{Binding StatusBg}" HorizontalAlignment="Left">
                        <TextBlock Text="{Binding Status}" FontSize="10" FontWeight="SemiBold" Foreground="{Binding StatusFg}"/>
                      </Border>
                    </DataTemplate></DataGridTemplateColumn.CellTemplate>
                  </DataGridTemplateColumn>
                  <DataGridTextColumn Header="Sessions"     Binding="{Binding Sessions}" Width="80"/>
                  <DataGridTextColumn Header="VM Size"      Binding="{Binding VMSize}"   Width="160"/>
                  <DataGridTemplateColumn Header="Drain" Width="90">
                    <DataGridTemplateColumn.CellTemplate><DataTemplate>
                      <Border CornerRadius="4" Padding="6,2" Background="{Binding DrainBg}" HorizontalAlignment="Left">
                        <TextBlock Text="{Binding DrainText}" FontSize="10" Foreground="{Binding DrainFg}"/>
                      </Border>
                    </DataTemplate></DataGridTemplateColumn.CellTemplate>
                  </DataGridTemplateColumn>
                  <DataGridTextColumn Header="OS"           Binding="{Binding OSVer}"    Width="120"/>
                  <DataGridTextColumn Header="Last Heartbeat" Binding="{Binding Heartbeat}" Width="*"/>
                </DataGrid.Columns>
                <DataGrid.ContextMenu>
                  <ContextMenu Background="#112240" BorderBrush="#1E3A5F">
                    <MenuItem x:Name="CtxSessEnableDrain"  Header="Enable Drain Mode"  Foreground="#E2E8F0"/>
                    <MenuItem x:Name="CtxSessDisableDrain" Header="Disable Drain Mode" Foreground="#E2E8F0"/>
                    <Separator Background="#1E3A5F"/>
                    <MenuItem x:Name="CtxSessRemove"  Header="Remove from Pool" Foreground="#FCA5A5"/>
                    <Separator Background="#1E3A5F"/>
                    <MenuItem x:Name="CtxSessPortal"  Header="Open in Portal"   Foreground="#50ABF1"/>
                  </ContextMenu>
                </DataGrid.ContextMenu>
              </DataGrid>
            </Border>
            <!-- Active User Sessions sub-panel -->
            <Border Style="{StaticResource Card}" Margin="0,10,0,0">
              <DockPanel>
                <Grid DockPanel.Dock="Top" Margin="0,0,0,8">
                  <TextBlock Text="Active User Sessions" FontSize="14" FontWeight="SemiBold"/>
                  <StackPanel HorizontalAlignment="Right" Orientation="Horizontal">
                    <Button x:Name="BtnRefUserSess"    Content="Refresh"       Style="{StaticResource BtnSec}" Padding="10,5" Margin="4,0" FontSize="11"/>
                    <Button x:Name="BtnMsgUser"        Content="Send Message"  Style="{StaticResource BtnSec}" Padding="10,5" Margin="4,0" FontSize="11"/>
                    <Button x:Name="BtnDisconnectUser" Content="Disconnect"    Style="{StaticResource BtnSec}" Padding="10,5" Margin="4,0" FontSize="11"/>
                    <Button x:Name="BtnLogoffUser"     Content="Force Logoff"  Style="{StaticResource BtnRed}"  Padding="10,5" Margin="4,0" FontSize="11"/>
                    <Button x:Name="BtnExportUserSess" Content="Export CSV"    Style="{StaticResource BtnSec}" Padding="10,5" Margin="4,0" FontSize="11"/>
                  </StackPanel>
                </Grid>
                <DataGrid x:Name="GridUserSess" Height="190" CanUserSortColumns="True">
                  <DataGrid.Columns>
                    <DataGridTextColumn Header="User UPN"      Binding="{Binding UserUPN}"   Width="230"/>
                    <DataGridTextColumn Header="Host Pool"     Binding="{Binding PoolName}"  Width="160"/>
                    <DataGridTextColumn Header="Session Host"  Binding="{Binding HostName}"  Width="180"/>
                    <DataGridTextColumn Header="Session ID"    Binding="{Binding SessionId}" Width="85"/>
                    <DataGridTemplateColumn Header="State" Width="100">
                      <DataGridTemplateColumn.CellTemplate><DataTemplate>
                        <Border CornerRadius="4" Padding="6,2" Background="{Binding StateBg}" HorizontalAlignment="Left">
                          <TextBlock Text="{Binding State}" FontSize="10" FontWeight="SemiBold" Foreground="{Binding StateFg}"/>
                        </Border>
                      </DataTemplate></DataGridTemplateColumn.CellTemplate>
                    </DataGridTemplateColumn>
                    <DataGridTextColumn Header="Application"   Binding="{Binding App}"       Width="*"/>
                    <DataGridTextColumn Header="Login Time"    Binding="{Binding LoginTime}"  Width="130"/>
                  </DataGrid.Columns>
                </DataGrid>
              </DockPanel>
            </Border>
          </StackPanel>
        </ScrollViewer>

        <!-- ===== APP GROUPS ===== -->
        <ScrollViewer x:Name="PanelAG" Visibility="Collapsed" Padding="20">
          <StackPanel>
            <Grid Margin="0,0,0,14">
              <TextBlock Text="Application Groups" FontSize="22" FontWeight="SemiBold" VerticalAlignment="Bottom"/>
              <StackPanel HorizontalAlignment="Right" VerticalAlignment="Bottom" Orientation="Horizontal">
                <Button x:Name="BtnRefAG" Content="Refresh" Style="{StaticResource BtnSec}" Padding="12,7" Margin="4,0"/>
                <Button x:Name="BtnNewAG" Content="+ New App Group" Style="{StaticResource BtnPrimary}" Padding="14,7" Margin="4,0"/>
              </StackPanel>
            </Grid>
            <Grid>
              <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="320"/></Grid.ColumnDefinitions>
              <Border Grid.Column="0" Style="{StaticResource Card}" Margin="0,0,8,0">
                <DockPanel>
                  <TextBlock DockPanel.Dock="Top" Text="App Groups" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,10"/>
                  <DataGrid x:Name="GridAG" CanUserSortColumns="True">
                    <DataGrid.Columns>
                      <DataGridTextColumn Header="Name"      Binding="{Binding Name}"       Width="200"/>
                      <DataGridTemplateColumn Header="Type" Width="100">
                        <DataGridTemplateColumn.CellTemplate><DataTemplate>
                          <Border CornerRadius="4" Padding="6,2" Background="{Binding TypeBg}" HorizontalAlignment="Left">
                            <TextBlock Text="{Binding Type}" FontSize="10" FontWeight="SemiBold" Foreground="{Binding TypeFg}"/>
                          </Border>
                        </DataTemplate></DataGridTemplateColumn.CellTemplate>
                      </DataGridTemplateColumn>
                      <DataGridTextColumn Header="Host Pool"   Binding="{Binding HostPool}"   Width="*"/>
                      <DataGridTextColumn Header="Apps"        Binding="{Binding AppCount}"   Width="60"/>
                      <DataGridTextColumn Header="Users"       Binding="{Binding UserCount}"  Width="70"/>
                    </DataGrid.Columns>
                  </DataGrid>
                </DockPanel>
              </Border>
              <Border Grid.Column="1" Style="{StaticResource Card}" Margin="8,0,0,0">
                <DockPanel>
                  <Grid DockPanel.Dock="Top" Margin="0,0,0,10">
                    <TextBlock Text="Published Apps" FontSize="14" FontWeight="SemiBold"/>
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                      <Button x:Name="BtnAssignUsers" Content="Assign Users" Style="{StaticResource BtnSec}" Padding="10,5" Margin="0,0,6,0" FontSize="11"/>
                      <Button x:Name="BtnPubApp" Content="+ Publish App" Style="{StaticResource BtnPrimary}" Padding="10,5" FontSize="11"/>
                    </StackPanel>
                  </Grid>
                  <DataGrid x:Name="GridApps" CanUserSortColumns="False">
                    <DataGrid.Columns>
                      <DataGridTextColumn Header="App Name" Binding="{Binding AppName}" Width="*"/>
                      <DataGridTextColumn Header="Path"     Binding="{Binding AppPath}" Width="140"/>
                    </DataGrid.Columns>
                  </DataGrid>
                </DockPanel>
              </Border>
            </Grid>
          </StackPanel>
        </ScrollViewer>

        <!-- ===== FSLOGIX ===== -->
        <ScrollViewer x:Name="PanelFSL" Visibility="Collapsed" Padding="20">
          <StackPanel>
            <TextBlock Text="FSLogix Profile Management" FontSize="22" FontWeight="SemiBold" Margin="0,0,0,14"/>
            <Grid Margin="0,0,0,12">
              <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition/><ColumnDefinition/></Grid.ColumnDefinitions>
              <Border Grid.Column="0" Style="{StaticResource Card}" Margin="0,0,6,0">
                <StackPanel>
                  <TextBlock Text="PROFILES" FontSize="9" Foreground="#64748B" FontWeight="SemiBold"/>
                  <TextBlock x:Name="FSLTotal" Text="-" FontSize="30" FontWeight="Bold" Foreground="#50ABF1" Margin="0,4,0,2"/>
                  <TextBlock Text="VHDX containers" FontSize="11" Foreground="#475569"/>
                </StackPanel>
              </Border>
              <Border Grid.Column="1" Style="{StaticResource Card}" Margin="3,0">
                <StackPanel>
                  <TextBlock Text="TOTAL SIZE" FontSize="9" Foreground="#64748B" FontWeight="SemiBold"/>
                  <TextBlock x:Name="FSLSize" Text="-" FontSize="30" FontWeight="Bold" Foreground="#F59E0B" Margin="0,4,0,2"/>
                  <TextBlock Text="GB used" FontSize="11" Foreground="#475569"/>
                </StackPanel>
              </Border>
              <Border Grid.Column="2" Style="{StaticResource Card}" Margin="6,0,0,0">
                <StackPanel>
                  <TextBlock Text="QUOTA USED" FontSize="9" Foreground="#64748B" FontWeight="SemiBold"/>
                  <TextBlock x:Name="FSLQuota" Text="-" FontSize="30" FontWeight="Bold" Foreground="#10B981" Margin="0,4,0,2"/>
                  <TextBlock Text="of provisioned" FontSize="11" Foreground="#475569"/>
                </StackPanel>
              </Border>
            </Grid>
            <Grid>
              <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="260"/></Grid.ColumnDefinitions>
              <Border Grid.Column="0" Style="{StaticResource Card}" Margin="0,0,8,0">
                <DockPanel>
                  <TextBlock DockPanel.Dock="Top" Text="Profiles by Size" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,10"/>
                  <DataGrid x:Name="GridFSL" CanUserSortColumns="True">
                    <DataGrid.Columns>
                      <DataGridTextColumn Header="Username"     Binding="{Binding Username}" Width="160"/>
                      <DataGridTextColumn Header="Share"        Binding="{Binding Share}"    Width="110"/>
                      <DataGridTextColumn Header="Size (MB)"    Binding="{Binding SizeMB}"   Width="90"/>
                      <DataGridTemplateColumn Header="Health" Width="80">
                        <DataGridTemplateColumn.CellTemplate><DataTemplate>
                          <Border CornerRadius="4" Padding="6,2" Background="{Binding HBg}" HorizontalAlignment="Left">
                            <TextBlock Text="{Binding Health}" FontSize="10" FontWeight="SemiBold" Foreground="{Binding HFg}"/>
                          </Border>
                        </DataTemplate></DataGridTemplateColumn.CellTemplate>
                      </DataGridTemplateColumn>
                      <DataGridTextColumn Header="Last Mounted" Binding="{Binding LastMount}" Width="*"/>
                    </DataGrid.Columns>
                  </DataGrid>
                </DockPanel>
              </Border>
              <Border Grid.Column="1" Style="{StaticResource Card}" Margin="8,0,0,0">
                <StackPanel>
                  <TextBlock Text="Quick Actions" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,10"/>
                  <Button x:Name="BtnFSLLocks" Content="Clear .lock Files"    Style="{StaticResource BtnSec}" Margin="0,4" Padding="0,9"/>
                  <Button x:Name="BtnFSLTmp"   Content="Remove .tmp VHDXs"   Style="{StaticResource BtnSec}" Margin="0,4" Padding="0,9"/>
                  <Button x:Name="BtnFSLDiag"  Content="Run Diagnostics"     Style="{StaticResource BtnPrimary}" Margin="0,4" Padding="0,9"/>
                  <Separator Margin="0,8"/>
                  <TextBlock Text="Alert Threshold" FontSize="11" Foreground="#64748B"/>
                  <StackPanel Orientation="Horizontal" Margin="0,4">
                    <Slider x:Name="SliderFSLAlert" Minimum="1" Maximum="50" Value="10" Width="150" VerticalAlignment="Center"/>
                    <TextBlock x:Name="TxtFSLAlert" Text="10 GB" FontSize="12" Foreground="#50ABF1" Margin="8,0,0,0" VerticalAlignment="Center"/>
                  </StackPanel>
                  <Button x:Name="BtnFSLRef" Content="Refresh Profiles" Style="{StaticResource BtnSec}" Margin="0,8,0,0" Padding="0,9"/>
                </StackPanel>
              </Border>
            </Grid>
          </StackPanel>
        </ScrollViewer>

        <!-- ===== SCALING ===== -->
        <ScrollViewer x:Name="PanelScale" Visibility="Collapsed" Padding="20">
          <StackPanel>
            <Grid Margin="0,0,0,14">
              <TextBlock Text="Auto-Scaling Plans" FontSize="22" FontWeight="SemiBold" VerticalAlignment="Bottom"/>
              <StackPanel HorizontalAlignment="Right" VerticalAlignment="Bottom" Orientation="Horizontal">
                <Button x:Name="BtnRefScale" Content="Refresh" Style="{StaticResource BtnSec}" Padding="12,7" Margin="4,0"/>
                <Button x:Name="BtnNewScale" Content="+ New Scaling Plan" Style="{StaticResource BtnPrimary}" Padding="14,7" Margin="4,0"/>
              </StackPanel>
            </Grid>
            <Border Style="{StaticResource Card}">
              <DockPanel>
                <TextBlock DockPanel.Dock="Top" Text="Scaling Plans" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,10"/>
                <DataGrid x:Name="GridScale" CanUserSortColumns="True" Height="300">
                  <DataGrid.Columns>
                    <DataGridTextColumn Header="Plan Name"  Binding="{Binding Name}"      Width="200"/>
                    <DataGridTextColumn Header="Host Pool"  Binding="{Binding HP}"        Width="*"/>
                    <DataGridTemplateColumn Header="Status" Width="90">
                      <DataGridTemplateColumn.CellTemplate><DataTemplate>
                        <Border CornerRadius="4" Padding="6,2" Background="{Binding StatusBg}" HorizontalAlignment="Left">
                          <TextBlock Text="{Binding Status}" FontSize="10" FontWeight="SemiBold" Foreground="{Binding StatusFg}"/>
                        </Border>
                      </DataTemplate></DataGridTemplateColumn.CellTemplate>
                    </DataGridTemplateColumn>
                    <DataGridTextColumn Header="Time Zone"  Binding="{Binding TZ}"        Width="180"/>
                    <DataGridTextColumn Header="Schedules"  Binding="{Binding Schedules}"  Width="80"/>
                  </DataGrid.Columns>
                </DataGrid>
              </DockPanel>
            </Border>
          </StackPanel>
        </ScrollViewer>

        <!-- ===== MONITORING ===== -->
        <ScrollViewer x:Name="PanelMon" Visibility="Collapsed" Padding="20">
          <StackPanel>
            <TextBlock Text="Monitoring and Alerts" FontSize="22" FontWeight="SemiBold" Margin="0,0,0,14"/>
            <Border Style="{StaticResource Card}" Margin="0,0,0,10">
              <DockPanel>
                <TextBlock DockPanel.Dock="Top" Text="Log Analytics Workspaces" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,10"/>
                <DataGrid x:Name="GridLAW" Height="130" CanUserSortColumns="True">
                  <DataGrid.Columns>
                    <DataGridTextColumn Header="Workspace"     Binding="{Binding Name}"      Width="*"/>
                    <DataGridTextColumn Header="Resource Group" Binding="{Binding RG}"       Width="200"/>
                    <DataGridTextColumn Header="SKU"           Binding="{Binding SKU}"       Width="100"/>
                    <DataGridTextColumn Header="Retention"     Binding="{Binding Retention}" Width="100"/>
                  </DataGrid.Columns>
                </DataGrid>
              </DockPanel>
            </Border>
            <Border Style="{StaticResource Card}">
              <DockPanel>
                <Grid DockPanel.Dock="Top" Margin="0,0,0,10">
                  <TextBlock Text="Alert Rules" FontSize="14" FontWeight="SemiBold"/>
                  <Button x:Name="BtnAddAlert" Content="+ Add Alert" Style="{StaticResource BtnPrimary}" HorizontalAlignment="Right" Padding="12,5" FontSize="11"/>
                </Grid>
                <DataGrid x:Name="GridAlerts" Height="200" CanUserSortColumns="True">
                  <DataGrid.Columns>
                    <DataGridTextColumn Header="Alert Name" Binding="{Binding Name}"      Width="*"/>
                    <DataGridTextColumn Header="Condition"  Binding="{Binding Condition}" Width="220"/>
                    <DataGridTextColumn Header="Severity"   Binding="{Binding Severity}"  Width="80"/>
                    <DataGridTemplateColumn Header="Enabled" Width="80">
                      <DataGridTemplateColumn.CellTemplate><DataTemplate>
                        <Border CornerRadius="4" Padding="6,2" Background="{Binding EBg}" HorizontalAlignment="Left">
                          <TextBlock Text="{Binding Enabled}" FontSize="10" FontWeight="SemiBold" Foreground="{Binding EFg}"/>
                        </Border>
                      </DataTemplate></DataGridTemplateColumn.CellTemplate>
                    </DataGridTemplateColumn>
                  </DataGrid.Columns>
                </DataGrid>
              </DockPanel>
            </Border>
          </StackPanel>
        </ScrollViewer>

        <!-- ===== RDP AND SECURITY ===== -->
        <ScrollViewer x:Name="PanelRDP" Visibility="Collapsed" Padding="20">
          <StackPanel>
            <Grid Margin="0,0,0,14">
              <TextBlock Text="RDP Properties and Security" FontSize="22" FontWeight="SemiBold" VerticalAlignment="Bottom"/>
              <Button x:Name="BtnApplyRDP" Content="Apply to Pool" Style="{StaticResource BtnPrimary}" HorizontalAlignment="Right" VerticalAlignment="Bottom" Padding="16,8"/>
            </Grid>
            <Grid>
              <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="320"/></Grid.ColumnDefinitions>
              <Border Grid.Column="0" Style="{StaticResource Card}" Margin="0,0,8,0">
                <StackPanel>
                  <Grid Margin="0,0,0,12">
                    <TextBlock Text="Custom RDP Properties" FontSize="14" FontWeight="SemiBold"/>
                    <StackPanel HorizontalAlignment="Right" Orientation="Horizontal">
                      <TextBlock Text="Target pool: " FontSize="11" Foreground="#64748B" VerticalAlignment="Center" Margin="0,0,6,0"/>
                      <ComboBox x:Name="RDPPool" Width="220" FontSize="11"/>
                    </StackPanel>
                  </Grid>
                  <Label Content="Security Preset"/>
                  <ComboBox x:Name="RDPPreset" Margin="0,2,0,10">
                    <ComboBoxItem Content="Balanced (recommended for office workers)" IsSelected="True"/>
                    <ComboBoxItem Content="Strict / PCI (maximum security, no redirection)"/>
                    <ComboBoxItem Content="Kiosk / Frontline (minimal, single monitor)"/>
                    <ComboBoxItem Content="Developer / Admin (full access)"/>
                    <ComboBoxItem Content="Custom"/>
                  </ComboBox>
                  <Label Content="RDP Property String"/>
                  <TextBox x:Name="RDPProps" Height="340" AcceptsReturn="True" TextWrapping="Wrap"
                           VerticalScrollBarVisibility="Auto" FontFamily="Cascadia Code, Consolas" FontSize="11" Margin="0,2"/>
                </StackPanel>
              </Border>
              <Border Grid.Column="1" Style="{StaticResource Card}" Margin="8,0,0,0">
                <ScrollViewer>
                  <StackPanel>
                    <TextBlock Text="Security Toggles" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,12"/>
                    <TextBlock Text="REDIRECTION" FontSize="10" Foreground="#64748B" FontWeight="SemiBold" Margin="0,0,0,6"/>
                    <CheckBox x:Name="RDPClip"   Content="Clipboard"           IsChecked="True" Margin="0,3"/>
                    <CheckBox x:Name="RDPDrives" Content="Drive redirection"   Margin="0,3"/>
                    <CheckBox x:Name="RDPPrint"  Content="Printer redirection" IsChecked="True" Margin="0,3"/>
                    <CheckBox x:Name="RDPCam"    Content="Camera and mic"      IsChecked="True" Margin="0,3"/>
                    <CheckBox x:Name="RDPUSB"    Content="USB redirection"     Margin="0,3"/>
                    <CheckBox x:Name="RDPSC"     Content="Smart card"          Margin="0,3"/>
                    <Separator Margin="0,8"/>
                    <TextBlock Text="SECURITY" FontSize="10" Foreground="#64748B" FontWeight="SemiBold" Margin="0,0,0,6"/>
                    <CheckBox x:Name="RDPWatermark" Content="Session watermarking"        IsChecked="True" Margin="0,3"/>
                    <CheckBox x:Name="RDPScrCap"    Content="Screen capture protection"  IsChecked="True" Margin="0,3"/>
                    <CheckBox x:Name="RDPWebAuthn"  Content="WebAuthn (FIDO2 passthrough)" IsChecked="True" Margin="0,3"/>
                    <Separator Margin="0,8"/>
                    <TextBlock Text="DISPLAY" FontSize="10" Foreground="#64748B" FontWeight="SemiBold" Margin="0,0,0,6"/>
                    <CheckBox x:Name="RDPMultiMon"  Content="Multiple monitors"    IsChecked="True" Margin="0,3"/>
                    <CheckBox x:Name="RDPDynRes"    Content="Dynamic resolution"   IsChecked="True" Margin="0,3"/>
                    <CheckBox x:Name="RDPSmartSz"   Content="Smart sizing"         Margin="0,3"/>
                    <Separator Margin="0,8"/>
                    <Button x:Name="BtnBuildRDP" Content="Build Property String" Style="{StaticResource BtnPrimary}" Padding="0,9"/>
                  </StackPanel>
                </ScrollViewer>
              </Border>
            </Grid>
          </StackPanel>
        </ScrollViewer>


        <!-- ===== BEST PRACTICES ===== -->
        <ScrollViewer x:Name="PanelBP" Visibility="Collapsed" Padding="20">
          <StackPanel>
            <Grid Margin="0,0,0,14">
              <StackPanel>
                <TextBlock Text="AVD Best Practices Assessment" FontSize="22" FontWeight="SemiBold"/>
                <TextBlock Text="Checks your environment against Microsoft AVD best practices and security baseline" FontSize="12" Foreground="#64748B"/>
              </StackPanel>
              <StackPanel HorizontalAlignment="Right" VerticalAlignment="Bottom" Orientation="Horizontal">
                <Button x:Name="BtnRunBP"    Content="Run Assessment"  Style="{StaticResource BtnPrimary}" Padding="16,8" Margin="4,0"/>
                <Button x:Name="BtnExportBP" Content="Export CSV"      Style="{StaticResource BtnSec}"     Padding="12,8" Margin="4,0"/>
              </StackPanel>
            </Grid>
            <!-- Score cards -->
            <Grid Margin="0,0,0,12">
              <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition/><ColumnDefinition/><ColumnDefinition/></Grid.ColumnDefinitions>
              <Border Grid.Column="0" Style="{StaticResource Card}" Margin="0,0,6,0">
                <StackPanel>
                  <TextBlock Text="OVERALL SCORE" FontSize="9" Foreground="#64748B" FontWeight="SemiBold"/>
                  <TextBlock x:Name="BPScore" Text="-" FontSize="34" FontWeight="Bold" Foreground="#50ABF1" Margin="0,4,0,2"/>
                  <TextBlock x:Name="BPScoreLbl" Text="/ 100" FontSize="11" Foreground="#475569"/>
                </StackPanel>
              </Border>
              <Border Grid.Column="1" Style="{StaticResource Card}" Margin="3,0">
                <StackPanel>
                  <TextBlock Text="PASSED" FontSize="9" Foreground="#64748B" FontWeight="SemiBold"/>
                  <TextBlock x:Name="BPPass" Text="-" FontSize="34" FontWeight="Bold" Foreground="#10B981" Margin="0,4,0,2"/>
                  <TextBlock Text="checks" FontSize="11" Foreground="#475569"/>
                </StackPanel>
              </Border>
              <Border Grid.Column="2" Style="{StaticResource Card}" Margin="3,0">
                <StackPanel>
                  <TextBlock Text="WARNINGS" FontSize="9" Foreground="#64748B" FontWeight="SemiBold"/>
                  <TextBlock x:Name="BPWarn" Text="-" FontSize="34" FontWeight="Bold" Foreground="#F59E0B" Margin="0,4,0,2"/>
                  <TextBlock Text="checks" FontSize="11" Foreground="#475569"/>
                </StackPanel>
              </Border>
              <Border Grid.Column="3" Style="{StaticResource Card}" Margin="6,0,0,0">
                <StackPanel>
                  <TextBlock Text="FAILED" FontSize="9" Foreground="#64748B" FontWeight="SemiBold"/>
                  <TextBlock x:Name="BPFail" Text="-" FontSize="34" FontWeight="Bold" Foreground="#EF4444" Margin="0,4,0,2"/>
                  <TextBlock Text="checks" FontSize="11" Foreground="#475569"/>
                </StackPanel>
              </Border>
            </Grid>
            <Border Style="{StaticResource Card}">
              <DockPanel>
                <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="0,0,0,10">
                  <TextBlock Text="Filter category: " FontSize="12" Foreground="#64748B" VerticalAlignment="Center" Margin="0,0,8,0"/>
                  <ComboBox x:Name="BPCatFilter" Width="200" FontSize="11">
                    <ComboBoxItem Content="All Categories"         IsSelected="True"/>
                    <ComboBoxItem Content="Host Pool Configuration"/>
                    <ComboBoxItem Content="Session Hosts"/>
                    <ComboBoxItem Content="Identity and Security"/>
                    <ComboBoxItem Content="Networking"/>
                    <ComboBoxItem Content="Monitoring"/>
                    <ComboBoxItem Content="FSLogix Profiles"/>
                    <ComboBoxItem Content="Cost Optimization"/>
                    <ComboBoxItem Content="Prerequisites"/>
                  </ComboBox>
                </StackPanel>
                <DataGrid x:Name="GridBP" CanUserSortColumns="True" Height="540">
                  <DataGrid.Columns>
                    <DataGridTemplateColumn Header="Status" Width="80">
                      <DataGridTemplateColumn.CellTemplate><DataTemplate>
                        <Border CornerRadius="4" Padding="6,2" Background="{Binding StatusBg}" HorizontalAlignment="Left">
                          <TextBlock Text="{Binding Status}" FontSize="10" FontWeight="SemiBold" Foreground="{Binding StatusFg}"/>
                        </Border>
                      </DataTemplate></DataGridTemplateColumn.CellTemplate>
                    </DataGridTemplateColumn>
                    <DataGridTextColumn Header="Category"    Binding="{Binding Category}"   Width="170"/>
                    <DataGridTextColumn Header="Check"       Binding="{Binding Check}"      Width="280"/>
                    <DataGridTextColumn Header="Finding"     Binding="{Binding Finding}"    Width="*"/>
                    <DataGridTextColumn Header="Remediation" Binding="{Binding Remediation}" Width="300"/>
                    <DataGridTextColumn Header="Priority"    Binding="{Binding Priority}"   Width="80"/>
                  </DataGrid.Columns>
                </DataGrid>
              </DockPanel>
            </Border>
          </StackPanel>
        </ScrollViewer>

        <!-- ===== LOG ===== -->
        <Border x:Name="PanelLog" Visibility="Collapsed" Style="{StaticResource Card}" Margin="16">
          <DockPanel>
            <Grid DockPanel.Dock="Top" Margin="0,0,0,10">
              <StackPanel>
                <TextBlock Text="Activity Log" FontSize="22" FontWeight="SemiBold"/>
                <TextBlock Text="Real-time deployment and management operations" FontSize="12" Foreground="#64748B"/>
              </StackPanel>
              <Button x:Name="BtnClearLog" Content="Clear" Style="{StaticResource BtnSec}" HorizontalAlignment="Right" VerticalAlignment="Top" Padding="10,6"/>
            </Grid>
            <Border Background="#060E1A" CornerRadius="8" BorderBrush="#1E3A5F" BorderThickness="1" Padding="6">
              <RichTextBox x:Name="MainLogBox" Background="Transparent" Foreground="#94A3B8"
                           BorderThickness="0" IsReadOnly="True" FontFamily="Cascadia Code, Consolas" FontSize="12"
                           VerticalScrollBarVisibility="Auto"/>
            </Border>
          </DockPanel>
        </Border>

        <!-- ===== SETTINGS ===== -->
        <ScrollViewer x:Name="PanelSet" Visibility="Collapsed" Padding="20">
          <StackPanel>
            <TextBlock Text="Settings" FontSize="22" FontWeight="SemiBold" Margin="0,0,0,14"/>
            <Grid>
              <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
              <Border Grid.Column="0" Style="{StaticResource Card}" Margin="0,0,8,0">
                <StackPanel>
                  <TextBlock Text="Azure Connection Defaults" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,14"/>
                  <Label Content="Default Subscription ID"/>
                  <TextBox x:Name="SetSubId" Margin="0,2,0,10"/>
                  <Label Content="Default Tenant ID"/>
                  <TextBox x:Name="SetTenantId" Margin="0,2,0,10"/>
                  <Label Content="Default Location"/>
                  <ComboBox x:Name="SetLocation" Margin="0,2,0,14"/>
                  <Separator/>
                  <TextBlock Text="Auto-Refresh Interval" FontSize="11" Foreground="#64748B" Margin="0,8,0,4"/>
                  <StackPanel Orientation="Horizontal">
                    <Slider x:Name="SetRefresh" Minimum="30" Maximum="300" Value="60" Width="180" VerticalAlignment="Center"/>
                    <TextBlock x:Name="TxtSetRefresh" Text="60 seconds" FontSize="12" Foreground="#50ABF1" Margin="8,0,0,0" VerticalAlignment="Center"/>
                  </StackPanel>
                </StackPanel>
              </Border>
              <Border Grid.Column="1" Style="{StaticResource Card}" Margin="8,0,0,0">
                <StackPanel>
                  <TextBlock Text="Preferences" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,14"/>
                  <CheckBox x:Name="SetAutoLic"     Content="Auto-scan licenses on connect"  IsChecked="True" Margin="0,4"/>
                  <CheckBox x:Name="SetAutoRefresh" Content="Enable dashboard auto-refresh"  IsChecked="True" Margin="0,4"/>
                  <CheckBox x:Name="SetShowCost"    Content="Show cost estimates in wizard"  IsChecked="True" Margin="0,4"/>
                  <Separator Margin="0,12"/>
                  <StackPanel Orientation="Horizontal">
                    <Button x:Name="BtnSaveSet" Content="Save Settings"   Style="{StaticResource BtnPrimary}" Padding="14,8" Margin="0,0,8,0"/>
                    <Button x:Name="BtnTestConn" Content="Test Connection" Style="{StaticResource BtnSec}" Padding="14,8"/>
                  </StackPanel>
                </StackPanel>
              </Border>
            </Grid>
          </StackPanel>
        </ScrollViewer>

      </Grid><!-- /content -->
    </Grid><!-- /body -->

    <!-- STATUS BAR -->
    <Border Grid.Row="2" Background="#060C18" BorderBrush="#1B324E" BorderThickness="0,1,0,0">
      <Grid Margin="16,0">
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock x:Name="TxtStatus" Text="Ready" FontSize="11" Foreground="#64748B" Margin="0,0,14,0"/>
          <ProgressBar x:Name="ProgStatus" Width="100" Height="4" Maximum="100" Value="0" Margin="0,0,8,0"/>
          <TextBlock x:Name="TxtStatusDetail" Text="" FontSize="11" Foreground="#475569"/>
        </StackPanel>
        <StackPanel HorizontalAlignment="Right" Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock x:Name="TxtCountdown" Text="" FontSize="11" Foreground="#334155" Margin="0,0,12,0"/>
          <Ellipse x:Name="StatusDot" Width="7" Height="7" Fill="#334155" Margin="0,0,6,0"/>
          <TextBlock x:Name="TxtConnStatus" Text="Disconnected" FontSize="11" Foreground="#64748B"/>
        </StackPanel>
      </Grid>
    </Border>
  </Grid>
</Window>
'@

# ============================================================================
# LOAD XAML
# ============================================================================
Write-Host "Loading XAML..." -ForegroundColor DarkGray
try {
    $Reader = [System.Xml.XmlNodeReader]::new($XAML)
    $Window = [System.Windows.Markup.XamlReader]::Load($Reader)
    Write-Host "XAML OK" -ForegroundColor Green
} catch {
    [System.Console]::Error.WriteLine("XAML LOAD FAILED: $($_.Exception.Message)")
    if ($_.Exception.InnerException) { [System.Console]::Error.WriteLine($_.Exception.InnerException.Message) }
    Read-Host "Press Enter"
    exit 1
}

# ============================================================================
# ELEMENT REFERENCES
# ============================================================================
function X { param([string]$n) $Window.FindName($n) }

# Header
$BtnConnect     = X "BtnConnect";    $BtnRefreshAll = X "BtnRefreshAll"
$TxtSubName     = X "TxtSubName";    $CboSubSwitcher = X "CboSubSwitcher"
$TxtLicBadge    = X "TxtLicBadge";   $DotConn        = X "DotConn"

# Nav
$NavDash  = X "NavDash";  $NavLicense = X "NavLicense"
$NavWizard= X "NavWizard";$NavHP      = X "NavHP"
$NavSess  = X "NavSess";  $NavAG      = X "NavAG"
$NavFSL   = X "NavFSL";   $NavScale   = X "NavScale"
$NavMon   = X "NavMon";   $NavRDP     = X "NavRDP"
$NavBP    = X "NavBP"
$NavLog   = X "NavLog";   $NavSet     = X "NavSet"
$AllNavBtns = @($NavDash,$NavLicense,$NavWizard,$NavHP,$NavSess,$NavAG,$NavFSL,$NavScale,$NavMon,$NavRDP,$NavBP,$NavLog,$NavSet)

# Panels
$PanelDash  = X "PanelDash";  $PanelLic   = X "PanelLic"
$PanelWiz   = X "PanelWiz";   $PanelHP    = X "PanelHP"
$PanelSess  = X "PanelSess";  $PanelAG    = X "PanelAG"
$PanelFSL   = X "PanelFSL";   $PanelScale = X "PanelScale"
$PanelMon   = X "PanelMon";   $PanelRDP   = X "PanelRDP"
$PanelBP    = X "PanelBP";    $BPScore    = X "BPScore";   $BPPass  = X "BPPass"
$BPWarn     = X "BPWarn";     $BPFail     = X "BPFail";    $GridBP  = X "GridBP"
$BPCatFilter= X "BPCatFilter"; $BtnRunBP  = X "BtnRunBP";  $BtnExportBP = X "BtnExportBP"
$PanelLog   = X "PanelLog";   $PanelSet   = X "PanelSet"
$AllPanels  = @($PanelDash,$PanelLic,$PanelWiz,$PanelHP,$PanelSess,$PanelAG,$PanelFSL,$PanelScale,$PanelMon,$PanelRDP,$PanelBP,$PanelLog,$PanelSet)
$PanelMap   = @{ Dash=$PanelDash; Lic=$PanelLic; Wiz=$PanelWiz; HP=$PanelHP; Sess=$PanelSess; AG=$PanelAG; FSL=$PanelFSL; Scale=$PanelScale; Mon=$PanelMon; RDP=$PanelRDP; BP=$PanelBP; Log=$PanelLog; Set=$PanelSet }

# Dashboard
$MetHP=$null;$MetHosts=$null;$MetSess=$null;$MetLic=$null;$MetWS=$null;$GridDash=$null;$TxtDashSub=$null;$BtnDashRefresh=$null;$BtnNewDeploy=$null
try {
$MetHP        = X "MetHP";         $MetHosts      = X "MetHosts"
$MetSess      = X "MetSess";       $MetLic        = X "MetLic"
$MetWS        = X "MetWS";         $GridDash      = X "GridDash"
$TxtDashSub   = X "TxtDashSub";    $BtnDashRefresh= X "BtnDashRefresh"
$BtnNewDeploy = X "BtnNewDeploy"
} catch {}

# License
$LicTotalUsers = X "LicTotalUsers"; $LicTotalSub = X "LicTotalSub"
$LicType       = X "LicType";       $LicSupport  = X "LicSupport"
$GridLic       = X "GridLic";       $GridLicReqs = X "GridLicReqs"
$TxtLicRec     = X "TxtLicRec";     $BtnScanLic  = X "BtnScanLic"

# Wizard
$TxtWizType   = X "TxtWizType";   $WizStepList = X "WizStepList"
$WizProg      = X "WizProg";      $TxtWizStep  = X "TxtWizStep"
$RdoPooled    = X "RdoPooled";    $RdoPersonal = X "RdoPersonal"
$RdoRemoteApp = X "RdoRemoteApp"; $BtnWizNext  = X "BtnWizNext"
$BtnWizBack   = X "BtnWizBack";   $BtnWizCancel= X "BtnWizCancel"
$BtnDeploy    = X "BtnDeploy";    $BtnCancelDeploy = X "BtnCancelDeploy"
$WizHPName    = X "WizHPName";    $WizHPFriendly = X "WizHPFriendly"
$WizSub       = X "WizSub";       $WizRG       = X "WizRG"
$WizRegion    = X "WizRegion";    $WizWS       = X "WizWS"
$WizTagK      = X "WizTagK";      $WizTagV     = X "WizTagV"
$BtnAddTag    = X "BtnAddTag";    $GridTags    = X "GridTags"
$SliderMaxSess= X "SliderMaxSess";$TxtMaxSess  = X "TxtMaxSess"
$RdoBreadth   = X "RdoBreadth";   $RdoDepth    = X "RdoDepth"
$RdoAutoAssign= X "RdoAutoAssign";$RdoDirectAssign = X "RdoDirectAssign"
$ChkStartVM   = X "ChkStartVM";   $ChkValidation = X "ChkValidation"
$SliderTokenHrs = X "SliderTokenHrs"; $TxtTokenHrs = X "TxtTokenHrs"
$WizVNetRG    = X "WizVNetRG";    $WizVNet     = X "WizVNet"
$WizSubnet    = X "WizSubnet";    $BtnLoadVNet = X "BtnLoadVNet"
$WizKV        = X "WizKV";        $WizAdminUser= X "WizAdminUser"
$WizAdminPass = X "WizAdminPass"
$SliderVMCnt  = X "SliderVMCnt";  $TxtVMCnt    = X "TxtVMCnt"
$WizVMPrefix  = X "WizVMPrefix";  $WizVMSize   = X "WizVMSize"
$TxtVMDetail  = X "TxtVMDetail"
$WizDiskType  = X "WizDiskType";  $WizAvail    = X "WizAvail"
$ChkHybridBenefit = X "ChkHybridBenefit"
$RdoMarket    = X "RdoMarket";    $RdoGallery  = X "RdoGallery"
$WizMktImg    = X "WizMktImg";    $WizGallery  = X "WizGallery"
$WizGalleryImg = X "WizGalleryImg"
$RdoEntra     = X "RdoEntra";     $RdoHybrid   = X "RdoHybrid"
$EntraPanel   = X "EntraPanel";   $HybridPanel = X "HybridPanel"
$ChkIntune    = X "ChkIntune"
$WizDomain    = X "WizDomain";    $WizOU       = X "WizOU"
$WizDomainUser= X "WizDomainUser";$WizDomainPass = X "WizDomainPass"
$WizDCIPs     = X "WizDCIPs"
$ChkFSL       = X "ChkFSL";       $WizFSLBackend = X "WizFSLBackend"
$WizSA        = X "WizSA";        $WizShareName  = X "WizShareName"
$SliderProfGB = X "SliderProfGB"; $TxtProfGB     = X "TxtProfGB"
$WizFSLType   = X "WizFSLType";   $ChkCloudCache = X "ChkCloudCache"
$ChkKFM       = X "ChkKFM"
$WizRDPPreset = X "WizRDPPreset"; $WizChkClip = X "WizChkClip"
$WizChkDrives = X "WizChkDrives"; $WizChkCam  = X "WizChkCam"
$WizChkUSB    = X "WizChkUSB";    $WizChkWM   = X "WizChkWM"
$WizChkSCP    = X "WizChkSCP"
$ChkScaling   = X "ChkScaling";   $WizRampUp  = X "WizRampUp"
$WizPeakStart = X "WizPeakStart"; $WizRampDown = X "WizRampDown"
$WizOffPeak   = X "WizOffPeak"
$SliderCap    = X "SliderCap";    $TxtCap     = X "TxtCap"
$SliderMinH   = X "SliderMinH";   $TxtMinH    = X "TxtMinH"
$ChkMonitor   = X "ChkMonitor";   $WizLAW     = X "WizLAW"
$GridReview   = X "GridReview";   $WizLog     = X "WizLog"
$TxtEstCost   = X "TxtEstCost";   $TxtEstNote = X "TxtEstNote"
$DeployProg   = X "DeployProg";   $TxtDeployPhase = X "TxtDeployPhase"

$WS = @()
for ($i=1;$i -le 8;$i++) { $WS += X "WS$i" }

# HP / Sess / AG / FSL
$GridHP      = X "GridHP";       $BtnRefHP      = X "BtnRefHP";    $BtnNewHP      = X "BtnNewHP"
$FilterHP    = X "FilterHP";     $BtnExportHP   = X "BtnExportHP"
$GridSess    = X "GridSess";     $SessFilter    = X "SessFilter";   $BtnRefSess    = X "BtnRefSess"
$FilterSess  = X "FilterSess";   $BtnExportSess = X "BtnExportSess"
$BtnHealAll  = X "BtnHealAll";   $BtnDrainAll   = X "BtnDrainAll"
$GridUserSess     = X "GridUserSess"
$BtnRefUserSess   = X "BtnRefUserSess";   $BtnMsgUser        = X "BtnMsgUser"
$BtnDisconnectUser= X "BtnDisconnectUser"; $BtnLogoffUser     = X "BtnLogoffUser"
$BtnExportUserSess= X "BtnExportUserSess"
$GridAG    = X "GridAG";    $GridApps   = X "GridApps";   $BtnRefAG = X "BtnRefAG"; $BtnNewAG = X "BtnNewAG"
$BtnAssignUsers = X "BtnAssignUsers"
$GridFSL   = X "GridFSL";   $FSLTotal   = X "FSLTotal";   $FSLSize    = X "FSLSize"
$FSLQuota  = X "FSLQuota";  $SliderFSLAlert = X "SliderFSLAlert"; $TxtFSLAlert = X "TxtFSLAlert"
$BtnFSLLocks = X "BtnFSLLocks"; $BtnFSLTmp = X "BtnFSLTmp"; $BtnFSLDiag = X "BtnFSLDiag"; $BtnFSLRef = X "BtnFSLRef"
$BtnRefScale = X "BtnRefScale"; $BtnNewScale = X "BtnNewScale"; $BtnAddAlert = X "BtnAddAlert"
$BtnNewLAW = X "BtnNewLAW"; $BtnNewSA = X "BtnNewSA"; $BtnNewKV = X "BtnNewKV"
$BtnNewRG = X "BtnNewRG"; $BtnNewWS = X "BtnNewWS"; $BtnPubApp = X "BtnPubApp"
$GridScale = X "GridScale"; $GridLAW = X "GridLAW"; $GridAlerts = X "GridAlerts"
$CboSubSwitcher = X "CboSubSwitcher"

# RDP
$RDPPool    = X "RDPPool";   $RDPPreset  = X "RDPPreset"; $RDPProps = X "RDPProps"
$BtnBuildRDP= X "BtnBuildRDP"; $BtnApplyRDP = X "BtnApplyRDP"
$RDPClip    = X "RDPClip";   $RDPDrives  = X "RDPDrives"; $RDPPrint = X "RDPPrint"
$RDPCam     = X "RDPCam";    $RDPUSB     = X "RDPUSB";    $RDPSC    = X "RDPSC"
$RDPWatermark = X "RDPWatermark"; $RDPScrCap = X "RDPScrCap"; $RDPWebAuthn = X "RDPWebAuthn"
$RDPMultiMon= X "RDPMultiMon"; $RDPDynRes = X "RDPDynRes"; $RDPSmartSz = X "RDPSmartSz"

# Status / Log / Settings
$TxtStatus     = X "TxtStatus";     $ProgStatus     = X "ProgStatus"
$TxtStatusDetail = X "TxtStatusDetail"; $TxtCountdown = X "TxtCountdown"
$StatusDot     = X "StatusDot";     $TxtConnStatus  = X "TxtConnStatus"
$MainLogBox    = X "MainLogBox";    $Global:LogBox  = $MainLogBox
$BtnClearLog   = X "BtnClearLog"
$LoadingOverlay = X "LoadingOverlay"; $LoadingText = X "LoadingText"
$SetSubId      = X "SetSubId";      $SetTenantId    = X "SetTenantId"
$SetLocation   = X "SetLocation";   $SetRefresh     = X "SetRefresh"
$TxtSetRefresh = X "TxtSetRefresh"; $BtnSaveSet     = X "BtnSaveSet"
$BtnTestConn   = X "BtnTestConn";   $SetAutoLic     = X "SetAutoLic"

# Context menus
$CtxHPPortal       = X "CtxHPPortal";       $CtxHPDelete   = X "CtxHPDelete"
$CtxSessEnableDrain = X "CtxSessEnableDrain"; $CtxSessDisableDrain = X "CtxSessDisableDrain"
$CtxSessPortal     = X "CtxSessPortal";      $CtxSessRemove = X "CtxSessRemove"

# Style resources
$Script:NavActiveStyle   = $Window.FindResource("NavActive")
$Script:NavInactiveStyle = $Window.FindResource("NavBtn")

# ============================================================================
# NAVIGATION
# ============================================================================
function Show-Loading { param([string]$Text="Loading...")
    try { $script:LoadingText.Text = $Text; $script:LoadingOverlay.Visibility=[System.Windows.Visibility]::Visible } catch {}
}
function Hide-Loading {
    try { $script:LoadingOverlay.Visibility=[System.Windows.Visibility]::Collapsed } catch {}
}

function Switch-Panel {
    param([string]$Tag)
    $newPanel = if ($PanelMap.ContainsKey($Tag)) { $PanelMap[$Tag] } else { return }

    # Update nav styles immediately
    foreach ($b in $AllNavBtns) {
        try { $b.Style = if ($b.Tag -eq $Tag) { $Script:NavActiveStyle } else { $Script:NavInactiveStyle } } catch {}
    }

    $currentPanel = $AllPanels | Where-Object { $_.Visibility -eq [System.Windows.Visibility]::Visible } | Select-Object -First 1

    $dur150 = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(150))
    $dur220 = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(220))

    if ($currentPanel -and $currentPanel -ne $newPanel) {
        # Capture for closure
        $cp = $currentPanel; $np = $newPanel
        $fadeOut          = [System.Windows.Media.Animation.DoubleAnimation]::new()
        $fadeOut.From     = 1.0; $fadeOut.To = 0.0; $fadeOut.Duration = $dur150
        $fadeOut.EasingFunction = [System.Windows.Media.Animation.CubicEase]::new()
        $fadeOut.EasingFunction.EasingMode = [System.Windows.Media.Animation.EasingMode]::EaseIn

        $fadeOut.Add_Completed(({
            param($s,$e)
            $cp.Visibility = [System.Windows.Visibility]::Collapsed
            $cp.Opacity    = 1.0
            $np.Opacity    = 0.0
            $np.Visibility = [System.Windows.Visibility]::Visible
            $fadeIn         = [System.Windows.Media.Animation.DoubleAnimation]::new()
            $fadeIn.From    = 0.0; $fadeIn.To = 1.0; $fadeIn.Duration = $dur220
            $easeOut        = [System.Windows.Media.Animation.CubicEase]::new()
            $easeOut.EasingMode = [System.Windows.Media.Animation.EasingMode]::EaseOut
            $fadeIn.EasingFunction = $easeOut
            $np.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $fadeIn)
        }.GetNewClosure()))

        $currentPanel.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $fadeOut)
    } elseif ($newPanel) {
        foreach ($p in $AllPanels) { $p.Visibility = [System.Windows.Visibility]::Collapsed }
        $np = $newPanel; $np.Opacity = 0.0; $np.Visibility = [System.Windows.Visibility]::Visible
        $fadeIn = [System.Windows.Media.Animation.DoubleAnimation]::new()
        $fadeIn.From = 0.0; $fadeIn.To = 1.0; $fadeIn.Duration = $dur220
        $np.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $fadeIn)
    }

    # Load data for the panel
    switch ($Tag) {
        "HP"    { if ($Global:IsConnected) { Show-Loading "Loading host pools...";   Load-HostPools;       Hide-Loading } }
        "Sess"  { if ($Global:IsConnected) { Show-Loading "Loading session hosts...";Load-SessionHosts; Load-UserSessions; Hide-Loading } }
        "AG"    { if ($Global:IsConnected) { Show-Loading "Loading app groups...";   Load-AppGroups;       Hide-Loading } }
        "Scale" { if ($Global:IsConnected) { Show-Loading "Loading scaling plans...";Load-ScalingPlans;    Hide-Loading } }
        "Mon"   { if ($Global:IsConnected) { Show-Loading "Loading monitoring...";   Load-Monitoring;      Hide-Loading } }
        "Lic"   { if ($Global:IsConnected -and -not $Script:LicScanned) { Show-Loading "Scanning licenses..."; Invoke-LicenseScan; Hide-Loading } }
        "BP"    { if ($Global:IsConnected) { Show-Loading "Running assessment...";   Invoke-BestPracticesCheck; Hide-Loading } }
        "Wiz"   { if ($Global:IsConnected) { Load-WizardDropdowns } }
    }
}
foreach ($b in $AllNavBtns) {
    $b.Add_Click({ param($s,$e) Switch-Panel $s.Tag })
}

# ============================================================================
# UI STATE HELPERS
# ============================================================================
function Set-Status { param([string]$Msg, [int]$Pct=0, [string]$Detail="")
    try { if ($script:TxtStatus)        { $script:TxtStatus.Text       = $Msg } }       catch {}
    try { if ($script:ProgStatus)       { $script:ProgStatus.Value     = [double]$Pct } } catch {}
    try { if ($script:TxtStatusDetail)  { $script:TxtStatusDetail.Text = $Detail } }    catch {}
    $Global:Sync.StatusMsg = $Msg
}

function Set-Connected { param([bool]$State)
    $Global:IsConnected = $State
    $c = if ($State) { "#10B981" } else { "#EF4444" }
    $t = if ($State) { "Connected" }   else { "Disconnected" }
    $b = if ($State) { "Disconnect" }  else { "Connect to Azure" }
    try { $script:StatusDot.Fill                = Get-Brush $c } catch {}
    try { $script:TxtConnStatus.Text            = $t;  $script:TxtConnStatus.Foreground = Get-Brush $c } catch {}
    try { $script:BtnConnect.Content            = $b } catch {}
    try { $script:DotConn.Fill                  = Get-Brush $c } catch {}
    if ($State -and $Global:Subscription) {
        try { $script:TxtSubName.Text           = $Global:Subscription.Name; $script:TxtSubName.Foreground = Get-Brush "#E2E8F0" } catch {}
    }
    if ($State -and $Global:TenantId) {
        # TenantId now shown in subscription switcher tooltip, not a separate label
        try { $script:CboSubSwitcher.ToolTip = "Tenant: $($Global:TenantId)" } catch {}
    }
}

# ============================================================================
# LICENSE SCAN
# ============================================================================
$Script:LicScanned = $false

function Invoke-LicenseScan {
    if (-not $Global:IsConnected) { return }
    Set-Status "Scanning licenses via Microsoft Graph..." 20
    Write-Log "Starting license assessment..." "STEP"
    try {
        $token = Get-GraphToken   # cached, auto-refreshes before expiry
        $hdr   = @{ Authorization="Bearer $token"; "Content-Type"="application/json" }
        $skus  = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/subscribedSkus" -Headers $hdr -EA Stop).value

        $rows = [System.Collections.Generic.List[PSObject]]::new()
        $avdTotal=0; $supportCnt=0; $bestType="None"
        $recs = [System.Collections.Generic.List[string]]::new()

        foreach ($sku in $skus) {
            $sid = $sku.skuPartNumber
            $map = $Script:AvdSkuMap[$sid]
            $nm  = if ($map) { $map.Name } else { $sid }
            $avd = ($map -and $map.AVD)
            $tp  = if ($map) { $map.LicType } else { "Unknown" }
            $assigned  = $sku.consumedUnits
            $available = [Math]::Max(0, $sku.prepaidUnits.enabled - $assigned)
            if ($avd) {
                $avdTotal += $assigned
                if ($tp -eq "Desktop+RemoteApp") { $bestType = "Desktop+RemoteApp" }
                elseif ($tp -eq "RemoteApp only" -and $bestType -eq "None") { $bestType = "RemoteApp only" }
            } elseif ($map) { $supportCnt++ }
            $rows.Add([PSCustomObject]@{
                AvdIcon  = if ($avd) {"[AVD]"} else {"---"}
                LicName  = $nm; SkuId=$sid; Assigned=$assigned; Available=$available; LicType=$tp
            })
        }

        $hasAvd    = $avdTotal -gt 0
        $hasP1     = $skus | Where-Object { $_.skuPartNumber -in @("AAD_PREMIUM","EMS","EMSPREMIUM","SPE_E3","SPE_E5","SPE_F3","O365_BUSINESS_PREMIUM") } | Select-Object -First 1
        $hasIntune = $skus | Where-Object { $_.skuPartNumber -in @("INTUNE_A","INTUNE_A_D","EMS","EMSPREMIUM","SPE_E3","SPE_E5","SPE_F3","O365_BUSINESS_PREMIUM") } | Select-Object -First 1

        $reqs = @(
            [PSCustomObject]@{Icon=if($hasAvd){"[OK]"}else{"[X]"}; Req="AVD entitlement license ($avdTotal users)"}
            [PSCustomObject]@{Icon=if($hasP1){"[OK]"}else{"[!]"};  Req="Entra ID P1 (Conditional Access)"}
            [PSCustomObject]@{Icon=if($hasIntune){"[OK]"}else{"[!]"}; Req="Intune (Entra ID Join management)"}
            [PSCustomObject]@{Icon="[i]"; Req="Azure subscription (billed separately)"}
        )

        if (-not $hasAvd) { $recs.Add("No AVD license detected. Assign Microsoft 365 E3/E5, F3, Windows 10/11 E3/E5, or standalone AVD.") }
        if (-not $hasP1)  { $recs.Add("Entra ID P1 recommended for Conditional Access and MFA.") }
        if (-not $hasIntune) { $recs.Add("Microsoft Intune required for Entra ID-joined session host management.") }
        if ($recs.Count -eq 0) { $recs.Add("All required licenses detected. Environment is ready for AVD deployment.") }

        $licTxt = "$avdTotal"
        $licBadge = if ($avdTotal -gt 0) { "Licenses: $avdTotal AVD users" } else { "Licenses: None detected" }
        $licColor = if ($avdTotal -gt 0) { "#10B981" } else { "#EF4444" }

        try { $script:LicTotalUsers.Text  = $licTxt } catch {}
        try { $script:LicType.Text        = if ($bestType -eq "None") {"None detected"} else {$bestType} } catch {}
        try { $script:LicSupport.Text     = "$supportCnt" } catch {}
        try { $script:GridLic.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[PSObject]]($rows.ToArray()) } catch {}
        try { $script:GridLicReqs.ItemsSource = $reqs } catch {}
        try { $script:TxtLicRec.Text      = $recs -join "`n`n" } catch {}
        try { $script:TxtLicBadge.Text    = $licBadge; $script:TxtLicBadge.Foreground = Get-Brush $licColor } catch {}
        try { $script:MetLic.Text         = $licTxt } catch {}

        $Script:LicScanned = $true
        Write-Log "License scan complete: $avdTotal AVD-entitled users | $supportCnt supporting licenses" "OK"
    } catch {
        Write-Log "License scan error: $_" "WARN"
        try { $script:TxtLicRec.Text = "Scan failed: $_`nEnsure Microsoft Graph read access is available." } catch {}
    }
    Set-Status "Ready" 0
}

# ============================================================================
# DATA LOAD FUNCTIONS
# ============================================================================
function Load-Dashboard {
    if (-not $Global:IsConnected) { return }
    Set-Status "Loading dashboard..." 30
    try {
        $hpAll  = @(Get-AzWvdHostPool -EA SilentlyContinue | Select-Object -First 100)
        $wsAll  = @(Get-AzWvdWorkspace -EA SilentlyContinue | Select-Object -First 50)
        $totalSess=0; $totalAvail=0; $totalHosts=0
        $rows = [System.Collections.Generic.List[PSObject]]::new()
        foreach ($hp in $hpAll) {
            $rg    = ($hp.Id -split "/")[4]
            $hosts = @(Get-AzWvdSessionHost -HostPoolName $hp.Name -ResourceGroupName $rg -EA SilentlyContinue)
            $avail = @($hosts | Where-Object {$_.Status -eq "Available"}).Count
            $sess  = ($hosts | Measure-Object -Property Session -Sum).Sum; if (-not $sess) {$sess=0}
            $totalAvail  += $avail; $totalHosts += $hosts.Count; $totalSess += [int]$sess
            $typeStr = switch ($hp.HostPoolType) {
                "Pooled"   { if ($hp.PreferredAppGroupType -eq "RailApplications") {"RemoteApp"} else {"Pooled"} }
                "Personal" { "Personal" } default { $hp.HostPoolType }
            }
            $typeBg = switch ($typeStr) { "Pooled"{"#0D2547"} "Personal"{"#065F46"} "RemoteApp"{"#2A1A00"} default{"#1E3A5F"} }
            $typeFg = switch ($typeStr) { "Pooled"{"#50ABF1"} "Personal"{"#10B981"} "RemoteApp"{"#F59E0B"} default{"#94A3B8"} }
            $rows.Add([PSCustomObject]@{
                Name=$hp.Name; Type=$typeStr; TypeBg=$typeBg; TypeFg=$typeFg
                LB=$hp.LoadBalancerType; MaxSess=$hp.MaxSessionLimit
                Hosts="$avail / $($hosts.Count)"; Sessions=[int]$sess
                Status=if($hp.ValidationEnvironment){"Validation"}else{"Production"}
                StatusBg=if($hp.ValidationEnvironment){"#2A1A00"}else{"#0A2010"}
                StatusFg=if($hp.ValidationEnvironment){"#F59E0B"}else{"#10B981"}
                Location=$hp.Location; RG=$rg
            })
        }
        try {
            $script:MetHP.Text      = "$($hpAll.Count)"
            $script:MetHosts.Text   = "$totalAvail"
            $script:MetHostsSub.Text = "$totalAvail avail / $totalHosts total"
            $script:MetSess.Text    = "$totalSess"
            $script:MetWS.Text      = "$($wsAll.Count)"
            $script:GridDash.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[PSObject]]($rows.ToArray())
            $script:TxtDashSub.Text = "Subscription: $($Global:Subscription.Name)  |  $($hpAll.Count) host pool(s)  |  $totalHosts session host(s)"
        } catch {}
        Write-Log "Dashboard: $($hpAll.Count) pools, $totalHosts hosts, $totalSess sessions" "OK"
    } catch { Write-Log "Dashboard error: $_" "WARN" }
    Set-Status "Ready" 0
}

function Load-HostPools {
    if (-not $Global:IsConnected) { return }
    Set-Status "Loading host pools..." 30
    try {
        $hpAll = @(Get-AzWvdHostPool -EA SilentlyContinue)
        $rows  = [System.Collections.Generic.List[PSObject]]::new()
        $hpNames = @("(All Pools)")
        foreach ($hp in $hpAll) {
            $rg    = ($hp.Id -split "/")[4]
            $hosts = @(Get-AzWvdSessionHost -HostPoolName $hp.Name -ResourceGroupName $rg -EA SilentlyContinue)
            $avail = @($hosts | Where-Object {$_.Status -eq "Available"}).Count
            $sess  = ($hosts | Measure-Object -Property Session -Sum).Sum; if (-not $sess) {$sess=0}
            $typeStr = switch ($hp.HostPoolType) {
                "Pooled"   { if ($hp.PreferredAppGroupType -eq "RailApplications") {"RemoteApp"} else {"Pooled"} }
                "Personal" { "Personal" } default { $hp.HostPoolType }
            }
            $typeBg = switch ($typeStr) { "Pooled"{"#0D2547"} "Personal"{"#065F46"} "RemoteApp"{"#2A1A00"} default{"#1E3A5F"} }
            $typeFg = switch ($typeStr) { "Pooled"{"#50ABF1"} "Personal"{"#10B981"} "RemoteApp"{"#F59E0B"} default{"#94A3B8"} }
            $rows.Add([PSCustomObject]@{
                Name=$hp.Name; Type=$typeStr; TypeBg=$typeBg; TypeFg=$typeFg
                LB=$hp.LoadBalancerType; MaxSess=$hp.MaxSessionLimit
                Hosts="$avail / $($hosts.Count)"; Sessions=[int]$sess
                Assignment=if($hp.PersonalDesktopAssignmentType){$hp.PersonalDesktopAssignmentType}else{"N/A"}
                Status=if($hp.ValidationEnvironment){"Validation"}else{"Production"}
                StatusBg=if($hp.ValidationEnvironment){"#2A1A00"}else{"#0A2010"}
                StatusFg=if($hp.ValidationEnvironment){"#F59E0B"}else{"#10B981"}
                Location=$hp.Location; RG=$rg; HPObj=$hp
            })
            $hpNames += $hp.Name
        }
        try { $script:GridHP.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[PSObject]]($rows.ToArray()); $Global:Sync["CachedHPRows"] = $rows.ToArray() } catch {}
        try { $script:SessFilter.ItemsSource  = $hpNames; $script:SessFilter.SelectedIndex = 0 } catch {}
        try { $script:RDPPool.ItemsSource     = @($hpAll | ForEach-Object {$_.Name}) } catch {}
        Write-Log "Host pools loaded: $($hpAll.Count)" "OK"
    } catch { Write-Log "Load-HostPools error: $_" "WARN" }
    Set-Status "Ready" 0
}

function Load-SessionHosts {
    if (-not $Global:IsConnected) { return }
    Set-Status "Loading session hosts..." 30
    try {
        $filter = try {$script:SessFilter.SelectedItem} catch {"(All Pools)"}
        $hpAll  = @(Get-AzWvdHostPool -EA SilentlyContinue)
        $rows   = [System.Collections.Generic.List[PSObject]]::new()
        foreach ($hp in $hpAll) {
            if ($filter -and $filter -ne "(All Pools)" -and $hp.Name -ne $filter) { continue }
            $rg    = ($hp.Id -split "/")[4]
            $hosts = @(Get-AzWvdSessionHost -HostPoolName $hp.Name -ResourceGroupName $rg -EA SilentlyContinue)
            foreach ($h in $hosts) {
                $vmn    = ($h.Name -split "/")[-1]
                $status = if ($h.Status) {$h.Status} else {"Unknown"}
                $sess   = if ($h.Session -ne $null) {$h.Session} else {0}
                $drain  = ($h.AllowNewSession -eq $false)
                $statusBg = switch ($status) { "Available"{"#0A2010"} "Unavailable"{"#2A0A0A"} "NeedsAssistance"{"#2A1A00"} default{"#1E3A5F"} }
                $statusFg = switch ($status) { "Available"{"#10B981"} "Unavailable"{"#EF4444"} "NeedsAssistance"{"#F59E0B"} default{"#94A3B8"} }
                $rows.Add([PSCustomObject]@{
                    VMName=$vmn; Pool=$hp.Name; Status=$status; StatusBg=$statusBg; StatusFg=$statusFg
                    Sessions=$sess; VMSize=($h.VirtualMachineId -split "/")[-1]
                    DrainText=if($drain){"Drain ON"}else{"Normal"}; DrainBg=if($drain){"#2A1A00"}else{"#0A2010"}; DrainFg=if($drain){"#F59E0B"}else{"#10B981"}
                    OSVer=$h.OSVersion
                    Heartbeat=if($h.LastHeartBeat){$h.LastHeartBeat.ToString("MM/dd HH:mm")}else{"--"}
                    HostObj=$h; HPName=$hp.Name; RG=$rg
                })
            }
        }
        try { $script:GridSess.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[PSObject]]($rows.ToArray()); $Global:Sync["CachedSessRows"] = $rows.ToArray() } catch {}
    } catch { Write-Log "Load-SessionHosts error: $_" "WARN" }
    Set-Status "Ready" 0
}

function Load-AppGroups {
    if (-not $Global:IsConnected) { return }
    Set-Status "Loading app groups..." 20
    try {
        $ags  = @(Get-AzWvdApplicationGroup -EA SilentlyContinue)
        $rows = [System.Collections.Generic.List[PSObject]]::new()
        foreach ($ag in $ags) {
            $rg   = ($ag.Id -split "/")[4]
            $apps = @(Get-AzWvdApplication -ApplicationGroupName $ag.Name -ResourceGroupName $rg -EA SilentlyContinue)
            $usrs = @(Get-AzRoleAssignment -Scope $ag.Id -RoleDefinitionName "Desktop Virtualization User" -EA SilentlyContinue).Count
            $typeBg = if ($ag.ApplicationGroupType -eq "Desktop") {"#0D2547"} else {"#2A1A00"}
            $typeFg = if ($ag.ApplicationGroupType -eq "Desktop") {"#50ABF1"} else {"#F59E0B"}
            $rows.Add([PSCustomObject]@{ Name=$ag.Name; Type=$ag.ApplicationGroupType; TypeBg=$typeBg; TypeFg=$typeFg; HostPool=($ag.HostPoolArmPath -split "/")[-1]; AppCount=$apps.Count; UserCount=$usrs })
        }
        try { $script:GridAG.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[PSObject]]($rows.ToArray()) } catch {}
    } catch { Write-Log "Load-AppGroups error: $_" "WARN" }
    Set-Status "Ready" 0
}

function Load-ScalingPlans {
    if (-not $Global:IsConnected) { return }
    Set-Status "Loading scaling plans..." 20
    try {
        $sps  = @(Get-AzWvdScalingPlan -EA SilentlyContinue)
        $rows = [System.Collections.Generic.List[PSObject]]::new()
        foreach ($sp in $sps) {
            $hps = @($sp.HostPoolReference | ForEach-Object {($_.HostPoolArmPath -split "/")[-1]})
            $rows.Add([PSCustomObject]@{ Name=$sp.Name; HP=($hps -join ", "); Status="Active"; StatusBg="#0A2010"; StatusFg="#10B981"; TZ=$sp.TimeZone; Schedules=$sp.Schedule.Count })
        }
        try { $script:GridScale.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[PSObject]]($rows.ToArray()) } catch {}
    } catch { Write-Log "Load-ScalingPlans error: $_" "WARN" }
    Set-Status "Ready" 0
}

function Load-Monitoring {
    if (-not $Global:IsConnected) { return }
    Set-Status "Loading monitoring..." 20
    try {
        $laws = @(Get-AzOperationalInsightsWorkspace -EA SilentlyContinue)
        $rows = [System.Collections.Generic.List[PSObject]]::new()
        foreach ($law in $laws) {
            $rows.Add([PSCustomObject]@{ Name=$law.Name; RG=($law.ResourceId -split "/")[4]; SKU=$law.Sku; Retention=$law.RetentionInDays })
        }
        try { $script:GridLAW.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[PSObject]]($rows.ToArray()) } catch {}
    } catch { Write-Log "Load-Monitoring error: $_" "WARN" }
    Set-Status "Ready" 0
}

# ============================================================================
# WIZARD
# ============================================================================
$Script:WizStep       = 1
$Script:WizTotalSteps = 8
$Script:WizTitles     = @(
    @{Title="Basics";             Sub="Name, region, resource group"}
    @{Title="Host Pool Config";   Sub="Load balancing, session limits"}
    @{Title="Networking";         Sub="VNet, subnet, credentials"}
    @{Title="Session Hosts";      Sub="VM size, count, OS image"}
    @{Title="Identity";           Sub="Entra ID or Hybrid AD Join"}
    @{Title="FSLogix Profiles";   Sub="Profile storage configuration"}
    @{Title="RDP and Scaling";    Sub="Session settings, auto-scale"}
    @{Title="Review and Deploy";  Sub="Confirm and create resources"}
)

function Update-WizardUI {
    param([int]$Step)
    for ($i=0;$i -lt $WS.Count;$i++) {
        $WS[$i].Visibility = if ($i -eq ($Step-1)) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    }
    $pct = [Math]::Round(($Step-1)*100.0/($Script:WizTotalSteps-1), 0)
    try { $script:WizProg.Value   = $pct } catch {}
    try { $script:TxtWizStep.Text = "Step $Step of $Script:WizTotalSteps" } catch {}
    try { $script:BtnWizBack.IsEnabled = ($Step -gt 1) } catch {}
    try { $script:BtnWizNext.Content   = if ($Step -eq $Script:WizTotalSteps) {"Finish"} else {"Next"} } catch {}
    try { $script:BtnDeploy.Visibility    = if ($Step -eq $Script:WizTotalSteps) {[System.Windows.Visibility]::Visible} else {[System.Windows.Visibility]::Collapsed} } catch {}
    try { $script:BtnWizNext.Visibility   = if ($Step -eq $Script:WizTotalSteps) {[System.Windows.Visibility]::Collapsed} else {[System.Windows.Visibility]::Visible} } catch {}
    $stepRows = [System.Collections.Generic.List[PSObject]]::new()
    for ($i=0;$i -lt $Script:WizTitles.Count;$i++) {
        $n = $i+1
        $active = ($n -eq $Step); $done = ($n -lt $Step)
        $stepRows.Add([PSCustomObject]@{
            Num    = if ($done) {"OK"} else {"$n"}
            NumFg  = if ($active) {"White"} elseif ($done) {"#10B981"} else {"#64748B"}
            StepBg = if ($active) {"#0078D4"} elseif ($done) {"#065F46"} else {"#1E3A5F"}
            Title  = $Script:WizTitles[$i].Title; Sub = $Script:WizTitles[$i].Sub
            TitleFg= if ($active) {"White"} elseif ($done) {"#10B981"} else {"#64748B"}
            Weight = if ($active) {"SemiBold"} else {"Normal"}
        })
    }
    try { $script:WizStepList.ItemsSource = $stepRows } catch {}
    if ($Step -eq $Script:WizTotalSteps) { Build-ReviewSummary }
}

function Load-WizardDropdowns {
    if (-not $Global:IsConnected) { return }
    try { $subs = @(Get-AzSubscription -EA SilentlyContinue); $script:WizSub.ItemsSource = @($subs | ForEach-Object {"$($_.Name) ($($_.Id))"}); $script:WizSub.SelectedIndex=0 } catch {}
    try { $rgs = @(Get-AzResourceGroup -EA SilentlyContinue | Sort-Object ResourceGroupName); $script:WizRG.ItemsSource = @($rgs | ForEach-Object {$_.ResourceGroupName}); $script:WizVNetRG.ItemsSource = @($rgs | ForEach-Object {$_.ResourceGroupName}) } catch {}
    try { $ws = @(Get-AzWvdWorkspace -EA SilentlyContinue); $script:WizWS.ItemsSource = @($ws | ForEach-Object {$_.Name}) } catch {}
    try { $kvs = @(Get-AzKeyVault -EA SilentlyContinue); $script:WizKV.ItemsSource = @($kvs | ForEach-Object {$_.VaultName}) } catch {}
    try { $sas = @(Get-AzStorageAccount -EA SilentlyContinue); $script:WizSA.ItemsSource = @($sas | ForEach-Object {$_.StorageAccountName}) } catch {}
    try { $laws = @(Get-AzOperationalInsightsWorkspace -EA SilentlyContinue); $script:WizLAW.ItemsSource = @($laws | ForEach-Object {$_.Name}) } catch {}
    try { $script:WizRegion.ItemsSource = $Script:AzureRegions; $script:WizRegion.SelectedValue = "eastus" } catch {}
    try { $vmItems = $Script:VmSizes | ForEach-Object {"$($_.Size) ($($_.vCPU) vCPU, $($_.RAM)) - $($_.UseCase)"}; $script:WizVMSize.ItemsSource = $vmItems; $script:WizVMSize.SelectedIndex=1 } catch {}
    try { $script:WizMktImg.ItemsSource = $Script:MktImages; $script:WizMktImg.SelectedIndex=0 } catch {}
    try { $script:GridTags.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[PSObject]]::new() } catch {}
}

function Build-ReviewSummary {
    # Pre-compute all values before building table (try{} inside if() is invalid PS5)
    $hpType = "Pooled"
    try { if ($script:RdoPersonal.IsChecked)  { $hpType = "Personal"  } } catch {}
    try { if ($script:RdoRemoteApp.IsChecked) { $hpType = "RemoteApp" } } catch {}
    $lb = "BreadthFirst"
    try { if ($script:RdoDepth.IsChecked) { $lb = "DepthFirst" } } catch {}
    $vmSzRaw = ""; try { $vmSzRaw = $script:WizVMSize.SelectedItem } catch {}
    $vmSz   = if ($vmSzRaw) { ($vmSzRaw -split " ")[0] } else { "Standard_D4s_v5" }
    $vmCnt  = 2;  try { $vmCnt  = [int]$script:SliderVMCnt.Value   } catch {}
    $join   = "Entra ID Join"
    try { if ($script:RdoHybrid.IsChecked) { $join = "Hybrid AD Join" } } catch {}
    $fsl   = "Enabled";  try { if (-not $script:ChkFSL.IsChecked)     { $fsl   = "Disabled" } } catch {}
    $scale = "Enabled";  try { if (-not $script:ChkScaling.IsChecked) { $scale = "Disabled" } } catch {}
    $hpName = "--"; try { $hpName = $script:WizHPName.Text   } catch {}
    $rgVal  = "--"; try { $rgVal  = $script:WizRG.SelectedItem } catch {}
    $locVal = "--"; try { $locVal = $script:WizRegion.SelectedValue } catch {}
    $vnetV  = "--"; try { $vnetV  = $script:WizVNet.SelectedItem   } catch {}
    $subnetV = "--"; try { $subnetV = $script:WizSubnet.SelectedItem } catch {}
    $maxSV  = "--"; try { $maxSV  = [int]$script:SliderMaxSess.Value } catch {}

    $rows = [System.Collections.Generic.List[PSObject]]::new()
    $settings = @(
        @{Section="Host Pool";       Setting="Name";           Value=$hpName}
        @{Section="Host Pool";       Setting="Type";           Value=$hpType}
        @{Section="Host Pool";       Setting="Load Balancing"; Value=$lb}
        @{Section="Host Pool";       Setting="Max Sessions";   Value=$maxSV}
        @{Section="Infrastructure";  Setting="Resource Group"; Value=$rgVal}
        @{Section="Infrastructure";  Setting="Location";       Value=$locVal}
        @{Section="Networking";      Setting="VNet";           Value=$vnetV}
        @{Section="Networking";      Setting="Subnet";         Value=$subnetV}
        @{Section="Session Hosts";   Setting="VM Size";        Value=$vmSz}
        @{Section="Session Hosts";   Setting="VM Count";       Value=$vmCnt}
        @{Section="Identity";        Setting="Join Type";      Value=$join}
        @{Section="Profiles";        Setting="FSLogix";        Value=$fsl}
        @{Section="Scaling";         Setting="Auto-Scale";     Value=$scale}
    )
    foreach ($s in $settings) { $rows.Add([PSCustomObject]$s) }
    try { $script:GridReview.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[PSObject]]($rows.ToArray()) } catch {}

    $priceMap = @{"Standard_D2s_v5"=55;"Standard_D4s_v5"=110;"Standard_D8s_v5"=220;"Standard_D16s_v5"=440;"Standard_D4ds_v5"=120;"Standard_D8ds_v5"=240;"Standard_E4s_v5"=185;"Standard_E8s_v5"=370;"Standard_B4ms"=60;"Standard_B8ms"=120}
    $unitCost = $priceMap[$vmSz]; if (-not $unitCost) { $unitCost = 100 }
    $totalEst = $vmCnt * $unitCost
    try { $script:TxtEstCost.Text = "~`$$totalEst/mo"; $script:TxtEstNote.Text = "($vmCnt x $vmSz, compute only)" } catch {}
}

function Load-WizardWorkspaces {
    if (-not $Global:IsConnected) { return }
    try { $ws = @(Get-AzWvdWorkspace -EA SilentlyContinue); $script:WizWS.ItemsSource = @($ws | ForEach-Object {$_.Name}) } catch {}
}
function Load-WizardLAWs {
    if (-not $Global:IsConnected) { return }
    try { $laws = @(Get-AzOperationalInsightsWorkspace -EA SilentlyContinue); $script:WizLAW.ItemsSource = @($laws | ForEach-Object {$_.Name}) } catch {}
}

# ============================================================================
# PREREQUISITES CHECK
# ============================================================================
function Test-Prerequisites {
    $issues = [System.Collections.Generic.List[string]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()

    # PowerShell version
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        $issues.Add("PowerShell 5.1+ required (found $($PSVersionTable.PSVersion))")
    }

    # Required Az modules with minimum versions
    $required = @{
        "Az.Accounts"                = "2.12.0"
        "Az.DesktopVirtualization"   = "3.1.0"
        "Az.Compute"                 = "5.0.0"
        "Az.Network"                 = "5.0.0"
        "Az.KeyVault"                = "4.0.0"
        "Az.OperationalInsights"     = "3.0.0"
        "Az.Resources"               = "6.0.0"
        "Az.Storage"                 = "5.0.0"
    }
    foreach ($mod in $required.GetEnumerator()) {
        $installed = Get-Module -ListAvailable -Name $mod.Key | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $installed) {
            $issues.Add("Missing module: $($mod.Key) (install with: Install-Module $($mod.Key) -Scope CurrentUser)")
        } elseif ([Version]$installed.Version -lt [Version]$mod.Value) {
            $warnings.Add("$($mod.Key) v$($installed.Version) is below recommended v$($mod.Value). Run: Update-Module $($mod.Key)")
        }
    }

    # ExecutionPolicy
    $ep = Get-ExecutionPolicy -Scope CurrentUser
    if ($ep -eq "Restricted") {
        $issues.Add("ExecutionPolicy is Restricted. Run: Set-ExecutionPolicy RemoteSigned -Scope CurrentUser")
    }

    if ($issues.Count -gt 0) {
        $msg = "Prerequisites check failed:`n`n" + ($issues -join "`n") + "`n`n" +
               $(if ($warnings.Count -gt 0) {"Warnings:`n" + ($warnings -join "`n") + "`n`n"} else {""}) +
               "The tool may not work correctly until these are resolved."
        [System.Windows.MessageBox]::Show($msg, "AVD Manager - Prerequisites", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
        Write-Log "Prerequisites: $($issues.Count) issue(s), $($warnings.Count) warning(s)" "WARN"
        $issues | ForEach-Object { Write-Log "  [PREREQ] $_" "ERROR" }
    } elseif ($warnings.Count -gt 0) {
        Write-Log "Prerequisites OK with $($warnings.Count) module version warning(s)" "WARN"
        $warnings | ForEach-Object { Write-Log "  [WARN] $_" "WARN" }
    } else {
        Write-Log "Prerequisites check passed (all modules OK)" "OK"
    }
    return ($issues.Count -eq 0)
}

# ============================================================================
# AUDIT LOG  (separate file — tracks all writes/deletes)
# ============================================================================
$Global:AuditLog = ".\AVD-Manager-Audit-$(Get-Date -f 'yyyyMMdd').log"

function Write-Audit {
    param([string]$Action, [string]$Resource, [string]$Detail="", [string]$Result="OK")
    $ts  = Get-Date -f "yyyy-MM-dd HH:mm:ss"
    $who = try { (Get-AzContext).Account.Id } catch { "unknown" }
    $sub = try { (Get-AzContext).Subscription.Name } catch { "unknown" }
    $line = "[$ts] [$Result] User=$who Sub=$sub Action=$Action Resource=$Resource $Detail"
    Add-Content $Global:AuditLog $line -EA SilentlyContinue
    Write-Log "[AUDIT] $Action $Resource" "INFO"
}

# ============================================================================
# CREDENTIAL HYGIENE  (clear sensitive strings from memory)
# ============================================================================
function Clear-SensitiveData {
    param([hashtable]$Params)
    foreach ($key in @("AdminPass","DomainPass","RegToken")) {
        if ($Params.ContainsKey($key)) {
            $Params[$key] = [string]::new('*', [Math]::Min(8, $Params[$key].Length))
        }
    }
}

function Protect-LogMessage {
    # Strip anything that looks like a token/password from log strings
    param([string]$Message)
    $Message = $Message -replace '(?i)(token|password|secret|key)\s*[:=]\s*\S+', '$1: [REDACTED]'
    $Message = $Message -replace 'eyJ[A-Za-z0-9+/]{40,}', '[JWT_REDACTED]'
    return $Message
}

# ============================================================================
# WINDOW STATE PERSISTENCE
# ============================================================================
$Global:WinStateFile = ".\AVD-Manager-WinState.json"

function Save-WindowState {
    try {
        $state = @{
            Left   = [int]$Window.Left
            Top    = [int]$Window.Top
            Width  = [int]$Window.Width
            Height = [int]$Window.Height
            State  = $Window.WindowState.ToString()
        }
        $state | ConvertTo-Json | Set-Content $Global:WinStateFile -Encoding UTF8 -EA SilentlyContinue
    } catch {}
}

function Restore-WindowState {
    if (-not (Test-Path $Global:WinStateFile)) { return }
    try {
        $s = Get-Content $Global:WinStateFile -Raw | ConvertFrom-Json -EA Stop
        if ($s.State -eq "Normal") {
            if ($s.Left   -ge 0 -and $s.Left   -lt [System.Windows.SystemParameters]::VirtualScreenWidth)  { $Window.Left   = $s.Left }
            if ($s.Top    -ge 0 -and $s.Top    -lt [System.Windows.SystemParameters]::VirtualScreenHeight) { $Window.Top    = $s.Top }
            if ($s.Width  -ge 800)  { $Window.Width  = $s.Width }
            if ($s.Height -ge 600)  { $Window.Height = $s.Height }
        } elseif ($s.State -eq "Maximized") {
            $Window.WindowState = [System.Windows.WindowState]::Maximized
        }
    } catch {}
}

# ============================================================================
# CONFIGURATION EXPORT / IMPORT
# ============================================================================
function Export-ToolConfig {
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.FileName = "AVD-Manager-Config-$(Get-Date -f 'yyyyMMdd').json"
    $dlg.Filter   = "JSON Files (*.json)|*.json|All Files (*.*)|*.*"
    $dlg.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    if (-not $dlg.ShowDialog()) { return }
    try {
        $exportBy   = "unknown"; try { $exportBy   = (Get-AzContext).Account.Id }        catch {}
        $lastSub    = "";        try { $lastSub    = $script:SetSubId.Text }              catch {}
        $lastTenant = "";        try { $lastTenant = $script:SetTenantId.Text }           catch {}
        $lastLoc    = "eastus"; try { $lastLoc    = $script:SetLocation.SelectedValue }  catch {}
        $rdpPreset  = "";        try { $rdpPreset  = $script:RDPProps.Text }              catch {}
        $cfg = @{
            ExportDate    = (Get-Date -f "yyyy-MM-dd HH:mm:ss")
            ExportBy      = $exportBy
            LastSub       = $lastSub
            LastTenant    = $lastTenant
            LastLocation  = $lastLoc
            RefreshSecs   = $Global:RefreshSecs
            RDPPresets    = @{ Balanced = $rdpPreset }
        }
        $cfg | ConvertTo-Json -Depth 10 | Set-Content $dlg.FileName -Encoding UTF8
        Show-Toast "Configuration exported to $([System.IO.Path]::GetFileName($dlg.FileName))"
        Write-Log "Config exported: $($dlg.FileName)" "OK"
    } catch { Write-Log "Export config error: $_" "ERROR" }
}

function Import-ToolConfig {
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*"
    $dlg.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    if (-not $dlg.ShowDialog()) { return }
    try {
        $cfg = Get-Content $dlg.FileName -Raw | ConvertFrom-Json -EA Stop
        if ($cfg.LastSub)      { try{$script:SetSubId.Text      = $cfg.LastSub}catch{} }
        if ($cfg.LastTenant)   { try{$script:SetTenantId.Text   = $cfg.LastTenant}catch{} }
        if ($cfg.LastLocation) { try{$script:SetLocation.SelectedValue = $cfg.LastLocation}catch{} }
        if ($cfg.RefreshSecs)  { try{$script:SetRefresh.Value   = $cfg.RefreshSecs}catch{} }
        if ($cfg.RDPPresets.Balanced) { try{$script:RDPProps.Text = $cfg.RDPPresets.Balanced}catch{} }
        Show-Toast "Configuration imported from $([System.IO.Path]::GetFileName($dlg.FileName))"
        Write-Log "Config imported: $($dlg.FileName)" "OK"
    } catch { Write-Log "Import config error: $_" "ERROR" }
}

# ============================================================================
# PRE-DEPLOYMENT VALIDATION (quota, network, naming)
# ============================================================================
function Test-DeploymentReadiness {
    param([hashtable]$P)
    $issues = [System.Collections.Generic.List[string]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()

    # 1. Naming convention check
    if ($P.HpName -notmatch '^[a-zA-Z][a-zA-Z0-9\-]{0,62}[a-zA-Z0-9]$') {
        $issues.Add("Host pool name must start with a letter, be 2-64 chars, alphanumeric + hyphens")
    }
    if ($P.Prefix -notmatch '^[a-zA-Z][a-zA-Z0-9\-]{0,11}$') {
        $warnings.Add("VM prefix should be short (max 12 chars) to avoid OS hostname truncation on Win10")
    }

    # 2. VM quota check
    try {
        $quota = Get-AzVMUsage -Location $P.Loc -EA SilentlyContinue
        $cores  = $quota | Where-Object { $_.Name.Value -eq "cores" }
        $vmCores = @{ "Standard_D2s_v5"=2; "Standard_D4s_v5"=4; "Standard_D8s_v5"=8; "Standard_D16s_v5"=16;
                      "Standard_E4s_v5"=4; "Standard_E8s_v5"=8; "Standard_E16s_v5"=16;
                      "Standard_B4ms"=4;   "Standard_B8ms"=8 }
        $coreCount = $vmCores[$P.VmSz]
        if (-not $coreCount) { $coreCount = 4 }
        $needed = $P.VmCnt * $coreCount
        if ($cores) {
            $available = $cores.Limit - $cores.CurrentValue
            if ($available -lt $needed) {
                $issues.Add("Insufficient vCPU quota in $($P.Loc): need $needed, available $available. Request increase: https://portal.azure.com")
            } else {
                Write-Log "Quota OK: need $needed vCPUs, $available available in $($P.Loc)" "OK"
            }
        }
    } catch { $warnings.Add("Could not verify vCPU quota — check manually before deploying") }

    # 3. Network reachability (DNS check for AVD endpoint)
    try {
        $dns = [System.Net.Dns]::Resolve("rdweb.wvd.microsoft.com")
        Write-Log "DNS check OK: rdweb.wvd.microsoft.com resolves to $($dns.AddressList[0])" "OK"
    } catch { $warnings.Add("DNS resolution failed for rdweb.wvd.microsoft.com — verify outbound connectivity") }

    # 4. Subnet space check
    if ($P.SubnetName -and $P.VNetName -and $P.VNetRG) {
        try {
            $vnet   = Get-AzVirtualNetwork -Name $P.VNetName -ResourceGroupName $P.VNetRG -EA Stop
            $subnet = $vnet.Subnets | Where-Object { $_.Name -eq $P.SubnetName } | Select-Object -First 1
            if ($subnet) {
                # Count available IPs (Azure reserves 5 per subnet)
                $prefix  = $subnet.AddressPrefix[0]
                $bits    = [int]($prefix -split "/")[1]
                $total   = [Math]::Pow(2, 32-$bits) - 5
                $used    = $subnet.IpConfigurations.Count
                $avail   = $total - $used
                if ($avail -lt $P.VmCnt) {
                    $issues.Add("Subnet '$($P.SubnetName)' has ~$avail IPs available but you need $($P.VmCnt) for VMs (+NICs)")
                } else {
                    Write-Log "Subnet space OK: ~$avail IPs available, need $($P.VmCnt)" "OK"
                }
            }
        } catch { $warnings.Add("Could not verify subnet IP space — check subnet has enough addresses") }
    }

    # 5. Admin credential complexity (already validated in wizard, re-check here)
    if ($P.AdminPass.Length -lt 12) { $issues.Add("Admin password must be at least 12 characters") }

    # 6. Hybrid AD specific: DC reachability hint
    if ($P.JoinType -eq "HybridAD" -and -not $P.DomainName) {
        $issues.Add("Hybrid AD Join selected but Domain FQDN is empty")
    }

    # 7. Duplicate VM name check
    if ($P.VNetRG) {
        for ($i=0; $i -lt $P.VmCnt; $i++) {
            $vmName = "$($P.Prefix)-$i"
            $exists = Get-AzVM -ResourceGroupName $P.RG -Name $vmName -EA SilentlyContinue
            if ($exists) { $warnings.Add("VM '$vmName' already exists in RG '$($P.RG)' — it will be skipped") }
        }
    }

    # Report
    if ($issues.Count -gt 0 -or $warnings.Count -gt 0) {
        $msg = ""
        if ($issues.Count -gt 0) { $msg += "BLOCKING ISSUES:`n" + ($issues -join "`n") + "`n`n" }
        if ($warnings.Count -gt 0) { $msg += "WARNINGS (non-blocking):`n" + ($warnings -join "`n") }
        $btn = if ($issues.Count -gt 0) { [System.Windows.MessageBoxButton]::OK } else { [System.Windows.MessageBoxButton]::YesNo }
        $ico = if ($issues.Count -gt 0) { [System.Windows.MessageBoxImage]::Error } else { [System.Windows.MessageBoxImage]::Warning }
        $title = if ($issues.Count -gt 0) { "Deployment Blocked" } else { "Deployment Warnings" }
        $r = [System.Windows.MessageBox]::Show($msg, $title, $btn, $ico)
        if ($issues.Count -gt 0) { return $false }
        return ($r -eq [System.Windows.MessageBoxResult]::Yes)
    }
    return $true
}

# ============================================================================
# BEST PRACTICES ASSESSMENT ENGINE
# ============================================================================
$Script:BPRows = [System.Collections.Generic.List[PSObject]]::new()

function Add-BPRow {
    param([string]$Status, [string]$Category, [string]$Check, [string]$Finding, [string]$Remediation, [string]$Priority="Medium")
    $statusBg = switch ($Status) { "PASS"{"#0A2010"} "WARN"{"#2A1A00"} "FAIL"{"#2A0A0A"} default{"#1E3A5F"} }
    $statusFg = switch ($Status) { "PASS"{"#10B981"} "WARN"{"#F59E0B"} "FAIL"{"#EF4444"} default{"#94A3B8"} }
    $Script:BPRows.Add([PSCustomObject]@{
        Status=$Status; StatusBg=$statusBg; StatusFg=$statusFg
        Category=$Category; Check=$Check; Finding=$Finding
        Remediation=$Remediation; Priority=$Priority
    })
}

function Invoke-BestPracticesCheck {
    if (-not $Global:IsConnected) { return }
    Set-Status "Running best practices assessment..." 10
    Write-Log "=== Best Practices Assessment starting ===" "STEP"
    $Script:BPRows.Clear()

    try {
        $hpAll    = @(Invoke-AzWithRetry { Get-AzWvdHostPool -EA Stop })
        $wsAll    = @(Get-AzWvdWorkspace -EA SilentlyContinue)
        $lawAll   = @(Get-AzOperationalInsightsWorkspace -EA SilentlyContinue)
        $scaleAll = @(Get-AzWvdScalingPlan -EA SilentlyContinue)

        # ── Category: Prerequisites ──────────────────────────────────────────
        $azVer = (Get-Module -ListAvailable Az.DesktopVirtualization | Sort-Object Version -Descending | Select-Object -First 1).Version
        if ($azVer -and [Version]$azVer -ge [Version]"3.1.0") {
            Add-BPRow "PASS" "Prerequisites" "Az.DesktopVirtualization version" "v$azVer installed" "" "Low"
        } else {
            Add-BPRow "WARN" "Prerequisites" "Az.DesktopVirtualization version" "v$azVer — v3.1.0+ recommended" "Run: Update-Module Az.DesktopVirtualization" "Medium"
        }

        foreach ($hp in $hpAll) {
            $rg    = ($hp.Id -split "/")[4]
            $hosts = @(Get-AzWvdSessionHost -HostPoolName $hp.Name -ResourceGroupName $rg -EA SilentlyContinue)
            $ags   = @(Get-AzWvdApplicationGroup -EA SilentlyContinue | Where-Object {$_.HostPoolArmPath -like "*/$($hp.Name)"})
            $name  = $hp.Name

            # ── Category: Host Pool Configuration ───────────────────────────
            # Start VM on Connect
            if ($hp.StartVMOnConnect) {
                Add-BPRow "PASS" "Host Pool Configuration" "Start VM on Connect: $name" "Enabled — idle VMs deallocate saving cost" "" "High"
            } else {
                Add-BPRow "FAIL" "Host Pool Configuration" "Start VM on Connect: $name" "Disabled — VMs run 24/7 increasing cost" "Enable: Update-AzWvdHostPool -StartVMOnConnect:`$true" "High"
            }

            # Max session limit vs VM size
            if ($hp.HostPoolType -eq "Pooled") {
                $maxSess = $hp.MaxSessionLimit
                if ($maxSess -ge 2 -and $maxSess -le 30) {
                    Add-BPRow "PASS" "Host Pool Configuration" "Max session limit: $name" "$maxSess sessions/host — within recommended range" "" "Medium"
                } elseif ($maxSess -gt 30) {
                    Add-BPRow "WARN" "Host Pool Configuration" "Max session limit: $name" "$maxSess sessions/host — may cause performance degradation" "Validate VM size can support this density; consider D8s_v5 or larger" "Medium"
                } else {
                    Add-BPRow "WARN" "Host Pool Configuration" "Max session limit: $name" "$maxSess sessions/host — very low density" "Increase limit or reduce VM count to improve cost efficiency" "Low"
                }

                # Scaling plan
                $hasScale = $scaleAll | Where-Object { $_.HostPoolReference.HostPoolArmPath -like "*/$name" }
                if ($hasScale) {
                    Add-BPRow "PASS" "Host Pool Configuration" "Scaling plan: $name" "Assigned: $($hasScale.Name)" "" "High"
                } else {
                    Add-BPRow "FAIL" "Host Pool Configuration" "Scaling plan: $name" "No scaling plan — hosts run 24/7" "Create a scaling plan with peak/off-peak schedules to reduce cost by 40-60%" "High"
                }
            }

            # RDP properties
            $rdp = $hp.CustomRdpProperty
            if ($rdp -match "screen-capture-protection:i:1") {
                Add-BPRow "PASS" "Identity and Security" "Screen capture protection: $name" "Enabled" "" "High"
            } else {
                Add-BPRow "FAIL" "Identity and Security" "Screen capture protection: $name" "Disabled — users can screenshot session content" "Add: screen-capture-protection:i:1 to RDP properties" "High"
            }
            if ($rdp -match "watermarking:i:1") {
                Add-BPRow "PASS" "Identity and Security" "Session watermarking: $name" "Enabled" "" "Medium"
            } else {
                Add-BPRow "WARN" "Identity and Security" "Session watermarking: $name" "Disabled — harder to trace data exfil via photos" "Add: watermarking:i:1 to RDP properties" "Medium"
            }
            if ($rdp -match "redirectclipboard:i:0" -or $rdp -notmatch "redirectclipboard") {
                Add-BPRow "WARN" "Identity and Security" "Clipboard redirection: $name" "Clipboard may be unrestricted" "Consider: redirectclipboard:i:0 for sensitive environments" "Medium"
            } else {
                Add-BPRow "PASS" "Identity and Security" "Clipboard redirection: $name" "Explicitly configured in RDP properties" "" "Medium"
            }

            # App group has workspace
            if ($ags.Count -eq 0) {
                Add-BPRow "FAIL" "Host Pool Configuration" "App groups: $name" "No app groups found" "Create a Desktop or RemoteApp app group and associate with a workspace" "High"
            } else {
                $inWorkspace = $wsAll | Where-Object { $ags | Where-Object {$_.Id -in $_.ApplicationGroupReference} }
                if ($inWorkspace) {
                    Add-BPRow "PASS" "Host Pool Configuration" "Workspace assignment: $name" "$($ags.Count) app group(s) in workspace" "" "High"
                } else {
                    Add-BPRow "WARN" "Host Pool Configuration" "Workspace assignment: $name" "App groups may not be linked to a workspace" "Assign app groups to a workspace so users can discover them" "High"
                }
            }

            # Validation environment
            if ($hp.ValidationEnvironment) {
                Add-BPRow "WARN" "Host Pool Configuration" "Validation environment: $name" "Marked as validation — receives early service updates" "Only use validation pools for testing, not production users" "Medium"
            }

            # ── Category: Session Hosts ──────────────────────────────────────
            if ($hosts.Count -eq 0) {
                Add-BPRow "WARN" "Session Hosts" "Session hosts: $name" "No session hosts registered" "Add session hosts to make this pool functional" "High"
            } else {
                $unhealthy = @($hosts | Where-Object { $_.Status -notin @("Available","Upgrading") })
                if ($unhealthy.Count -eq 0) {
                    Add-BPRow "PASS" "Session Hosts" "Host health: $name" "All $($hosts.Count) host(s) Available" "" "High"
                } else {
                    Add-BPRow "FAIL" "Session Hosts" "Host health: $name" "$($unhealthy.Count)/$($hosts.Count) host(s) unhealthy: $($unhealthy.Name -join ', ')" "Check VM health, restart or drain unhealthy hosts, review event logs" "High"
                }

                # Drain mode check
                $drained = @($hosts | Where-Object { $_.AllowNewSession -eq $false })
                if ($drained.Count -gt 0) {
                    Add-BPRow "WARN" "Session Hosts" "Drain mode: $name" "$($drained.Count) host(s) in drain mode" "Remove drain mode on hosts unless intentionally draining: $($drained.Name -join ', ')" "Medium"
                }

                # OS version check (look for Windows 10 multi-session on older builds)
                $oldOS = @($hosts | Where-Object { $_.OSVersion -and $_.OSVersion -match "10\.0\.19041" })
                if ($oldOS.Count -gt 0) {
                    Add-BPRow "WARN" "Session Hosts" "OS version: $name" "$($oldOS.Count) host(s) on older Win10 build (19041)" "Consider upgrading to Windows 11 multi-session for better performance and support lifecycle" "Medium"
                }

                # Recent heartbeat
                $stale = @($hosts | Where-Object { $_.LastHeartBeat -and (New-TimeSpan -Start $_.LastHeartBeat -End (Get-Date)).TotalMinutes -gt 15 })
                if ($stale.Count -gt 0) {
                    Add-BPRow "WARN" "Session Hosts" "Heartbeat: $name" "$($stale.Count) host(s) haven't reported in 15+ min" "Check VM power state and AVD agent health on: $($stale.Name -join ', ')" "High"
                }
            }

            # ── Category: Identity and Security ─────────────────────────────
            # Note: join type isn't directly on HP object; check first host
            if ($hosts.Count -gt 0) {
                $firstHost = $hosts[0]
                if ($firstHost.VirtualMachineId) {
                    $vmRG   = ($firstHost.VirtualMachineId -split "/")[4]
                    $vmName = ($firstHost.VirtualMachineId -split "/")[-1]
                    $vm     = Get-AzVM -ResourceGroupName $vmRG -Name $vmName -EA SilentlyContinue
                    if ($vm) {
                        $entraExt   = $vm.Extensions | Where-Object { $_.Publisher -eq "Microsoft.Azure.ActiveDirectory" }
                        $domainExt  = $vm.Extensions | Where-Object { $_.Publisher -eq "Microsoft.Compute" -and $_.TypePropertiesType -eq "JsonADDomainExtension" }
                        $intuneExt  = $vm.Extensions | Where-Object { $_.Publisher -eq "Microsoft.Azure.ActiveDirectory" -and $_.Settings.mdmId }
                        if ($entraExt) {
                            Add-BPRow "PASS" "Identity and Security" "Entra ID join: $name" "VMs are Entra ID joined" "" "High"
                            if ($intuneExt -or ($entraExt.Settings -match "mdmId")) {
                                Add-BPRow "PASS" "Identity and Security" "Intune enrollment: $name" "Intune MDM enrollment configured" "" "High"
                            } else {
                                Add-BPRow "WARN" "Identity and Security" "Intune enrollment: $name" "Entra ID joined but Intune MDM not detected" "Enable Intune enrollment for compliance policies and patch management" "High"
                            }
                        } elseif ($domainExt) {
                            Add-BPRow "PASS" "Identity and Security" "Hybrid AD join: $name" "VMs are domain joined" "" "Medium"
                            Add-BPRow "INFO" "Identity and Security" "Join type recommendation: $name" "Using Hybrid AD join" "Consider migrating to Entra ID join for cloud-native management and simpler ops" "Low"
                        }

                        # Disk encryption
                        $diskEnc = $vm.StorageProfile.OsDisk.ManagedDisk.DiskEncryptionSet
                        if ($diskEnc) {
                            Add-BPRow "PASS" "Identity and Security" "Disk encryption: $vmName" "Customer-managed key encryption enabled" "" "Medium"
                        } else {
                            Add-BPRow "WARN" "Identity and Security" "Disk encryption: $vmName" "Using platform-managed keys (PMK)" "Consider customer-managed keys (CMK) for regulated industries" "Low"
                        }
                    }
                }
            }
        }

        # ── Category: Monitoring ─────────────────────────────────────────────
        if ($lawAll.Count -gt 0) {
            Add-BPRow "PASS" "Monitoring" "Log Analytics workspace" "$($lawAll.Count) workspace(s) found" "" "High"
            foreach ($law in $lawAll) {
                if ($law.RetentionInDays -ge 30) {
                    Add-BPRow "PASS" "Monitoring" "Log retention: $($law.Name)" "$($law.RetentionInDays) day retention" "" "Medium"
                } else {
                    Add-BPRow "WARN" "Monitoring" "Log retention: $($law.Name)" "$($law.RetentionInDays) day retention — below 30 day minimum" "Increase retention for compliance: Update-AzOperationalInsightsWorkspace -RetentionInDays 90" "Medium"
                }
            }

            # Check if diagnostic settings exist on any host pool
            foreach ($hp in $hpAll) {
                try {
                    $diag = Get-AzDiagnosticSetting -ResourceId $hp.Id -EA SilentlyContinue
                    if ($diag) {
                        Add-BPRow "PASS" "Monitoring" "Diagnostics: $($hp.Name)" "Diagnostic settings configured" "" "High"
                    } else {
                        Add-BPRow "FAIL" "Monitoring" "Diagnostics: $($hp.Name)" "No diagnostic settings — AVD Insights blind" "Enable: Portal > Host Pool > Diagnostic settings > Add setting to LAW" "High"
                    }
                } catch {}
            }
        } else {
            Add-BPRow "FAIL" "Monitoring" "Log Analytics workspace" "No workspace found — no visibility into AVD health" "Create a Log Analytics workspace and enable AVD Insights/Diagnostics" "High"
        }

        # ── Category: Cost Optimization ──────────────────────────────────────
        $pooledNoScale = @($hpAll | Where-Object { $_.HostPoolType -eq "Pooled" -and -not ($scaleAll | Where-Object { $_.HostPoolReference.HostPoolArmPath -like "*/$($_.Name)" }) })
        if ($pooledNoScale.Count -eq 0) {
            Add-BPRow "PASS" "Cost Optimization" "Scaling plans coverage" "All pooled host pools have scaling plans" "" "High"
        } else {
            Add-BPRow "FAIL" "Cost Optimization" "Scaling plans coverage" "$($pooledNoScale.Count) pooled pool(s) without scaling plans" "Add scaling plans to: $($pooledNoScale.Name -join ', ')" "High"
        }

        # Summarise
        $pass = ($Script:BPRows | Where-Object {$_.Status -eq "PASS"}).Count
        $warn = ($Script:BPRows | Where-Object {$_.Status -eq "WARN"}).Count
        $fail = ($Script:BPRows | Where-Object {$_.Status -eq "FAIL"}).Count
        $total = $pass + $warn + $fail
        $score = if ($total -gt 0) { [Math]::Round(($pass * 100.0 + $warn * 50.0) / $total, 0) } else { 0 }

        try { $script:BPScore.Text = "$score"; $script:BPPass.Text = "$pass"; $script:BPWarn.Text = "$warn"; $script:BPFail.Text = "$fail" } catch {}
        $scoreColor = if ($score -ge 80) {"#10B981"} elseif ($score -ge 60) {"#F59E0B"} else {"#EF4444"}
        try { $script:BPScore.Foreground = Get-Brush $scoreColor } catch {}
        try { $script:GridBP.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[PSObject]]($Script:BPRows.ToArray()) } catch {}

        Write-Log "Best practices complete: Score=$score/100  PASS=$pass  WARN=$warn  FAIL=$fail" "OK"
        Show-Toast "Assessment complete — Score: $score/100" $scoreColor

    } catch { Write-Log "Best practices error: $_" "ERROR" }
    Set-Status "Ready" 0
}

# ============================================================================
# UTILITY HELPERS
# ============================================================================

# Retry wrapper for transient Azure errors (429, 503, 504)
function Invoke-AzWithRetry {
    param([scriptblock]$Action, [int]$MaxAttempts=3, [int]$DelaySeconds=3)
    for ($attempt=1; $attempt -le $MaxAttempts; $attempt++) {
        try { return (& $Action) }
        catch {
            $msg = $_.Exception.Message
            $isTransient = $msg -match "429|503|504|TooManyRequests|ServiceUnavailable|GatewayTimeout|throttl"
            if ($isTransient -and $attempt -lt $MaxAttempts) {
                Write-Log "Transient error (attempt $attempt/$MaxAttempts), retrying in $($DelaySeconds*$attempt)s..." "WARN"
                Start-Sleep -Seconds ($DelaySeconds * $attempt)
            } else { throw }
        }
    }
}

# Toast - 4-second auto-clearing status bar message
function Show-Toast {
    param([string]$Message, [string]$Color="#10B981")
    try { $script:TxtStatus.Text=$Message; $script:TxtStatus.Foreground=Get-Brush $Color } catch {}
    $t = [System.Windows.Threading.DispatcherTimer]::new()
    $t.Interval = [TimeSpan]::FromSeconds(4)
    $t.Add_Tick({ try{$script:TxtStatus.Text="Ready";$script:TxtStatus.Foreground=Get-Brush "#64748B"}catch{}; $this.Stop() })
    $t.Start()
}

# Confirmation dialog
function Confirm-Action {
    param([string]$Message, [string]$Title="Confirm Action")
    $r = [System.Windows.MessageBox]::Show($Message,$Title,
        [System.Windows.MessageBoxButton]::YesNo,[System.Windows.MessageBoxImage]::Warning)
    return ($r -eq [System.Windows.MessageBoxResult]::Yes)
}

# Export DataGrid to CSV via SaveFileDialog
function Export-GridData {
    param([System.Windows.Controls.DataGrid]$Grid, [string]$DefaultName="export")
    if (-not $Grid -or $Grid.Items.Count -eq 0) { Show-Toast "No data to export" "#F59E0B"; return }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.FileName = "$DefaultName-$(Get-Date -f 'yyyyMMdd-HHmm').csv"
    $dlg.Filter   = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $dlg.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    if (-not $dlg.ShowDialog()) { return }
    try {
        @($Grid.Items) | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
        Show-Toast "Exported $($Grid.Items.Count) rows to $([System.IO.Path]::GetFileName($dlg.FileName))"
        Write-Log "Exported to: $($dlg.FileName)" "OK"
    } catch { Write-Log "Export error: $_" "ERROR" }
}

# Wizard step input validation
function Test-WizardStep {
    param([int]$Step)
    $errs = [System.Collections.Generic.List[string]]::new()
    switch ($Step) {
        1 {
            $n = try{$script:WizHPName.Text.Trim()}catch{""}
            if (-not $n)                           { $errs.Add("Host Pool Name is required") }
            elseif ($n -notmatch '^[a-zA-Z0-9\-]{1,64}$') { $errs.Add("Name: alphanumeric + hyphens only, max 64 chars") }
            if (-not (try{$script:WizRG.SelectedItem}catch{$null}))     { $errs.Add("Resource Group is required") }
            if (-not (try{$script:WizRegion.SelectedValue}catch{$null})) { $errs.Add("Azure Region is required") }
        }
        3 {
            if (-not (try{$script:WizVNet.SelectedItem}catch{$null}))   { $errs.Add("Virtual Network is required") }
            if (-not (try{$script:WizSubnet.SelectedItem}catch{$null})) { $errs.Add("Subnet is required") }
            $admin = try{$script:WizAdminUser.Text.Trim()}catch{""}
            if (-not $admin) { $errs.Add("Admin Username is required") }
            $pass = try{$script:WizAdminPass.Password}catch{""}
            if ($pass.Length -lt 12) { $errs.Add("Password must be at least 12 characters") }
            elseif ($pass -notmatch '[A-Z]' -or $pass -notmatch '[0-9]') { $errs.Add("Password needs uppercase letter and number") }
        }
        4 {
            if (-not (try{$script:WizVMSize.SelectedItem}catch{$null})) { $errs.Add("VM Size is required") }
            $pfx = try{$script:WizVMPrefix.Text.Trim()}catch{""}
            if (-not $pfx) { $errs.Add("VM Name Prefix is required") }
        }
        5 {
            $hybrid = $false; try{$hybrid=$script:RdoHybrid.IsChecked}catch{}
            if ($hybrid) {
                if (-not (try{$script:WizDomain.Text.Trim()}catch{""}))     { $errs.Add("Domain FQDN is required for Hybrid AD Join") }
                if (-not (try{$script:WizDomainUser.Text.Trim()}catch{""})) { $errs.Add("Domain Join Account is required") }
                if (-not (try{$script:WizDomainPass.Password}catch{""}))    { $errs.Add("Domain Join Password is required") }
            }
        }
    }
    if ($errs.Count -gt 0) {
        [System.Windows.MessageBox]::Show(
            "Please fix before continuing:`n`n" + ($errs -join "`n"),
            "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return $false
    }
    return $true
}

# ============================================================================
# USER SESSION MANAGEMENT
# ============================================================================
function Load-UserSessions {
    if (-not $Global:IsConnected) { return }
    Set-Status "Loading user sessions..." 30
    try {
        $filter = try{$script:SessFilter.SelectedItem}catch{"(All Pools)"}
        $hpAll  = @(Invoke-AzWithRetry { Get-AzWvdHostPool -EA Stop })
        $rows   = [System.Collections.Generic.List[PSObject]]::new()
        foreach ($hp in $hpAll) {
            if ($filter -and $filter -ne "(All Pools)" -and $hp.Name -ne $filter) { continue }
            $rg = ($hp.Id -split "/")[4]
            $sessions = @(Get-AzWvdUserSession -HostPoolName $hp.Name -ResourceGroupName $rg -EA SilentlyContinue)
            foreach ($s in $sessions) {
                $hostName = ($s.Name -split "/")[1]
                $state    = if ($s.SessionState) {$s.SessionState} else {"Connected"}
                $stateBg  = switch ($state) { "Active"{"#0A2010"} "Disconnected"{"#2A0A0A"} "Pending"{"#2A1A00"} default{"#1E3A5F"} }
                $stateFg  = switch ($state) { "Active"{"#10B981"} "Disconnected"{"#EF4444"} "Pending"{"#F59E0B"} default{"#94A3B8"} }
                $rows.Add([PSCustomObject]@{
                    UserUPN  =$s.UserPrincipalName; PoolName=$hp.Name; HostName=$hostName
                    SessionId=($s.Name -split "/")[-1]; State=$state; StateBg=$stateBg; StateFg=$stateFg
                    App=if($s.ApplicationType -eq "RemoteApp"){"RemoteApp"}else{"Full Desktop"}
                    LoginTime=if($s.CreateTime){$s.CreateTime.ToString("MM/dd HH:mm")}else{"--"}
                    RG=$rg; HPName=$hp.Name; SessObj=$s
                })
            }
        }
        try { $script:GridUserSess.ItemsSource=[System.Collections.ObjectModel.ObservableCollection[PSObject]]($rows.ToArray()) } catch {}
        Write-Log "User sessions: $($rows.Count) active" "OK"
    } catch { Write-Log "Load-UserSessions error: $_" "WARN" }
    Set-Status "Ready" 0
}

function Send-UserMessage {
    param($Row)
    if (-not $Row) { Show-Toast "Select a user session first" "#F59E0B"; return }
    $msg = [Microsoft.VisualBasic.Interaction]::InputBox("Message to $($Row.UserUPN):", "Send Message", "")
    if (-not $msg) { return }
    try {
        Send-AzWvdUserSessionMessage -HostPoolName $Row.HPName -ResourceGroupName $Row.RG `
            -SessionHostName $Row.HostName -UserSessionId $Row.SessionId `
            -MessageTitle "IT Administrator" -MessageBody $msg -EA Stop | Out-Null
        Show-Toast "Message sent to $($Row.UserUPN)"
        Write-Log "Message sent to $($Row.UserUPN)" "OK"
    } catch { Write-Log "Send message error: $_" "ERROR"; Show-Toast "Send failed" "#EF4444" }
}

function Invoke-Disconnect {
    param($Row)
    if (-not $Row) { Show-Toast "Select a user session first" "#F59E0B"; return }
    if (-not (Confirm-Action "Disconnect $($Row.UserUPN)?`nSession remains in Disconnected state.")) { return }
    try {
        Disconnect-AzWvdUserSession -HostPoolName $Row.HPName -ResourceGroupName $Row.RG `
            -SessionHostName $Row.HostName -Id $Row.SessionId -EA Stop | Out-Null
        Show-Toast "Disconnected $($Row.UserUPN)"; Write-Log "Disconnected: $($Row.UserUPN)" "OK"
        Load-UserSessions
    } catch { Write-Log "Disconnect error: $_" "ERROR"; Show-Toast "Disconnect failed" "#EF4444" }
}

function Invoke-ForceLogoff {
    param($Row)
    if (-not $Row) { Show-Toast "Select a user session first" "#F59E0B"; return }
    if (-not (Confirm-Action "Force logoff $($Row.UserUPN)?`nUnsaved work will be LOST.")) { return }
    try {
        Remove-AzWvdUserSession -HostPoolName $Row.HPName -ResourceGroupName $Row.RG `
            -SessionHostName $Row.HostName -Id $Row.SessionId -Force -EA Stop | Out-Null
        Show-Toast "Logged off $($Row.UserUPN)"; Write-Log "Force logoff: $($Row.UserUPN)" "OK"
        Load-UserSessions
    } catch { Write-Log "Logoff error: $_" "ERROR"; Show-Toast "Logoff failed" "#EF4444" }
}

# ============================================================================
# SUBSCRIPTION SWITCHER
# ============================================================================
function Load-Subscriptions {
    try {
        $subs = @(Get-AzSubscription -EA SilentlyContinue)
        $items = @($subs | ForEach-Object {"$($_.Name)  [$($_.Id.Substring(0,8))...]"})
        try { $script:CboSubSwitcher.ItemsSource = $items } catch {}
        $ctx = Get-AzContext -EA SilentlyContinue
        if ($ctx) {
            $idx = 0
            for ($i=0; $i -lt $subs.Count; $i++) {
                if ($subs[$i].Id -eq $ctx.Subscription.Id) { $idx=$i; break }
            }
            try { $script:CboSubSwitcher.SelectedIndex = $idx } catch {}
        }
        $Global:SubList = $subs
    } catch {}
}

# ============================================================================
# REAL COST API (Azure Cost Management)
# ============================================================================
function Get-AzRealCost {
    param([string]$Scope, [int]$Days=30)
    if (-not $Global:IsConnected) { return $null }
    try {
        $tokenObj = Get-AzAccessToken -ResourceUrl "https://management.azure.com" -EA Stop
        $token = if ($tokenObj.Token -is [System.Security.SecureString]) {
            $b=[System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($tokenObj.Token)
            $t=[System.Runtime.InteropServices.Marshal]::PtrToStringAuto($b)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($b); $t
        } else { $tokenObj.Token }
        $end   = (Get-Date).ToString("yyyy-MM-dd")
        $start = (Get-Date).AddDays(-$Days).ToString("yyyy-MM-dd")
        $body  = @{
            type="ActualCost"; timeframe="Custom"
            timePeriod=@{from=$start;to=$end}
            dataset=@{granularity="None";aggregation=@{totalCost=@{name="Cost";function="Sum"}}}
        } | ConvertTo-Json -Depth 10
        $hdr  = @{Authorization="Bearer $token";"Content-Type"="application/json"}
        $uri  = "https://management.azure.com/$Scope/providers/Microsoft.CostManagement/query?api-version=2023-03-01"
        $resp = Invoke-RestMethod -Uri $uri -Method Post -Headers $hdr -Body $body -EA Stop
        return [Math]::Round(($resp.properties.rows | ForEach-Object {$_[0]} | Measure-Object -Sum).Sum, 2)
    } catch { Write-Log "Cost API: $_" "WARN"; return $null }
}

# ============================================================================
# LOG ANALYTICS / AVD INSIGHTS
# ============================================================================
function Get-AVDInsights {
    param([string]$WorkspaceId)
    if (-not $WorkspaceId -or -not $Global:IsConnected) { return $null }
    try {
        $tokenObj = Get-AzAccessToken -ResourceUrl "https://api.loganalytics.io" -EA Stop
        $token = if ($tokenObj.Token -is [System.Security.SecureString]) {
            $b=[System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($tokenObj.Token)
            $t=[System.Runtime.InteropServices.Marshal]::PtrToStringAuto($b)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($b); $t
        } else { $tokenObj.Token }
        $hdr = @{Authorization="Bearer $token";"Content-Type"="application/json"}
        $queries = @{
            Connections = "WVDConnections | where TimeGenerated > ago(24h) | summarize count()"
            Errors      = "WVDErrors | where TimeGenerated > ago(24h) | summarize count()"
            TopUsers    = "WVDConnections | where TimeGenerated > ago(7d) | summarize Sessions=count() by UserName | top 5 by Sessions desc"
        }
        $results = @{}
        foreach ($q in $queries.GetEnumerator()) {
            try {
                $body = @{query=$q.Value} | ConvertTo-Json
                $r    = Invoke-RestMethod -Uri "https://api.loganalytics.io/v1/workspaces/$WorkspaceId/query" -Method Post -Headers $hdr -Body $body -EA Stop
                $results[$q.Key] = $r.tables[0].rows
            } catch { $results[$q.Key] = $null }
        }
        return $results
    } catch { Write-Log "AVD Insights error: $_" "WARN"; return $null }
}

# ============================================================================
# GRAPH TOKEN CACHE (auto-refresh before expiry)
# ============================================================================
$Global:GraphToken       = $null
$Global:GraphTokenExpiry = [DateTime]::MinValue

function Get-GraphToken {
    if ($Global:GraphToken -and [DateTime]::UtcNow -lt $Global:GraphTokenExpiry.AddMinutes(-5)) {
        return $Global:GraphToken
    }
    $obj = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com" -EA Stop
    $raw = if ($obj.Token -is [System.Security.SecureString]) {
        $b=[System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($obj.Token)
        $t=[System.Runtime.InteropServices.Marshal]::PtrToStringAuto($b)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($b); $t
    } else { $obj.Token }
    $Global:GraphToken       = $raw
    $Global:GraphTokenExpiry = if ($obj.ExpiresOn) {$obj.ExpiresOn.UtcDateTime} else {[DateTime]::UtcNow.AddHours(1)}
    return $Global:GraphToken
}

# ============================================================================
# ACTUAL SESSION HOST VM DEPLOYMENT
# ============================================================================
function Deploy-SessionHosts {
    param([hashtable]$P)
    function DLog { param($m,$l="DEPLOY") $P.Sync.LogQueue.Enqueue(@{Timestamp=(Get-Date -f "HH:mm:ss");Level=$l;Message=$m}) }
    function SP   { param([int]$p) $P.Sync.Progress=$p }

    $imgMap = @{
        "Windows 11 Enterprise multi-session + Microsoft 365 Apps (latest) [RECOMMENDED]"=@{Pub="MicrosoftWindowsDesktop";Off="office-365";Sku="win11-23h2-avd-m365"}
        "Windows 11 Enterprise multi-session (latest)"                                   =@{Pub="MicrosoftWindowsDesktop";Off="windows-11";Sku="win11-23h2-avd"}
        "Windows 10 Enterprise multi-session + Microsoft 365 Apps (latest)"              =@{Pub="MicrosoftWindowsDesktop";Off="office-365";Sku="win10-22h2-avd-m365"}
        "Windows 10 Enterprise multi-session (latest)"                                   =@{Pub="MicrosoftWindowsDesktop";Off="windows-10";Sku="win10-22h2-avd"}
        "Windows Server 2022 Datacenter Azure Edition"                                   =@{Pub="MicrosoftWindowsServer";Off="WindowsServer";Sku="2022-datacenter-azure-edition"}
        "Windows Server 2019 Datacenter Azure Edition"                                   =@{Pub="MicrosoftWindowsServer";Off="WindowsServer";Sku="2019-datacenter-azure-edition"}
    }
    $img = $imgMap[$P.ImageName]
    if (-not $img) { $img = @{Pub="MicrosoftWindowsDesktop";Off="office-365";Sku="win11-23h2-avd-m365"} }

    $secPass = ConvertTo-SecureString $P.AdminPass -AsPlainText -Force
    $cred    = New-Object System.Management.Automation.PSCredential($P.AdminUser, $secPass)

    # Resolve subnet object
    DLog "Resolving network: $($P.VNetRG) / $($P.VNetName) / $($P.SubnetName)..." "STEP"
    $vnetObj   = Get-AzVirtualNetwork -Name $P.VNetName -ResourceGroupName $P.VNetRG -EA Stop
    $subnetObj = $vnetObj.Subnets | Where-Object {$_.Name -eq $P.SubnetName} | Select-Object -First 1
    if (-not $subnetObj) { throw "Subnet '$($P.SubnetName)' not found in VNet '$($P.VNetName)'" }

    $baseProgress = 65
    for ($i=0; $i -lt $P.VmCount; $i++) {
        if ($P.Sync.CancelToken) { DLog "Cancelled at VM $i" "WARN"; return }
        $vmName = "$($P.Prefix)-$i"
        $step   = $i + 1
        DLog "VM $step/$($P.VmCount): $vmName ..." "STEP"
        SP ($baseProgress + [Math]::Round($i * 30.0 / [Math]::Max($P.VmCount,1), 0))

        try {
            # NIC
            $ipCfg = New-AzNetworkInterfaceIpConfig -Name "ipconfig1" -SubnetId $subnetObj.Id -Primary
            $nic   = New-AzNetworkInterface -Name "$vmName-nic" -ResourceGroupName $P.RG -Location $P.Loc -IpConfiguration $ipCfg -EA Stop
            DLog "  [$step] NIC: $vmName-nic" "OK"

            # VM config
            $vmCfg = New-AzVMConfig -VMName $vmName -VMSize $P.VmSz |
                     Set-AzVMOperatingSystem -Windows -ComputerName $vmName -Credential $cred -ProvisionVMAgent -EnableAutoUpdate |
                     Set-AzVMSourceImage -PublisherName $img.Pub -Offer $img.Off -Skus $img.Sku -Version "latest" |
                     Add-AzVMNetworkInterface -Id $nic.Id |
                     Set-AzVMOSDisk -CreateOption FromImage -StorageAccountType "Premium_LRS"
            if ($P.HybridBenefit) { $vmCfg = $vmCfg | Set-AzVMLicenseType -LicenseType "Windows_Client" }

            DLog "  [$step] Creating VM (3-5 min)..." "INFO"
            New-AzVM -ResourceGroupName $P.RG -Location $P.Loc -VM $vmCfg -EA Stop | Out-Null
            DLog "  [$step] VM provisioned: $vmName" "OK"

            # Entra ID join
            if ($P.JoinType -eq "EntraID") {
                $mdmId = if ($P.IntuneEnroll) {"0000000a-0000-0000-c000-000000000000"} else {""}
                Set-AzVMExtension -ResourceGroupName $P.RG -VMName $vmName `
                    -Name "AADLoginForWindows" -Publisher "Microsoft.Azure.ActiveDirectory" `
                    -ExtensionType "AADLoginForWindows" -TypeHandlerVersion "1.0" `
                    -Settings @{mdmId=$mdmId} -EA Stop | Out-Null
                DLog "  [$step] Entra ID join extension OK" "OK"
            }

            # Hybrid AD join
            if ($P.JoinType -eq "HybridAD") {
                $djSet  = @{Name=$P.DomainName;User=$P.DomainUser;Restart="true";Options="3"}
                $djProt = @{Password=$P.DomainPass}
                if ($P.OUPath) { $djSet.OUPath = $P.OUPath }
                Set-AzVMExtension -ResourceGroupName $P.RG -VMName $vmName `
                    -Name "JsonADDomainExtension" -Publisher "Microsoft.Compute" `
                    -ExtensionType "JsonADDomainExtension" -TypeHandlerVersion "1.3" `
                    -Settings $djSet -ProtectedSettings $djProt -EA Stop | Out-Null
                DLog "  [$step] Domain joined: $($P.DomainName)" "OK"
            }

            # AVD DSC registration extension
            Set-AzVMExtension -ResourceGroupName $P.RG -VMName $vmName `
                -Name "DSC" -Publisher "Microsoft.Powershell" -ExtensionType "DSC" `
                -TypeHandlerVersion "2.73" `
                -Settings @{hostPoolName=$P.HpName; registrationInfoToken=$P.RegToken; aadJoin=($P.JoinType -eq "EntraID")} `
                -EA Stop | Out-Null
            DLog "  [$step] AVD agent registered with host pool" "OK"

        } catch { DLog "  [$step] ERROR on ${vmName}: $_" "ERROR" }
    }
    SP 100
    DLog "Session host deployment complete: $($P.VmCount) VM(s)" "OK"
}

# ============================================================================
# APP GROUP USER ASSIGNMENT
# ============================================================================
function Set-AppGroupUsers {
    param([string]$AppGroupId, [string[]]$UserUPNs)
    foreach ($upn in $UserUPNs) {
        try {
            $uid = (Get-AzADUser -UserPrincipalName $upn -EA Stop).Id
            New-AzRoleAssignment -ObjectId $uid -RoleDefinitionName "Desktop Virtualization User" `
                -Scope $AppGroupId -EA Stop | Out-Null
            Write-Log "Assigned $upn to app group" "OK"
        } catch { Write-Log "Assignment error for ${upn}: $_" "WARN" }
    }
}

# ============================================================================
# DEPLOYMENT
# ============================================================================
function Start-AVDDeployment {
    if (-not $Global:IsConnected) { [System.Windows.MessageBox]::Show("Connect to Azure first.", "AVD Manager") | Out-Null; return }
    if ($Global:Sync.IsDeploying)  { [System.Windows.MessageBox]::Show("Deployment already running.", "AVD Manager") | Out-Null; return }

    $hpName  = ""; try { $hpName  = $script:WizHPName.Text } catch {}
    $rg      = ""; try { $rg      = $script:WizRG.SelectedItem } catch {}
    $loc     = "eastus"; try { $loc = $script:WizRegion.SelectedValue } catch {}
    $vmCnt   = 2; try { $vmCnt   = [int]$script:SliderVMCnt.Value } catch {}
    $prefix  = "avd-host"; try { $prefix = $script:WizVMPrefix.Text } catch {}
    $vmSzRaw = ""; try { $vmSzRaw = $script:WizVMSize.SelectedItem } catch {}
    $vmSz    = if ($vmSzRaw) { ($vmSzRaw -split " ")[0] } else { "Standard_D4s_v5" }
    $maxS    = 8; try { $maxS  = [int]$script:SliderMaxSess.Value } catch {}
    $startVM = $true; try { $startVM = [bool]$script:ChkStartVM.IsChecked } catch {}
    $hpType  = "Pooled"
    try { if ($script:RdoPersonal.IsChecked)  { $hpType = "Personal"  } } catch {}
    try { if ($script:RdoRemoteApp.IsChecked) { $hpType = "Pooled"    } } catch {}
    $lb = "BreadthFirst"
    try { if ($script:RdoDepth.IsChecked) { $lb = "DepthFirst" } } catch {}
    $agType = "Desktop"
    try { if ($script:RdoRemoteApp.IsChecked) { $agType = "RailApplications" } } catch {}

    if (-not $hpName -or -not $rg) { [System.Windows.MessageBox]::Show("Host Pool Name and Resource Group are required.", "AVD Manager") | Out-Null; return }

    # Pre-deployment readiness check
    Set-Status "Running pre-deployment validation..." 10
    $preCheck = @{ HpName=$hpName; RG=$rg; Loc=$loc; VmCnt=$vmCnt; Prefix=$prefix; VmSz=$vmSz
                   JoinType=$joinType; DomainName=$domainName; AdminPass=$adminPass
                   VNetRG=$vnetRG; VNetName=$vnetName; SubnetName=$subnetName }
    if (-not (Test-DeploymentReadiness $preCheck)) { Set-Status "Ready" 0; return }

    # Audit the deployment initiation
    Write-Audit "DEPLOY_START" "HostPool/$hpName" "Type=$hpType VMs=$vmCnt x $vmSz Region=$loc Join=$joinType" "OK"

    # Gather additional deployment parameters
    $vmSzRawFull  = try{$script:WizVMSize.SelectedItem}catch{""}
    $imgName      = try{$script:WizMktImg.SelectedItem}catch{""}
    $vnetRG       = try{$script:WizVNetRG.SelectedItem}catch{""}
    $vnetName     = try{$script:WizVNet.SelectedItem}catch{""}
    $subnetName   = try{$script:WizSubnet.SelectedItem}catch{""}
    $adminPass    = try{$script:WizAdminPass.Password}catch{""}
    $joinType     = "EntraID"; try{if($script:RdoHybrid.IsChecked){$joinType="HybridAD"}}catch{}
    $intuneEnroll = $true;  try{$intuneEnroll=[bool]$script:ChkIntune.IsChecked}catch{}
    $hybridBen    = $false; try{$hybridBen=[bool]$script:ChkHybridBenefit.IsChecked}catch{}
    $domainName   = try{$script:WizDomain.Text}catch{""}
    $domainUser   = try{$script:WizDomainUser.Text}catch{""}
    $domainPass   = try{$script:WizDomainPass.Password}catch{""}
    $ouPath       = try{$script:WizOU.Text}catch{""}

    $Global:Sync.IsDeploying = $true; $Global:Sync.CancelToken = $false; $Global:Sync.Progress = 0
    try { $script:BtnDeploy.IsEnabled=$false; $script:BtnCancelDeploy.IsEnabled=$true } catch {}
    try { $script:WizLog.Document.Blocks.Clear() } catch {}
    $Global:WizLogBox = $null; try { $Global:WizLogBox = $script:WizLog } catch {}

    Write-Log "=== Deploying AVD: $hpName ($hpType) ===" "STEP"
    Write-Log "Region: $loc | VMs: $vmCnt x $vmSz | MaxSessions: $maxS | Join: $joinType" "INFO"

    $dv = @{
        HpName=$hpName; RG=$rg; Loc=$loc; VmCnt=$vmCnt; Prefix=$prefix; VmSz=$vmSz
        MaxS=$maxS; HpType=$hpType; LB=$lb; AgType=$agType; StartVM=$startVM
        ImageName=$imgName; VNetRG=$vnetRG; VNetName=$vnetName; SubnetName=$subnetName
        AdminUser=$($script:WizAdminUser.Text); AdminPass=$adminPass
        JoinType=$joinType; IntuneEnroll=$intuneEnroll; HybridBenefit=$hybridBen
        DomainName=$domainName; DomainUser=$domainUser; DomainPass=$domainPass; OUPath=$ouPath
        Sync=$Global:Sync; RegToken=""
    }

    $deployScript = {
        param($v)
        function DLog { param([string]$m,[string]$l="DEPLOY") $v.Sync.LogQueue.Enqueue(@{Timestamp=(Get-Date -f "HH:mm:ss");Level=$l;Message=$m}) }
        function SP   { param([int]$p) $v.Sync.Progress=$p }
        try {
            # 1. Resource Group
            DLog "Step 1/5: Resource group '$($v.RG)' in $($v.Loc)..." "STEP"
            SP 10
            if (-not (Get-AzResourceGroup -Name $v.RG -EA SilentlyContinue)) {
                New-AzResourceGroup -Name $v.RG -Location $v.Loc -EA Stop | Out-Null
                DLog "  Created: $($v.RG)" "OK"
            } else { DLog "  Exists: $($v.RG)" "INFO" }
            if ($v.Sync.CancelToken) { DLog "Cancelled." "WARN"; return }

            # 2. Host Pool
            DLog "Step 2/5: Creating host pool '$($v.HpName)'..." "STEP"
            SP 25
            if (-not (Get-AzWvdHostPool -Name $v.HpName -ResourceGroupName $v.RG -EA SilentlyContinue)) {
                $exp = (Get-Date).ToUniversalTime().AddHours(48).ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
                New-AzWvdHostPool -Name $v.HpName -ResourceGroupName $v.RG -Location $v.Loc `
                    -HostPoolType $v.HpType -LoadBalancerType $v.LB -PreferredAppGroupType $v.AgType `
                    -MaxSessionLimit $v.MaxS -StartVMOnConnect $v.StartVM `
                    -ExpirationTime $exp -RegistrationTokenOperation Update -EA Stop | Out-Null
                DLog "  Host pool created: $($v.HpName)" "OK"
            } else { DLog "  Exists: $($v.HpName)" "INFO" }
            if ($v.Sync.CancelToken) { DLog "Cancelled." "WARN"; return }

            # 3. Workspace and App Group
            DLog "Step 3/5: Workspace and app group..." "STEP"
            SP 45
            $wsName = "ws-$($v.HpName)"; $agName = "$($v.HpName)-ag"
            $ws = Get-AzWvdWorkspace -ResourceGroupName $v.RG -EA SilentlyContinue | Where-Object {$_.Name -eq $wsName} | Select-Object -First 1
            if (-not $ws) {
                $ws = New-AzWvdWorkspace -Name $wsName -ResourceGroupName $v.RG -Location $v.Loc -FriendlyName "Workspace for $($v.HpName)" -EA Stop
                DLog "  Workspace created: $wsName" "OK"
            } else { DLog "  Workspace exists: $wsName" "INFO" }
            if (-not (Get-AzWvdApplicationGroup -Name $agName -ResourceGroupName $v.RG -EA SilentlyContinue)) {
                $hp = Get-AzWvdHostPool -Name $v.HpName -ResourceGroupName $v.RG -EA Stop
                $agTypeStr = if ($v.AgType -eq "RailApplications") {"RemoteApp"} else {"Desktop"}
                $ag = New-AzWvdApplicationGroup -Name $agName -ResourceGroupName $v.RG -Location $v.Loc `
                    -HostPoolArmPath $hp.Id -ApplicationGroupType $agTypeStr -FriendlyName "Apps for $($v.HpName)" -EA Stop
                $refs = @($ws.ApplicationGroupReference) + $ag.Id
                Update-AzWvdWorkspace -Name $wsName -ResourceGroupName $v.RG -ApplicationGroupReference $refs -EA SilentlyContinue | Out-Null
                DLog "  App group created: $agName ($agTypeStr)" "OK"
            } else { DLog "  App group exists: $agName" "INFO" }
            if ($v.Sync.CancelToken) { DLog "Cancelled." "WARN"; return }

            # 4. Registration Token
            DLog "Step 4/6: Generating registration token..." "STEP"
            SP 60
            $exp = (Get-Date).ToUniversalTime().AddHours(48).ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
            $regInfo = New-AzWvdRegistrationInfo -HostPoolName $v.HpName -ResourceGroupName $v.RG `
                -ExpirationTime $exp -RegistrationTokenOperation Update -EA Stop
            $v.RegToken = $regInfo.Token
            DLog "  Token generated (valid 48 hrs)" "OK"
            if ($v.Sync.CancelToken) { DLog "Cancelled." "WARN"; return }

            # 5. Deploy Session Host VMs
            DLog "Step 5/6: Deploying $($v.VmCnt) session host VM(s)..." "STEP"
            if ($v.VNetName -and $v.SubnetName -and $v.AdminUser -and $v.AdminPass) {
                Deploy-SessionHosts -P $v
            } else {
                DLog "  Skipping VM creation — network/credentials not fully configured" "WARN"
                DLog "  Registration token: $($regInfo.Token.Substring(0,[Math]::Min(60,$regInfo.Token.Length)))..." "INFO"
                DLog "  Use the token above to manually register VMs via Portal or ARM template" "INFO"
            }
            if ($v.Sync.CancelToken) { DLog "Cancelled." "WARN"; return }

            # 6. Post-deployment
            SP 100
            DLog "" "INFO"
            DLog "=== Deployment complete ===" "OK"
            DLog "Host Pool : $($v.HpName) ($($v.HpType), $($v.LB))" "OK"
            DLog "App Group : $agName" "OK"
            DLog "Workspace : $wsName" "OK"
            DLog "" "INFO"
            DLog "Post-deployment checklist:" "INFO"
            DLog "  [1] Verify session hosts appear in host pool within 5-10 min" "INFO"
            DLog "  [2] Assign users: Portal > App Groups > $agName > Assignments" "INFO"
            DLog "  [3] Configure FSLogix registry on hosts if not done via GPO" "INFO"
            DLog "  [4] Test a user connection via https://client.wvd.microsoft.com" "INFO"
        } catch {
            DLog "DEPLOYMENT FAILED: $_" "ERROR"; SP 0
        }
        $v.Sync.IsDeploying = $false
    }

    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace(); $rs.Open()
    $ps = [System.Management.Automation.PowerShell]::Create(); $ps.Runspace=$rs
    $ps.AddScript($deployScript).AddArgument($dv) | Out-Null
    $handle = $ps.BeginInvoke()

    # Monitor completion
    $monRS = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace(); $monRS.Open()
    $monPS = [System.Management.Automation.PowerShell]::Create(); $monPS.Runspace=$monRS
    $monPS.AddScript({ param($p,$h,$sync)
        while (-not $h.IsCompleted) { Start-Sleep -Milliseconds 400 }
        try { $p.EndInvoke($h) | Out-Null; $p.Dispose() } catch {}
        $sync.IsDeploying=$false
    }).AddArgument($ps).AddArgument($handle).AddArgument($Global:Sync) | Out-Null
    $monPS.BeginInvoke() | Out-Null
}

# ============================================================================
# TIMER
# ============================================================================
$Script:Countdown = $Global:RefreshSecs
$Script:Timer = [System.Windows.Threading.DispatcherTimer]::new()
$Script:Timer.Interval = [TimeSpan]::FromSeconds(1)
$Script:Timer.Add_Tick({
    try {
        Flush-UILog
        if ($Global:Sync.IsDeploying) {
            try { $script:DeployProg.Value    = [double]$Global:Sync.Progress } catch {}
            try { $script:TxtDeployPhase.Text = "Deploying... $($Global:Sync.Progress)%" } catch {}
        }
        $Script:Countdown--
        if ($Script:Countdown -le 0) {
            $Script:Countdown = $Global:RefreshSecs
            try { $script:TxtCountdown.Text = "" } catch {}
        } else {
            try { $script:TxtCountdown.Text = "Refresh in $($Script:Countdown)s" } catch {}
        }
    } catch {}
})

# ============================================================================
# EVENT HANDLERS
# ============================================================================

# Connect
$BtnConnect.Add_Click({
    if ($Global:IsConnected) {
        Disconnect-AzAccount -EA SilentlyContinue | Out-Null
        Set-Connected $false; Set-Status "Disconnected"
        Write-Log "Disconnected from Azure" "WARN"; return
    }
    Set-Status "Connecting to Azure..." 50
    Write-Log "Connecting to Azure..." "STEP"
    try {
        $ctx = Get-AzContext -EA SilentlyContinue
        if (-not $ctx) {
            $tid = try{$script:SetTenantId.Text}catch{""}
            if ($tid) { Connect-AzAccount -TenantId $tid -EA Stop | Out-Null } else { Connect-AzAccount -EA Stop | Out-Null }
            $ctx = Get-AzContext
        }
        $sid = try{$script:SetSubId.Text}catch{""}
        if ($sid) { Set-AzContext -SubscriptionId $sid -EA SilentlyContinue | Out-Null; $ctx = Get-AzContext }
        $Global:Subscription = $ctx.Subscription; $Global:TenantId = $ctx.Tenant.Id
        Set-Connected $true; Set-Status "Connected: $($ctx.Subscription.Name)" 0
        Write-Log "Connected: $($ctx.Account.Id) | Sub: $($ctx.Subscription.Name)" "OK"
        Load-Subscriptions   # populate subscription switcher
        Load-Dashboard
        $autoLic = $true; try { $autoLic = [bool]$script:SetAutoLic.IsChecked } catch {}
        if ($autoLic) { Invoke-LicenseScan }
        # Real cost estimate in background
        $subScope = "/subscriptions/$($ctx.Subscription.Id)"
        $cost = Get-AzRealCost -Scope $subScope -Days 30
        if ($cost -ne $null) {
            Write-Log "Azure cost last 30 days (VMs + AVD): `$$cost" "INFO"
        }
    } catch {
        Set-Status "Connection failed" 0; Write-Log "Connection failed: $_" "ERROR"
        [System.Windows.MessageBox]::Show("Connection failed:`n$_`n`nCheck credentials and network.", "AVD Manager") | Out-Null
    }
})

$BtnRefreshAll.Add_Click({ if ($Global:IsConnected) { Load-Dashboard } })
$BtnDashRefresh.Add_Click({ if ($Global:IsConnected) { Load-Dashboard } })
$BtnNewDeploy.Add_Click({ Switch-Panel "Wiz" })
$BtnScanLic.Add_Click({ $Script:LicScanned=$false; Invoke-LicenseScan })

# Wizard nav
$BtnWizNext.Add_Click({
    if (-not (Test-WizardStep $Script:WizStep)) { return }   # validation gate
    if ($Script:WizStep -lt $Script:WizTotalSteps) {
        $Script:WizStep++
        Update-WizardUI $Script:WizStep
    }
})
$BtnWizBack.Add_Click({
    if ($Script:WizStep -gt 1) {
        $Script:WizStep--
        Update-WizardUI $Script:WizStep
    }
})
$BtnWizCancel.Add_Click({
    $Script:WizStep = 1; Update-WizardUI 1; Switch-Panel "Dash"
})

# Wizard type radios
$RdoPooled.Add_Checked({    try{$script:TxtWizType.Text="Pooled Desktop: multiple users share VMs"}catch{} })
$RdoPersonal.Add_Checked({  try{$script:TxtWizType.Text="Personal Desktop: dedicated VM per user"}catch{} })
$RdoRemoteApp.Add_Checked({ try{$script:TxtWizType.Text="RemoteApp: publish individual applications"}catch{} })

# Sliders
$SliderMaxSess.Add_ValueChanged({ try{$script:TxtMaxSess.Text="$([int]$script:SliderMaxSess.Value)"}catch{} })
$SliderTokenHrs.Add_ValueChanged({ try{$script:TxtTokenHrs.Text="$([int]$script:SliderTokenHrs.Value) hours"}catch{} })
$SliderVMCnt.Add_ValueChanged({ try{$script:TxtVMCnt.Text="$([int]$script:SliderVMCnt.Value)"}catch{} })
$SliderProfGB.Add_ValueChanged({ try{$script:TxtProfGB.Text="$([int]$script:SliderProfGB.Value) GB"}catch{} })
$SliderCap.Add_ValueChanged({ try{$script:TxtCap.Text="$([int]$script:SliderCap.Value)%"}catch{} })
$SliderMinH.Add_ValueChanged({ try{$script:TxtMinH.Text="$([int]$script:SliderMinH.Value)"}catch{} })
$SliderFSLAlert.Add_ValueChanged({ try{$script:TxtFSLAlert.Text="$([int]$script:SliderFSLAlert.Value) GB"}catch{} })
$SetRefresh.Add_ValueChanged({ try{$script:TxtSetRefresh.Text="$([int]$script:SetRefresh.Value) seconds"}catch{} })

# VM size detail
$WizVMSize.Add_SelectionChanged({
    $idx = try{$script:WizVMSize.SelectedIndex}catch{-1}
    if ($idx -ge 0 -and $idx -lt $Script:VmSizes.Count) {
        $vm = $Script:VmSizes[$idx]
        try{$script:TxtVMDetail.Text="$($vm.vCPU) vCPU / $($vm.RAM) / Max sessions: $($vm.MaxSessions) / $($vm.UseCase)"}catch{}
    }
})

# Domain join toggle
$RdoEntra.Add_Checked({ try{$script:EntraPanel.Visibility=[System.Windows.Visibility]::Visible; $script:HybridPanel.Visibility=[System.Windows.Visibility]::Collapsed}catch{} })
$RdoHybrid.Add_Checked({ try{$script:HybridPanel.Visibility=[System.Windows.Visibility]::Visible; $script:EntraPanel.Visibility=[System.Windows.Visibility]::Collapsed}catch{} })

# Image source
$RdoMarket.Add_Checked({ try{$script:WizMktImg.IsEnabled=$true; $script:WizGallery.IsEnabled=$false; $script:WizGalleryImg.IsEnabled=$false}catch{} })
$RdoGallery.Add_Checked({ try{$script:WizMktImg.IsEnabled=$false; $script:WizGallery.IsEnabled=$true; $script:WizGalleryImg.IsEnabled=$true}catch{} })

# VNet load
$BtnLoadVNet.Add_Click({
    $rg = try{$script:WizVNetRG.SelectedItem}catch{$null}
    if (-not $rg) { return }
    try { $vnets = @(Get-AzVirtualNetwork -ResourceGroupName $rg -EA SilentlyContinue); $script:WizVNet.ItemsSource = @($vnets | ForEach-Object {$_.Name}) } catch {}
})
$WizVNetRG.Add_SelectionChanged({
    $rg = try{$script:WizVNetRG.SelectedItem}catch{$null}
    if (-not $rg) { return }
    try { $vnets = @(Get-AzVirtualNetwork -ResourceGroupName $rg -EA SilentlyContinue); $script:WizVNet.ItemsSource = @($vnets | ForEach-Object {$_.Name}) } catch {}
})
$WizVNet.Add_SelectionChanged({
    $rg = try{$script:WizVNetRG.SelectedItem}catch{$null}; $vn = try{$script:WizVNet.SelectedItem}catch{$null}
    if (-not $rg -or -not $vn) { return }
    try { $vnetObj = Get-AzVirtualNetwork -Name $vn -ResourceGroupName $rg -EA SilentlyContinue; $script:WizSubnet.ItemsSource = @($vnetObj.Subnets | ForEach-Object {$_.Name}) } catch {}
})

# Add tag
$BtnAddTag.Add_Click({
    $k = try{$script:WizTagK.Text.Trim()}catch{""}; $v = try{$script:WizTagV.Text.Trim()}catch{""}
    if (-not $k) { return }
    $items = try{$script:GridTags.ItemsSource}catch{$null}
    if (-not $items) { $items = [System.Collections.ObjectModel.ObservableCollection[PSObject]]::new() }
    $items.Add([PSCustomObject]@{TagKey=$k; TagVal=$v})
    try{$script:GridTags.ItemsSource=$items; $script:WizTagK.Text=""; $script:WizTagV.Text=""}catch{}
})

# Deploy
$BtnDeploy.Add_Click({ Start-AVDDeployment })
$BtnCancelDeploy.Add_Click({ $Global:Sync.CancelToken=$true; Write-Log "Cancellation requested..." "WARN"; try{$script:TxtDeployPhase.Text="Cancelling..."}catch{} })

# HP buttons
$BtnRefHP.Add_Click({ Load-HostPools })
$BtnNewHP.Add_Click({ Switch-Panel "Wiz" })

# Session host buttons
$BtnRefSess.Add_Click({ Load-SessionHosts })
$SessFilter.Add_SelectionChanged({ Load-SessionHosts })
$BtnDrainAll.Add_Click({
    if (-not $Global:IsConnected) { return }
    if (-not (Confirm-Action "Enable drain mode on ALL visible session hosts?`nNew sessions will be blocked on all hosts.")) { return }
    $rows = $script:GridSess.Items; $cnt = 0
    foreach ($row in $rows) {
        try { Update-AzWvdSessionHost -HostPoolName $row.HPName -ResourceGroupName $row.RG -Name $row.VMName -AllowNewSession $false -EA SilentlyContinue | Out-Null; $cnt++ } catch {}
    }
    Show-Toast "Drain enabled on $cnt hosts"
    Write-Log "Drain mode enabled on $cnt hosts" "OK"
    Load-SessionHosts
})
$BtnHealAll.Add_Click({
    if (-not $Global:IsConnected) { return }
    Write-Log "Checking for unhealthy hosts..." "STEP"
    $rows = $script:GridSess.Items
    $cnt = 0
    foreach ($row in $rows) {
        if ($row.Status -in @("Unavailable","NeedsAssistance","NoHeartBeat")) {
            try {
                Restart-AzVM -ResourceGroupName $row.RG -Name $row.VMName -EA SilentlyContinue | Out-Null
                Write-Log "  Restart triggered: $($row.VMName)" "OK"; $cnt++
            } catch { Write-Log "  Restart failed: $($row.VMName) - $_" "WARN" }
        }
    }
    Write-Log "Heal complete: $cnt VMs restarted" "OK"
})

# Context menus
$CtxSessEnableDrain.Add_Click({
    $row = $script:GridSess.SelectedItem
    if ($row -and $Global:IsConnected) {
        try { Update-AzWvdSessionHost -HostPoolName $row.HPName -ResourceGroupName $row.RG -Name $row.VMName -AllowNewSession $false | Out-Null; Write-Log "Drain ON: $($row.VMName)" "OK"; Load-SessionHosts } catch { Write-Log "Error: $_" "ERROR" }
    }
})
$CtxSessDisableDrain.Add_Click({
    $row = $script:GridSess.SelectedItem
    if ($row -and $Global:IsConnected) {
        try { Update-AzWvdSessionHost -HostPoolName $row.HPName -ResourceGroupName $row.RG -Name $row.VMName -AllowNewSession $true | Out-Null; Write-Log "Drain OFF: $($row.VMName)" "OK"; Load-SessionHosts } catch { Write-Log "Error: $_" "ERROR" }
    }
})
$CtxSessPortal.Add_Click({ $row=$script:GridSess.SelectedItem; if($row){Start-Process "https://portal.azure.com/#resource/subscriptions/$($Global:Subscription.Id)/resourceGroups/$($row.RG)/providers/Microsoft.DesktopVirtualization/hostPools/$($row.HPName)/sessionHosts"} })
$CtxHPPortal.Add_Click({ $row=$script:GridHP.SelectedItem; if($row){Start-Process "https://portal.azure.com/#resource/subscriptions/$($Global:Subscription.Id)/resourceGroups/$($row.RG)/providers/Microsoft.DesktopVirtualization/hostPools/$($row.Name)/overview"} })

# App groups
$BtnRefAG.Add_Click({ Load-AppGroups })
$BtnNewAG.Add_Click({ Switch-Panel "Wiz" })

# RDP property builder
$BtnBuildRDP.Add_Click({
    $clip = try{$script:RDPClip.IsChecked}catch{$true}
    $drv  = try{$script:RDPDrives.IsChecked}catch{$false}
    $prn  = try{$script:RDPPrint.IsChecked}catch{$true}
    $cam  = try{$script:RDPCam.IsChecked}catch{$true}
    $usb  = try{$script:RDPUSB.IsChecked}catch{$false}
    $sc   = try{$script:RDPSC.IsChecked}catch{$false}
    $wm   = try{$script:RDPWatermark.IsChecked}catch{$true}
    $scp  = try{$script:RDPScrCap.IsChecked}catch{$true}
    $wa   = try{$script:RDPWebAuthn.IsChecked}catch{$true}
    $mm   = try{$script:RDPMultiMon.IsChecked}catch{$true}
    $dr   = try{$script:RDPDynRes.IsChecked}catch{$true}
    $ss   = try{$script:RDPSmartSz.IsChecked}catch{$false}
    $props = @(
        "targetisaadjoined:i:1","enablerdsaadredirection:i:1"
        "redirectclipboard:i:$(if($clip){1}else{0})"
        "drivestoredirect:s:$(if($drv){'*'}else{''})"
        "redirectprinters:i:$(if($prn){1}else{0})"
        "camerastoredirect:s:$(if($cam){'*'}else{''})"
        "audiocapturemode:i:$(if($cam){1}else{0})","audiomode:i:0"
        "usbdevicestoredirect:s:$(if($usb){'*'}else{''})"
        "redirectsmartcards:i:$(if($sc){1}else{0})"
        "watermarking:i:$(if($wm){1}else{0})"
        "screen-capture-protection:i:$(if($scp){1}else{0})"
        "redirectwebauthn:i:$(if($wa){1}else{0})"
        "use multimon:i:$(if($mm){1}else{0})"
        "dynamic resolution:i:$(if($dr){1}else{0})"
        "smart sizing:i:$(if($ss){1}else{0})"
        "autoreconnection enabled:i:1","authentication level:i:2","enablecredsspsupport:i:1"
        "bandwidthautodetect:i:1","networkautodetect:i:1","compression:i:1"
    )
    try{$script:RDPProps.Text=($props -join ";")}catch{}
    Write-Log "RDP property string built" "OK"
})
$RDPPreset.Add_SelectionChanged({
    $sel = try{$script:RDPPreset.SelectedIndex}catch{0}
    $presets = @(
        "targetisaadjoined:i:1;enablerdsaadredirection:i:1;redirectclipboard:i:1;drivestoredirect:s:;camerastoredirect:s:*;audiocapturemode:i:1;audiomode:i:0;usbdevicestoredirect:s:;redirectprinters:i:1;watermarking:i:1;screen-capture-protection:i:1;redirectwebauthn:i:1;use multimon:i:1;dynamic resolution:i:1;autoreconnection enabled:i:1;authentication level:i:2;enablecredsspsupport:i:1;bandwidthautodetect:i:1;networkautodetect:i:1;compression:i:1"
        "targetisaadjoined:i:1;enablerdsaadredirection:i:1;redirectclipboard:i:0;drivestoredirect:s:;camerastoredirect:s:;audiocapturemode:i:0;usbdevicestoredirect:s:;redirectprinters:i:0;watermarking:i:1;screen-capture-protection:i:1;redirectwebauthn:i:0;use multimon:i:1;dynamic resolution:i:1;autoreconnection enabled:i:1;authentication level:i:2"
        "targetisaadjoined:i:1;enablerdsaadredirection:i:1;redirectclipboard:i:0;drivestoredirect:s:;camerastoredirect:s:;audiocapturemode:i:0;usbdevicestoredirect:s:;redirectprinters:i:0;watermarking:i:1;screen-capture-protection:i:1;use multimon:i:0;dynamic resolution:i:0;smart sizing:i:1;autoreconnection enabled:i:1"
        "targetisaadjoined:i:1;enablerdsaadredirection:i:1;redirectclipboard:i:1;drivestoredirect:s:*;camerastoredirect:s:*;audiocapturemode:i:1;audiomode:i:0;usbdevicestoredirect:s:*;redirectprinters:i:1;watermarking:i:0;screen-capture-protection:i:0;redirectwebauthn:i:1;use multimon:i:1;dynamic resolution:i:1;autoreconnection enabled:i:1;authentication level:i:0"
    )
    if ($sel -lt $presets.Count) { try{$script:RDPProps.Text=$presets[$sel]}catch{} }
})
$BtnApplyRDP.Add_Click({
    if (-not $Global:IsConnected) { return }
    $hpn    = try{$script:RDPPool.SelectedItem}catch{$null}
    $rdpStr = try{$script:RDPProps.Text -replace "`r`n",";" -replace "`n",";"}catch{""}
    if (-not $hpn) { [System.Windows.MessageBox]::Show("Select a host pool first.", "AVD Manager") | Out-Null; return }
    try {
        $hp = Get-AzWvdHostPool -EA SilentlyContinue | Where-Object {$_.Name -eq $hpn} | Select-Object -First 1
        if ($hp) {
            $rg = ($hp.Id -split "/")[4]
            Update-AzWvdHostPool -Name $hpn -ResourceGroupName $rg -CustomRdpProperty $rdpStr -EA Stop | Out-Null
            Write-Log "RDP properties applied to: $hpn" "OK"
            [System.Windows.MessageBox]::Show("RDP properties applied to '$hpn' successfully.", "AVD Manager") | Out-Null
        }
    } catch { Write-Log "RDP apply error: $_" "ERROR" }
})

# Log

# -- Missing button stubs (New RG, WS, KV, SA, LAW) -------------------------
$BtnNewRG.Add_Click({
    if (-not $Global:IsConnected) { [System.Windows.MessageBox]::Show("Connect to Azure first.", "AVD Manager") | Out-Null; return }
    $name = [Microsoft.VisualBasic.Interaction]::InputBox("Enter new Resource Group name:", "Create Resource Group", "rg-avd-")
    if (-not $name) { return }
    $loc = try { $script:WizRegion.SelectedValue } catch { "eastus" }
    if (-not $loc) { $loc = "eastus" }
    try {
        New-AzResourceGroup -Name $name -Location $loc -EA Stop | Out-Null
        Write-Log "Created resource group: $name" "OK"
        # Refresh the RG dropdowns
        $rgs = @(Get-AzResourceGroup -EA SilentlyContinue | Sort-Object ResourceGroupName)
        $rgList = @($rgs | ForEach-Object { $_.ResourceGroupName })
        try { $script:WizRG.ItemsSource     = $rgList; $script:WizRG.SelectedValue     = $name } catch {}
        try { $script:WizVNetRG.ItemsSource = $rgList; $script:WizVNetRG.SelectedValue = $name } catch {}
        [System.Windows.MessageBox]::Show("Resource group '$name' created in $loc.", "AVD Manager") | Out-Null
    } catch {
        Write-Log "Create RG error: $_" "ERROR"
        [System.Windows.MessageBox]::Show("Failed to create resource group:`n$_", "AVD Manager") | Out-Null
    }
})

$BtnNewWS.Add_Click({
    if (-not $Global:IsConnected) { [System.Windows.MessageBox]::Show("Connect to Azure first.", "AVD Manager") | Out-Null; return }
    $name = [Microsoft.VisualBasic.Interaction]::InputBox("Enter new Workspace name:", "Create Workspace", "ws-avd-")
    if (-not $name) { return }
    $rg  = try { $script:WizRG.SelectedItem } catch { "" }
    $loc = try { $script:WizRegion.SelectedValue } catch { "eastus" }
    if (-not $rg) { [System.Windows.MessageBox]::Show("Select a Resource Group first.", "AVD Manager") | Out-Null; return }
    try {
        New-AzWvdWorkspace -Name $name -ResourceGroupName $rg -Location $loc -FriendlyName $name -EA Stop | Out-Null
        Write-Log "Created workspace: $name" "OK"
        Load-WizardWorkspaces
        try { $script:WizWS.SelectedValue = $name } catch {}
        [System.Windows.MessageBox]::Show("Workspace '$name' created.", "AVD Manager") | Out-Null
    } catch {
        Write-Log "Create Workspace error: $_" "ERROR"
        [System.Windows.MessageBox]::Show("Failed to create workspace:`n$_", "AVD Manager") | Out-Null
    }
})

$BtnNewKV.Add_Click({
    if (-not $Global:IsConnected) { [System.Windows.MessageBox]::Show("Connect to Azure first.", "AVD Manager") | Out-Null; return }
    [System.Windows.MessageBox]::Show("To create a Key Vault, go to:`nAzure Portal > Key Vaults > Create`n`nKey Vault name must be globally unique (3-24 chars, alphanumeric and hyphens).", "Create Key Vault") | Out-Null
    Start-Process "https://portal.azure.com/#create/Microsoft.KeyVault"
})

$BtnNewSA.Add_Click({
    if (-not $Global:IsConnected) { [System.Windows.MessageBox]::Show("Connect to Azure first.", "AVD Manager") | Out-Null; return }
    [System.Windows.MessageBox]::Show("To create a Storage Account for FSLogix:`n`n1. Azure Portal > Storage Accounts > Create`n2. Choose Premium performance, FileStorage kind for Azure Files Premium`n3. After creation, create a file share named 'profiles'`n4. Enable 'Azure Active Directory' (Entra ID) authentication on the share", "Create Storage Account") | Out-Null
    Start-Process "https://portal.azure.com/#create/Microsoft.StorageAccount-ARM"
})

$BtnNewLAW.Add_Click({
    if (-not $Global:IsConnected) { [System.Windows.MessageBox]::Show("Connect to Azure first.", "AVD Manager") | Out-Null; return }
    $name = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Log Analytics Workspace name:", "Create Log Analytics Workspace", "law-avd-")
    if (-not $name) { return }
    $rg  = try { $script:WizRG.SelectedItem } catch { "" }
    $loc = try { $script:WizRegion.SelectedValue } catch { "eastus" }
    if (-not $rg) { [System.Windows.MessageBox]::Show("Select a Resource Group first.", "AVD Manager") | Out-Null; return }
    try {
        New-AzOperationalInsightsWorkspace -Name $name -ResourceGroupName $rg -Location $loc -Sku PerGB2018 -EA Stop | Out-Null
        Write-Log "Created Log Analytics Workspace: $name" "OK"
        Load-WizardLAWs
        try { $script:WizLAW.SelectedValue = $name } catch {}
        [System.Windows.MessageBox]::Show("Log Analytics Workspace '$name' created.", "AVD Manager") | Out-Null
    } catch {
        Write-Log "Create LAW error: $_" "ERROR"
        [System.Windows.MessageBox]::Show("Failed to create workspace:`n$_", "AVD Manager") | Out-Null
    }
})

# -- App group publish --------------------------------------------------------
$BtnPubApp.Add_Click({
    if (-not $Global:IsConnected) { return }
    $ag = $script:GridAG.SelectedItem
    if (-not $ag) { [System.Windows.MessageBox]::Show("Select a RemoteApp application group first.", "AVD Manager") | Out-Null; return }
    if ($ag.Type -ne "RemoteApp") { [System.Windows.MessageBox]::Show("Select a RemoteApp-type app group to publish applications.", "AVD Manager") | Out-Null; return }
    $appName = [Microsoft.VisualBasic.Interaction]::InputBox("Application display name:", "Publish Application", "My App")
    if (-not $appName) { return }
    $appPath = [Microsoft.VisualBasic.Interaction]::InputBox("Executable path (e.g. C:\Windows\System32
otepad.exe):", "Publish Application", "C:\Windows\System32
otepad.exe")
    if (-not $appPath) { return }
    try {
        $hp = Get-AzWvdApplicationGroup -Name $ag.Name -EA SilentlyContinue | Select-Object -First 1
        $rg = ($hp.Id -split "/")[4]
        New-AzWvdApplication -Name ($appName -replace "\s","_") -ApplicationGroupName $ag.Name `
            -ResourceGroupName $rg -FilePath $appPath -IconPath $appPath -IconIndex 0 `
            -CommandLineSetting DoNotAllow -ApplicationType InBuilt -EA Stop | Out-Null
        Write-Log "Published application '$appName' to $($ag.Name)" "OK"
        Load-AppGroups
    } catch {
        Write-Log "Publish app error: $_" "ERROR"
        [System.Windows.MessageBox]::Show("Failed to publish application:`n$_", "AVD Manager") | Out-Null
    }
})

# -- FSLogix actions ---------------------------------------------------------
$BtnFSLRef.Add_Click({
    if (-not $Global:IsConnected) { return }
    Write-Log "FSLogix refresh: scan Azure Files shares for profile containers..." "STEP"
    try {
        $sas = @(Get-AzStorageAccount -EA SilentlyContinue)
        $rows = [System.Collections.Generic.List[PSObject]]::new()
        foreach ($sa in $sas) {
            $rg  = ($sa.Id -split "/")[4]
            $ctx = $sa.Context
            $shares = @(Get-AzStorageShare -Context $ctx -EA SilentlyContinue | Where-Object { $_.Name -match "profile" })
            foreach ($sh in $shares) {
                $rows.Add([PSCustomObject]@{
                    Username="(scan in session)"; Share=$sh.Name; SizeMB=$sh.ShareUsageBytes/1MB
                    Health="OK"; HBg="#0A2010"; HFg="#10B981"; LastMount="--"
                })
            }
        }
        if ($rows.Count -eq 0) { Write-Log "No profile shares found. Check storage account permissions." "WARN" }
        else {
            try { $script:GridFSL.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[PSObject]]($rows.ToArray()) } catch {}
            try { $script:FSLTotal.Text = "$($rows.Count)" } catch {}
            Write-Log "FSLogix: found $($rows.Count) profile share(s)" "OK"
        }
    } catch { Write-Log "FSLogix refresh error: $_" "WARN" }
})

$BtnFSLLocks.Add_Click({
    if (-not $Global:IsConnected) { return }
    $msg = "To clear FSLogix .lock files:`n`n1. Ensure the affected user is logged off`n2. Navigate to the profile storage share`n3. Delete any *.lock files in the user's VHDX folder`n`nThis tool cannot directly delete files on the share - use the Azure Storage Explorer or connect to the share via File Explorer."
    [System.Windows.MessageBox]::Show($msg, "Clear FSLogix Lock Files") | Out-Null
})

$BtnFSLTmp.Add_Click({
    [System.Windows.MessageBox]::Show("To remove temporary FSLogix VHDX files:`n`n1. Ensure no users are logged in`n2. Open Azure Storage Explorer`n3. Navigate to your profiles share`n4. Delete any *.vhdx.tmp or *.vhd.tmp files`n`nThese are created when a session terminates uncleanly.", "Remove Temp VHDXs") | Out-Null
})

$BtnFSLDiag.Add_Click({
    if (-not $Global:IsConnected) { return }
    Write-Log "Running FSLogix diagnostics..." "STEP"
    $results = [System.Collections.Generic.List[string]]::new()
    try {
        # Check storage accounts with file shares named 'profiles'
        $sas = @(Get-AzStorageAccount -EA SilentlyContinue)
        $results.Add("Storage accounts checked: $($sas.Count)")
        foreach ($sa in $sas) {
            $ctx = $sa.Context
            $shares = @(Get-AzStorageShare -Context $ctx -EA SilentlyContinue | Where-Object {$_.Name -match "profile"})
            if ($shares.Count -gt 0) { $results.Add("[OK] Found profile share in: $($sa.StorageAccountName)") }
        }
        # Check if FSLogix registry keys exist on a session host (requires PS remoting)
        $results.Add("")
        $results.Add("On each session host, verify:")
        $results.Add("  HKLM:\SOFTWARE\FSLogix\Profiles\Enabled = 1")
        $results.Add("  HKLM:\SOFTWARE\FSLogix\Profiles\VHDLocations = \\server\share")
        $results.Add("")
        $results.Add("FSLogix version check: Run 'frx version' on session hosts")
        $results.Add("Minimum recommended: 2.9.8440.42104")
    } catch { $results.Add("Error: $_") }
    [System.Windows.MessageBox]::Show($results -join "`n", "FSLogix Diagnostics") | Out-Null
    Write-Log "FSLogix diagnostic check complete" "OK"
})

# -- Scaling -----------------------------------------------------------------
$BtnRefScale.Add_Click({ if ($Global:IsConnected) { Load-ScalingPlans } })

$BtnNewScale.Add_Click({
    if (-not $Global:IsConnected) { [System.Windows.MessageBox]::Show("Connect to Azure first.", "AVD Manager") | Out-Null; return }
    $hp = $script:GridHP.SelectedItem
    if (-not $hp) {
        # Try getting from host pool grid
        $hps = @(Get-AzWvdHostPool -EA SilentlyContinue)
        if ($hps.Count -eq 0) { [System.Windows.MessageBox]::Show("No host pools found. Deploy a host pool first.", "AVD Manager") | Out-Null; return }
    }
    [System.Windows.MessageBox]::Show("To create a Scaling Plan:`n`nAzure Portal > Azure Virtual Desktop > Scaling Plans > Create`n`nOr deploy via wizard (Step 7: RDP and Scaling tab).", "New Scaling Plan") | Out-Null
    Start-Process "https://portal.azure.com/#blade/Microsoft_Azure_WVD/WvdManagerMenuBlade/scalingPlans"
})

# -- Monitoring ---------------------------------------------------------------
$BtnAddAlert.Add_Click({
    if (-not $Global:IsConnected) { return }
    [System.Windows.MessageBox]::Show("Common AVD alerts to configure:`n`n[!] Session host unavailable > 0`n[!] User connection failures > 5/min`n[!] CPU utilization > 85% sustained`n[!] Available session hosts = 0`n[!] FSLogix mount failures > 0`n`nCreate via: Azure Monitor > Alerts > Create alert rule`nOr use Azure Portal > Monitor > Alerts.", "Add Alert Rule") | Out-Null
    Start-Process "https://portal.azure.com/#blade/Microsoft_Azure_Monitoring/AzureMonitoringBrowseBlade/alertsV2"
})

# -- Subscription switcher ---------------------------------------------------
$CboSubSwitcher.Add_SelectionChanged({
    $idx = try{$script:CboSubSwitcher.SelectedIndex}catch{-1}
    if ($idx -lt 0 -or -not $Global:SubList -or $idx -ge $Global:SubList.Count) { return }
    $sub = $Global:SubList[$idx]
    if ($sub.Id -eq (try{(Get-AzContext).Subscription.Id}catch{""})) { return }
    try {
        Set-AzContext -SubscriptionId $sub.Id -EA Stop | Out-Null
        $Global:Subscription = (Get-AzContext).Subscription
        Set-Connected $true
        Write-Log "Switched to subscription: $($sub.Name)" "OK"
        Show-Toast "Switched to: $($sub.Name)"
        Load-Dashboard
        $Script:LicScanned = $false
    } catch { Write-Log "Subscription switch error: $_" "ERROR" }
})

# -- HP filter / CSV ---------------------------------------------------------
$FilterHP.Add_TextChanged({
    $filter = try{$script:FilterHP.Text.ToLower()}catch{""}
    if (-not $filter) {
        Load-HostPools; return
    }
    $all = $Global:Sync["CachedHPRows"]
    if (-not $all) { return }
    $filtered = @($all | Where-Object {
        $_.Name    -match $filter -or
        $_.Type    -match $filter -or
        $_.Location -match $filter -or
        $_.Status  -match $filter -or
        $_.RG      -match $filter
    })
    try { $script:GridHP.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[PSObject]]($filtered) } catch {}
})
$BtnExportHP.Add_Click({ Export-GridData $script:GridHP "host-pools" })

# -- Session filter / CSV / user sessions ------------------------------------
$FilterSess.Add_TextChanged({
    $filter = try{$script:FilterSess.Text.ToLower()}catch{""}
    if (-not $filter) { Load-SessionHosts; return }
    $all = $Global:Sync["CachedSessRows"]
    if (-not $all) { return }
    $filtered = @($all | Where-Object {
        $_.VMName  -match $filter -or
        $_.Pool    -match $filter -or
        $_.Status  -match $filter -or
        $_.VMSize  -match $filter
    })
    try { $script:GridSess.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[PSObject]]($filtered) } catch {}
})
$BtnExportSess.Add_Click({ Export-GridData $script:GridSess "session-hosts" })

# -- User session management -------------------------------------------------
$BtnRefUserSess.Add_Click({ Load-UserSessions })
$BtnMsgUser.Add_Click({
    $row = $script:GridUserSess.SelectedItem
    if (-not $row) { Show-Toast "Select a user session first" "#F59E0B"; return }
    Send-UserMessage $row
})
$BtnDisconnectUser.Add_Click({
    $row = $script:GridUserSess.SelectedItem
    if (-not $row) { Show-Toast "Select a user session first" "#F59E0B"; return }
    Invoke-Disconnect $row
})
$BtnLogoffUser.Add_Click({
    $row = $script:GridUserSess.SelectedItem
    if (-not $row) { Show-Toast "Select a user session first" "#F59E0B"; return }
    Invoke-ForceLogoff $row
})
$BtnExportUserSess.Add_Click({ Export-GridData $script:GridUserSess "user-sessions" })

# -- App group: assign users -------------------------------------------------
$BtnAssignUsers.Add_Click({
    if (-not $Global:IsConnected) { return }
    $ag = $script:GridAG.SelectedItem
    if (-not $ag) { Show-Toast "Select an app group first" "#F59E0B"; return }
    $upns = [Microsoft.VisualBasic.Interaction]::InputBox(
        "Enter UPNs to assign to '$($ag.Name)'`n(comma or newline separated):",
        "Assign Users to App Group", "")
    if (-not $upns) { return }
    $list = @($upns -split "[,\n]" | ForEach-Object {$_.Trim()} | Where-Object {$_})
    try {
        $agObj = Get-AzWvdApplicationGroup -Name $ag.Name -EA SilentlyContinue | Select-Object -First 1
        if (-not $agObj) { $agObj = Get-AzWvdApplicationGroup -EA SilentlyContinue | Where-Object {$_.Name -eq $ag.Name} | Select-Object -First 1 }
        if ($agObj) {
            Set-AppGroupUsers -AppGroupId $agObj.Id -UserUPNs $list
            Show-Toast "Assigned $($list.Count) user(s) to $($ag.Name)"
        } else { Show-Toast "App group not found" "#EF4444" }
    } catch { Write-Log "Assign error: $_" "ERROR" }
})

# -- Destructive confirmations -----------------------------------------------
$CtxSessRemove.Add_Click({
    $row = $script:GridSess.SelectedItem
    if (-not ($row -and $Global:IsConnected)) { return }
    if (-not (Confirm-Action "Remove '$($row.VMName)' from pool '$($row.Pool)'?`nThe VM itself is NOT deleted.")) { return }
    try {
        Remove-AzWvdSessionHost -HostPoolName $row.HPName -ResourceGroupName $row.RG -Name $row.VMName -Force -EA Stop | Out-Null
        Write-Audit "REMOVE_SESSION_HOST" "HP/$($row.HPName)/Host/$($row.VMName)" "" "OK"
        Show-Toast "Removed $($row.VMName)"; Write-Log "Removed: $($row.VMName)" "OK"; Load-SessionHosts
    } catch { Write-Log "Remove error: $_" "ERROR" }
})
$CtxHPDelete.Add_Click({
    $row = $script:GridHP.SelectedItem
    if (-not ($row -and $Global:IsConnected)) { return }
    if (-not (Confirm-Action "DELETE host pool '$($row.Name)'?`n`nAll session host registrations will be removed.`nVMs themselves are NOT deleted.")) { return }
    $confirm = [Microsoft.VisualBasic.Interaction]::InputBox("Type the pool name to confirm:", "Confirm Delete", "")
    if ($confirm -ne $row.Name) { Show-Toast "Name mismatch — cancelled" "#F59E0B"; return }
    try {
        Remove-AzWvdHostPool -Name $row.Name -ResourceGroupName $row.RG -Force -EA Stop | Out-Null
        Show-Toast "Deleted: $($row.Name)"; Write-Log "Deleted host pool: $($row.Name)" "OK"; Load-HostPools
    } catch { Write-Log "Delete error: $_" "ERROR" }
})

# -- AVD Insights on Monitoring panel ----------------------------------------
$BtnAddAlert.Add_Click({
    if (-not $Global:IsConnected) { return }
    # Try to get a LAW workspace ID to run insights query
    $sel = try{$script:GridLAW.SelectedItem}catch{$null}
    if ($sel) {
        Write-Log "Running AVD Insights queries against workspace: $($sel.Name)..." "STEP"
        $law = Get-AzOperationalInsightsWorkspace -Name $sel.Name -EA SilentlyContinue | Select-Object -First 1
        if (-not $law) { $law = Get-AzOperationalInsightsWorkspace -EA SilentlyContinue | Where-Object {$_.Name -eq $sel.Name} | Select-Object -First 1 }
        if ($law) {
            $insights = Get-AVDInsights -WorkspaceId $law.CustomerId
            if ($insights) {
                $conn = try{$insights.Connections[0][0]}catch{"N/A"}
                $errs = try{$insights.Errors[0][0]}catch{"N/A"}
                $msg = "AVD Insights (last 24h):`n`nConnections: $conn`nErrors: $errs`n`nTop users (7d):`n"
                try { $insights.TopUsers | ForEach-Object { $msg += "  $($_[0]) — $($_[1]) sessions`n" } } catch {}
                [System.Windows.MessageBox]::Show($msg, "AVD Insights") | Out-Null
                Write-Log "Insights: $conn connections, $errs errors" "OK"
                return
            }
        }
    }
    # Fallback: open Portal alert creation
    [System.Windows.MessageBox]::Show("Select a Log Analytics workspace in the grid above to run AVD Insights queries.`n`nOr click OK to open Azure Monitor alert creation in the Portal.", "Add Alert / AVD Insights") | Out-Null
    Start-Process "https://portal.azure.com/#blade/Microsoft_Azure_Monitoring/AzureMonitoringBrowseBlade/alertsV2"
})

# Load user sessions when switching to Sess panel (already handled in Switch-Panel)
# Also auto-load when HP filter changes
$SessFilter.Add_SelectionChanged({ Load-SessionHosts; Load-UserSessions })

$BtnClearLog.Add_Click({ try{$script:MainLogBox.Document.Blocks.Clear()}catch{} })

# Settings
$BtnRunBP.Add_Click({ Invoke-BestPracticesCheck })
$BtnExportBP.Add_Click({ Export-GridData $script:GridBP "avd-best-practices" })
$BPCatFilter.Add_SelectionChanged({
    $sel = try{$script:BPCatFilter.SelectedItem.Content}catch{"All Categories"}
    if ($sel -eq "All Categories" -or -not $Script:BPRows -or $Script:BPRows.Count -eq 0) {
        try { $script:GridBP.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[PSObject]]($Script:BPRows.ToArray()) } catch {}
        return
    }
    $filtered = @($Script:BPRows | Where-Object { $_.Category -eq $sel })
    try { $script:GridBP.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[PSObject]]($filtered) } catch {}
})

$BtnSaveSet.Add_Click({
    try {
        $Global:Cfg.LastSubscription = try{$script:SetSubId.Text}catch{""}
        $Global:Cfg.LastTenant       = try{$script:SetTenantId.Text}catch{""}
        $Global:Cfg.LastLocation     = try{$script:SetLocation.SelectedValue}catch{"eastus"}
        $Global:Cfg.RefreshSecs      = try{[int]$script:SetRefresh.Value}catch{60}
        $Global:RefreshSecs          = $Global:Cfg.RefreshSecs
        $Script:Countdown            = $Global:RefreshSecs
        $Global:Cfg | ConvertTo-Json -Depth 5 | Set-Content $Global:ConfigFile -Encoding UTF8
        Write-Log "Settings saved to $Global:ConfigFile" "OK"
        Write-Audit "SETTINGS_SAVED" "AVD-Manager-Config" "" "OK"
        Show-Toast "Settings saved"
    } catch { Write-Log "Error saving settings: $_" "ERROR" }
})

$BtnTestConn.Add_Click({
    try {
        $ctx  = Get-AzContext -EA Stop
        $cost = Get-AzRealCost -Scope "/subscriptions/$($ctx.Subscription.Id)" -Days 30
        $costTxt = if ($cost -ne $null) {"`$${cost} (last 30d VM+AVD)"} else {"(unavailable)"}
        [System.Windows.MessageBox]::Show(
            "Connected`n`n" +
            "Account      : $($ctx.Account.Id)`n" +
            "Subscription : $($ctx.Subscription.Name)`n" +
            "Tenant       : $($ctx.Tenant.Id)`n" +
            "Token Expiry : $(if($Global:GraphTokenExpiry -gt [DateTime]::MinValue){$Global:GraphTokenExpiry.ToLocalTime().ToString('HH:mm:ss')}else{'(not cached)'})`n`n" +
            "Azure Cost   : $costTxt",
            "Connection Status") | Out-Null
    } catch {
        [System.Windows.MessageBox]::Show("Not connected. Click 'Connect to Azure'.", "AVD Manager") | Out-Null
    }
})

# Config export / import (add buttons if they don't exist in XAML; handled via context)
# Keyboard shortcut: Ctrl+E = Export, Ctrl+I = Import
$Window.Add_KeyDown({
    param($s,$e)
    if ($e.Key -eq "E" -and [System.Windows.Input.Keyboard]::Modifiers -eq "Control") { Export-ToolConfig }
    if ($e.Key -eq "I" -and [System.Windows.Input.Keyboard]::Modifiers -eq "Control") { Import-ToolConfig }
    if ($e.Key -eq "F5")                                                               { if ($Global:IsConnected) { Load-Dashboard } }
    if ($e.Key -eq "F1") {
        [System.Windows.MessageBox]::Show(
            "AVD Manager Keyboard Shortcuts`n`n" +
            "F5          : Refresh dashboard`n" +
            "Ctrl+E      : Export configuration`n" +
            "Ctrl+I      : Import configuration`n" +
            "Ctrl+W      : Open in Azure Portal (when grid row selected)`n" +
            "Escape      : Cancel current deployment`n`n" +
            "Best Practices scoring:`n" +
            "  PASS = 100pts  WARN = 50pts  FAIL = 0pts`n" +
            "  Score = (PASS*100 + WARN*50) / total_checks",
            "Help - AVD Manager") | Out-Null
    }
    if ($e.Key -eq "Escape" -and $Global:Sync.IsDeploying) {
        $Global:Sync.CancelToken = $true
        Write-Log "Cancellation requested via Escape key" "WARN"
    }
})

# ============================================================================
# INITIALIZATION
# ============================================================================
try { $script:SetLocation.ItemsSource = $Script:AzureRegions; $script:SetLocation.SelectedValue = $Global:Cfg.LastLocation } catch {}
try { $script:SetSubId.Text      = $Global:Cfg.LastSubscription } catch {}
try { $script:SetTenantId.Text   = $Global:Cfg.LastTenant } catch {}
try { $script:SetRefresh.Value   = $Global:Cfg.RefreshSecs } catch {}
try { $script:WizRegion.ItemsSource = $Script:AzureRegions } catch {}
try { $script:WizVMSize.ItemsSource = $Script:VmSizes | ForEach-Object {"$($_.Size) ($($_.vCPU) vCPU, $($_.RAM)) - $($_.UseCase)"}; $script:WizVMSize.SelectedIndex=1 } catch {}
try { $script:WizMktImg.ItemsSource = $Script:MktImages; $script:WizMktImg.SelectedIndex=0 } catch {}
try { $script:GridTags.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[PSObject]]::new() } catch {}

Update-WizardUI 1

# Run prerequisites check at startup
Test-Prerequisites

# Restore window position/size from last session
Restore-WindowState

Write-Log "AVD Manager v$Global:AppVersion initialized" "OK"
Write-Log "Click 'Connect to Azure' to begin" "INFO"
Write-Log "Supports: Pooled, Personal, RemoteApp | License assessment | Best Practices | All Azure regions" "INFO"
Write-Log "Audit log: $Global:AuditLog" "INFO"

$Script:Timer.Start()

# Save window state on close, write audit entry
$Window.Add_Closing({
    try { $Script:Timer.Stop() } catch {}
    $Global:Sync.CancelToken = $true
    Save-WindowState
    try { $Global:Cfg | ConvertTo-Json -Depth 5 | Set-Content $Global:ConfigFile -Encoding UTF8 } catch {}
    try { Write-Audit "SESSION_END" "AVD-Manager" "Normal exit" "OK" } catch {}
    try { Write-Log "AVD Manager closed" "INFO" } catch {}
})

[System.Console]::WriteLine("Starting Azure Virtual Desktop Manager v$Global:AppVersion...")
[System.Console]::WriteLine("Log: $Global:LogFile")
[System.Console]::WriteLine("Audit: $Global:AuditLog")
[void]$Window.ShowDialog()

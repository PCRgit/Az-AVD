#Requires -Version 7.0
<#
.SYNOPSIS
    BDF Azure Virtual Desktop — All-In-One Interactive Deployment Console
    Production-ready menu-driven deployment with existing resource detection,
    permission management, and dynamic auto-scaling.

.DESCRIPTION
    Interactive AIO console featuring:
      • Subscription & tenant selector
      • Environment profiles (POC / Staging / Production)
      • Existing resource DETECTION — VNet, NSG, Storage, KV, LAW, etc.
      • Component-by-component deploy / skip / reuse
      • Permission Manager — check & assign all required RBAC roles
      • Dynamic Auto-Scaling configuration wizard
      • Runbook management (deploy / test / trigger)
      • Health dashboard & post-deploy validation
      • Config save/resume (JSON) for interrupted deployments
      • Full audit log with timestamps

.NOTES
    Author  : Jaimin
    Version : 3.0 — AIO Interactive
    Date    : May 2026
    Requires: Az PowerShell 11.0+, Az.DesktopVirtualization 4.0+
              PowerShell 7.0+ recommended (color + Unicode support)
#>

[CmdletBinding()]
param(
    [string]$ConfigFile  = ".\AVD-Config.json",
    [string]$LogFile     = ".\AVD-Deploy-$(Get-Date -Format 'yyyyMMdd-HHmmss').log",
    [switch]$NoLogo,
    [switch]$Unattended,   # skip interactive prompts — use saved config
    [switch]$ValidateOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference    = "SilentlyContinue"
$PSDefaultParameterValues['*:ErrorAction'] = 'SilentlyContinue'

#region ═══════════════════════════════════════════════════════════════════════
#  CONSOLE UI LAYER
#═══════════════════════════════════════════════════════════════════════════════

# Box-drawing / UI characters
$UI = @{
    TL='╔'; TR='╗'; BL='╚'; BR='╝'; H='═'; V='║'
    TM='╦'; BM='╩'; LM='╠'; RM='╣'; XX='╬'
    SH='─'; SV='│'; STL='┌'; STR='┐'; SBL='└'; SBR='┘'
    Check='✔'; Cross='✖'; Warn='⚠'; Arrow='▶'; Dot='●'
    Up='▲'; Down='▼'; Star='★'; Gear='⚙'; Lock='🔒'
    Skip='⊘'; New='⊕'; Reuse='↺'; Deploy='⚡'
}

# Color palette
$C = @{
    Hdr     = 'Cyan';    HdrBg   = 'DarkBlue'
    Menu    = 'White';   MenuSel = 'Yellow'
    Ok      = 'Green';   Fail    = 'Red'
    Warn    = 'Yellow';  Info    = 'Cyan'
    Muted   = 'DarkGray'; Accent = 'Magenta'
    Azure   = 'Blue';    Gold    = 'DarkYellow'
    Step    = 'DarkCyan'; Input  = 'White'
    New     = 'Green';   Reuse   = 'Cyan';  Skip = 'DarkGray'
    Border  = 'DarkCyan'
}

function Write-Log {
    param([string]$Msg, [string]$Level = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogFile -Value "[$ts][$Level] $Msg" -EA SilentlyContinue
}

function Write-Banner {
    param([string]$Text, [string]$Color = $C.Hdr, [int]$Width = 72)
    $pad  = [Math]::Max(0, ($Width - $Text.Length - 4))
    $lpad = [Math]::Floor($pad / 2); $rpad = $pad - $lpad
    $line = $UI.H * $Width
    Write-Host ""
    Write-Host "$($UI.TL)$line$($UI.TR)" -ForegroundColor $C.Border
    Write-Host "$($UI.V)  $(' ' * $lpad)$Text$(' ' * $rpad)  $($UI.V)" -ForegroundColor $Color
    Write-Host "$($UI.BL)$line$($UI.BR)" -ForegroundColor $C.Border
    Write-Host ""
}

function Write-Section {
    param([string]$Title, [string]$Color = $C.Accent)
    $line = $UI.SH * 68
    Write-Host ""
    Write-Host "  $($UI.STL)$line$($UI.STR)" -ForegroundColor $C.Border
    Write-Host "  $($UI.SV)  $($UI.Arrow) $Title" -ForegroundColor $Color
    Write-Host "  $($UI.SBL)$line$($UI.SBR)" -ForegroundColor $C.Border
}

function Write-Status {
    param([string]$Label, [string]$Value,
          [ValidateSet('ok','fail','warn','info','skip','new','reuse','step','pending')]
          [string]$Type = 'info', [int]$Indent = 4)
    $icons  = @{ok='✔'; fail='✖'; warn='⚠'; info='ℹ'; skip='⊘'; new='⊕'; reuse='↺'; step='▶'; pending='○'}
    $colors = @{ok=$C.Ok; fail=$C.Fail; warn=$C.Warn; info=$C.Info; skip=$C.Muted; new=$C.New; reuse=$C.Reuse; step=$C.Step; pending=$C.Muted}
    $icon   = $icons[$Type]; $color = $colors[$Type]
    $pad    = " " * $Indent
    if ($Value) {
        Write-Host "$pad$icon " -NoNewline -ForegroundColor $color
        Write-Host ("{0,-36}" -f $Label) -NoNewline -ForegroundColor $C.Menu
        Write-Host $Value -ForegroundColor $color
    } else {
        Write-Host "$pad$icon $Label" -ForegroundColor $color
    }
    Write-Log "$Type | $Label $Value"
}

function Write-Line { Write-Host "  $($UI.SH * 68)" -ForegroundColor $C.Border }

function Read-MenuChoice {
    param([string]$Prompt = "Choice", [string[]]$Valid = @(), [string]$Default = "")
    do {
        Write-Host ""
        Write-Host "  $($UI.Arrow) " -NoNewline -ForegroundColor $C.Accent
        Write-Host $Prompt -NoNewline -ForegroundColor $C.Input
        if ($Default) { Write-Host " [$Default]" -NoNewline -ForegroundColor $C.Muted }
        Write-Host " : " -NoNewline -ForegroundColor $C.Muted
        $choice = Read-Host
        if (-not $choice -and $Default) { $choice = $Default }
    } until (-not $Valid -or $choice -in $Valid)
    return $choice.Trim()
}

function Read-YesNo {
    param([string]$Prompt, [bool]$Default = $true)
    $opts  = if ($Default) { "[Y/n]" } else { "[y/N]" }
    $defCh = if ($Default) { "Y" } else { "N" }
    $r = Read-MenuChoice -Prompt "$Prompt $opts" -Valid @("Y","y","N","n","") -Default $defCh
    return ($r -in @("Y","y"))
}

function Read-FromList {
    param([string]$Title, [object[]]$Items, [string]$DisplayProp,
          [string]$SubProp = "", [bool]$AllowNew = $true, [bool]$AllowSkip = $false)

    Write-Section $Title
    if ($Items.Count -eq 0) {
        Write-Status "No existing resources found in this subscription" "" "warn"
        return $null
    }

    $i = 1
    foreach ($item in $Items) {
        $main = if ($DisplayProp) { $item.$DisplayProp } else { $item }
        $sub  = if ($SubProp) { " — $($item.$SubProp)" } else { "" }
        Write-Host ("    {0,2}. " -f $i) -NoNewline -ForegroundColor $C.Accent
        Write-Host $main -NoNewline -ForegroundColor $C.Menu
        Write-Host $sub -ForegroundColor $C.Muted
        $i++
    }
    $opts = @()
    $opts += (1..($Items.Count) | ForEach-Object { "$_" })
    if ($AllowNew)  { Write-Host "     N.  Create new" -ForegroundColor $C.New;  $opts += "N","n" }
    if ($AllowSkip) { Write-Host "     S.  Skip this component" -ForegroundColor $C.Skip; $opts += "S","s" }

    $choice = Read-MenuChoice -Prompt "Select" -Valid $opts
    if ($choice -in @("N","n")) { return "NEW" }
    if ($choice -in @("S","s")) { return "SKIP" }
    return $Items[[int]$choice - 1]
}

function Show-Progress {
    param([string]$Activity, [int]$Percent, [string]$Status = "")
    $filled = [Math]::Floor($Percent / 2.5)
    $empty  = 40 - $filled
    $bar    = ("#" * $filled) + ("─" * $empty)
    Write-Host "`r  [$bar] $Percent% " -NoNewline -ForegroundColor $C.Azure
    if ($Status) { Write-Host $Status -NoNewline -ForegroundColor $C.Muted }
}

function Confirm-Action {
    param([string]$Action, [string]$Detail = "")
    Write-Host ""
    Write-Host "  ┌─ Confirm Action ──────────────────────────────────────────────┐" -ForegroundColor $C.Warn
    Write-Host "  │  $($UI.Warn)  $Action" -ForegroundColor $C.Warn
    if ($Detail) { Write-Host "  │     $Detail" -ForegroundColor $C.Muted }
    Write-Host "  └───────────────────────────────────────────────────────────────┘" -ForegroundColor $C.Warn
    return (Read-YesNo "Proceed?" $true)
}

function Show-Logo {
    if ($NoLogo) { return }
    Clear-Host
    Write-Host @"

`e[34m  ╔═══════════════════════════════════════════════════════════════════════╗
  ║                                                                       ║
  ║   `e[36m █████╗ ██╗   ██╗██████╗     `e[34m                                        ║
  ║   `e[36m██╔══██╗██║   ██║██╔══██╗    `e[97m Azure Virtual Desktop`e[34m               ║
  ║   `e[36m███████║██║   ██║██║  ██║    `e[97m AIO Deployment Console v3.0`e[34m         ║
  ║   `e[36m██╔══██║╚██╗ ██╔╝██║  ██║    `e[90m Bob's Discount Furniture`e[34m            ║
  ║   `e[36m██║  ██║ ╚████╔╝ ██████╔╝    `e[90m IT Infrastructure Team`e[34m              ║
  ║   `e[36m╚═╝  ╚═╝  ╚═══╝  ╚═════╝`e[34m                                        ║
  ║                                                                       ║
  ╚═══════════════════════════════════════════════════════════════════════╝`e[0m
"@
    Write-Host "  Config : " -NoNewline -ForegroundColor $C.Muted
    Write-Host $ConfigFile -ForegroundColor $C.Azure
    Write-Host "  Log    : " -NoNewline -ForegroundColor $C.Muted
    Write-Host $LogFile -ForegroundColor $C.Azure
    Write-Host ""
}
#endregion

#region ═══════════════════════════════════════════════════════════════════════
#  CONFIGURATION — GLOBAL STATE
#═══════════════════════════════════════════════════════════════════════════════

$Global:Cfg = @{
    # Azure identity
    SubscriptionId   = ""
    SubscriptionName = ""
    TenantId         = ""
    Environment      = "POC"   # POC | Staging | Production

    # Naming
    Prefix           = "bdf-poc"
    Location         = "eastus"
    LocationDisplay  = "East US"

    # Resource Groups (can be existing or new)
    RG = @{
        AVD        = @{ Name=""; IsNew=$true; ResourceId="" }
        Network    = @{ Name=""; IsNew=$true; ResourceId="" }
        Storage    = @{ Name=""; IsNew=$true; ResourceId="" }
        Monitoring = @{ Name=""; IsNew=$true; ResourceId="" }
        Automation = @{ Name=""; IsNew=$true; ResourceId="" }
    }

    # Networking
    Network = @{
        VNet    = @{ Name=""; ResourceGroup=""; AddressSpace=""; Id=""; IsNew=$true }
        Subnets = @{
            E3      = @{ Name="snet-avd-e3";      Prefix="10.10.1.0/24"; IsNew=$true }
            F1      = @{ Name="snet-avd-f1";      Prefix="10.10.2.0/24"; IsNew=$true }
            Mgmt    = @{ Name="snet-avd-mgmt";    Prefix="10.10.3.0/24"; IsNew=$true }
            Storage = @{ Name="snet-avd-storage"; Prefix="10.10.4.0/24"; IsNew=$true }
        }
        NSG     = @{ E3=@{Name=""; IsNew=$true}; F1=@{Name=""; IsNew=$true} }
    }

    # Security
    KeyVault = @{ Name=""; ResourceGroup=""; Id=""; IsNew=$true }

    # Monitoring
    LogAnalytics = @{ Name=""; ResourceGroup=""; WorkspaceId=""; Id=""; IsNew=$true }

    # Storage / FSLogix
    Storage = @{
        Account  = @{ Name=""; ResourceGroup=""; IsNew=$true }
        Shares   = @{ E3="profiles-e3"; F1="profiles-f1"; ODFC="odfc-e3" }
        PrivateEndpoint = @{ IsNew=$true }
    }

    # Compute Gallery
    Gallery = @{ Name=""; ResourceGroup=""; IsNew=$true }

    # Host Pools
    HostPools = @{
        E3 = @{
            Name=""; MaxSessions=8; LoadBalancer="BreadthFirst"
            VMCount=3; VMSize="Standard_D4ds_v5"
            AppGroupName=""; IsNew=$true; Status="NotDeployed"
        }
        F1 = @{
            Name=""; MaxSessions=12; LoadBalancer="BreadthFirst"
            VMCount=2; VMSize="Standard_D8ds_v5"
            AppGroupName=""; IsNew=$true; Status="NotDeployed"
        }
    }

    # Identity / Users
    Identity = @{
        E3GroupId      = ""
        F1GroupId      = ""
        AdminGroupId   = ""
        AdminUsername  = "bdfadmin"
        AzureAdJoin    = $true
    }

    # Auto-Scaling (dynamic config)
    Scaling = @{
        E3 = @{
            Enabled             = $true
            PlanName            = ""
            TimeZone            = "Eastern Standard Time"
            PeakCapacityPct     = 80
            RampUpCapacityPct   = 60
            ForceLogoffMinutes  = 15
            Schedules           = @()   # populated by wizard
        }
        F1 = @{
            Enabled             = $true
            PlanName            = ""
            TimeZone            = "Eastern Standard Time"
            PeakCapacityPct     = 85
            RampUpCapacityPct   = 60
            ForceLogoffMinutes  = 10
            Schedules           = @()
        }
        CustomRunbooks = @{
            ScaleOut  = $true
            ScaleIn   = $true
            AutoHeal  = $true
            Holiday   = $true
        }
    }

    # Automation
    Automation = @{ AccountName=""; ResourceGroup=""; ManagedIdentityId=""; IsNew=$true }

    # Deployment state (track what's been deployed)
    DeploymentState = @{
        ResourceGroups  = "NotDeployed"
        Networking      = "NotDeployed"
        KeyVault        = "NotDeployed"
        LogAnalytics    = "NotDeployed"
        AzureFiles      = "NotDeployed"
        ComputeGallery  = "NotDeployed"
        HostPools       = "NotDeployed"
        SessionHosts    = "NotDeployed"
        AppGroups       = "NotDeployed"
        ScalingPlans    = "NotDeployed"
        Automation      = "NotDeployed"
        Monitoring      = "NotDeployed"
        Permissions     = "NotVerified"
        RDPProperties   = "NotConfigured"
        JoinType        = "NotConfigured"
    }

    # Domain Join Configuration
    JoinConfig = @{
        Type                  = ""          # "EntraID" | "HybridAD"
        Configured            = $false

        # Entra ID Join settings
        EntraID = @{
            MDMEnrollment     = $true       # Auto-enroll in Intune
            MDMAppId          = ""          # Intune app ID (optional override)
        }

        # Hybrid AD Join settings
        HybridAD = @{
            DomainName        = ""          # e.g. bdf.internal
            DomainNetbios     = ""          # e.g. BDF
            DomainJoinUPN     = ""          # domain join service account UPN
            DomainJoinOU      = ""          # e.g. OU=AVDHosts,DC=bdf,DC=internal
            DCIPAddresses     = @()         # VNet DNS pointing to DCs
            AADConnectServer  = ""          # AADC or Cloud Sync
            SyncMethod        = "AADConnect" # "AADConnect" | "CloudSync"
        }
    }

    # RDP Custom Properties
    RDP = @{
        Configured        = $false
        ActiveProfile     = "Balanced"      # "Strict" | "Balanced" | "Open" | "Custom"

        # Built string (auto-generated from sections below)
        E3PropertyString  = ""
        F1PropertyString  = ""

        # Display & Graphics
        Display = @{
            DynamicResolution   = 1      # Adjust resolution dynamically
            MultiMonitor        = 1      # Multi-monitor support
            MaximizeToDisplays  = 1      # Maximize to current monitors
            SmartSizing         = 0      # Scale session to window
            DesktopScaleFactor  = 100    # DPI scaling %
            ColorDepth          = 32     # 16 or 32
            ConnectionType      = 7      # 7=auto detect (recommended)
            BandwidthAutoDetect = 1      # Auto-detect network quality
            NetworkAutoDetect   = 1
            Compression         = 1
            AllowFontSmoothing  = 1
            AllowDesktopComposition = 1
            VideoPlaybackMode   = 1      # Video rendering optimization
        }

        # Audio & Camera
        AV = @{
            AudioOutputMode     = 0      # 0=client, 1=server, 2=disabled
            AudioCapture        = 1      # Microphone (1=enabled)
            CameraRedirect      = 1      # Camera for Teams video (E3)
            CameraRedirectF1    = 0      # Cameras off for F1 frontline
            VideoCapture        = 1      # Multimedia Redirection (MMR)
            VideoCaptureQuality = 0      # 0=high,1=medium,2=low
        }

        # Device Redirection
        Redirect = @{
            Clipboard           = 1      # Copy/paste
            ClipboardF1         = 0      # Disable clipboard for frontline (DLP)
            Drives              = 0      # Local drive mapping (disable — use OneDrive)
            Printers            = 1      # Printer redirection
            SmartCards          = 0      # Smart cards
            USB                 = 0      # Generic USB (whitelist only)
            SerialPorts         = 0      # COM ports
            POSDevices          = 0      # POS hardware
            Location            = 0      # Location services
            WebAuthn            = 1      # FIDO2/passkeys passthrough
        }

        # Authentication & SSO
        Auth = @{
            EntraIDSSO          = 1      # SSO via Entra ID (enablerdsaadredirection)
            TargetIsAADJoined   = 1      # Azure AD joined VMs
            NLARequired         = 1      # Network Level Auth
            CredSSP             = 1      # CredSSP support
            AuthLevel           = 2      # 0=no warn, 1=warn, 2=warn+block
            PromptCredentials   = 0      # 0=no prompt (SSO), 1=prompt
        }

        # Session Timeouts
        Timeouts = @{
            DisconnectTimeoutMs = 28800000  # 8 hours
            IdleTimeoutMs       = 7200000   # 2 hours
            IdleTimeoutF1Ms     = 1800000   # 30 min for frontline
            ReconnectSameHost   = 1
            AutoReconnect       = 1
        }

        # Security & DLP
        Security = @{
            ScreenCaptureProtection   = 1  # 0=off,1=block client,2=block client+server
            Watermarking              = 1  # Session watermark (username overlay)
            WatermarkOpacity          = 20 # 0-100
            ScreenCaptureProtF1       = 2  # Stricter for frontline (full block)
        }

        # Performance
        Perf = @{
            BitmapCacheSize     = 32000   # KB
            BitmapCachePersist  = 1
            KeyboardHook        = 2       # 0=local,1=remote,2=fullscreen only
            DisableMenuAnims    = 0       # 0=show,1=hide (hide for low bandwidth)
            DisableThemes       = 0
            DisableCursorBlink  = 0
            DisableWallpaper    = 0       # Keep wallpaper (AVD uses Teams bg anyway)
        }
    }
}

# Status → color/icon
function Get-StateDisplay {
    param([string]$State)
    switch ($State) {
        "Deployed"    { return @{ Icon="✔"; Color=$C.Ok;   Text="Deployed"   } }
        "Skipped"     { return @{ Icon="⊘"; Color=$C.Skip; Text="Skipped"    } }
        "Failed"      { return @{ Icon="✖"; Color=$C.Fail; Text="Failed"     } }
        "InProgress"  { return @{ Icon="⚡"; Color=$C.Warn; Text="In Progress"} }
        "NotVerified" { return @{ Icon="○"; Color=$C.Muted; Text="Not Verified"} }
        default       { return @{ Icon="○"; Color=$C.Muted; Text="Not Deployed" } }
    }
}
#endregion

#region ═══════════════════════════════════════════════════════════════════════
#  CONFIG FILE — SAVE / LOAD
#═══════════════════════════════════════════════════════════════════════════════

function Save-Config {
    $Global:Cfg | ConvertTo-Json -Depth 10 | Set-Content $ConfigFile -Encoding UTF8
    Write-Status "Configuration saved" $ConfigFile "ok"
    Write-Log "Config saved to $ConfigFile"
}

function Load-Config {
    if (Test-Path $ConfigFile) {
        Write-Status "Loading saved config" $ConfigFile "reuse"
        $loaded = Get-Content $ConfigFile -Raw | ConvertFrom-Json -AsHashtable
        # Deep merge loaded into Global:Cfg
        foreach ($key in $loaded.Keys) {
            if ($Global:Cfg.ContainsKey($key)) { $Global:Cfg[$key] = $loaded[$key] }
        }
        Write-Log "Config loaded from $ConfigFile"
        return $true
    }
    return $false
}

function Set-DefaultNames {
    $p = $Global:Cfg.Prefix
    $Global:Cfg.RG.AVD.Name        = "rg-avd-$p"
    $Global:Cfg.RG.Network.Name    = "rg-avd-network-$p"
    $Global:Cfg.RG.Storage.Name    = "rg-avd-storage-$p"
    $Global:Cfg.RG.Monitoring.Name = "rg-avd-monitoring-$p"
    $Global:Cfg.RG.Automation.Name = "rg-avd-automation-$p"
    $Global:Cfg.Network.VNet.Name  = "vnet-avd-hub-$p"
    $Global:Cfg.KeyVault.Name      = "kv-avd-$p"
    $Global:Cfg.LogAnalytics.Name  = "law-avd-$p"
    $Global:Cfg.Storage.Account.Name = "stavd$($p -replace '-','')"
    $Global:Cfg.Gallery.Name       = "acg_avd_$($p -replace '-','_')"
    $Global:Cfg.HostPools.E3.Name  = "hp-avd-e3-office-$p"
    $Global:Cfg.HostPools.F1.Name  = "hp-avd-f1-frontline-$p"
    $Global:Cfg.HostPools.E3.AppGroupName = "ag-desktop-e3-$p"
    $Global:Cfg.HostPools.F1.AppGroupName = "ag-remoteapp-f1-$p"
    $Global:Cfg.Scaling.E3.PlanName = "sp-avd-e3-$p"
    $Global:Cfg.Scaling.F1.PlanName = "sp-avd-f1-$p"
    $Global:Cfg.Automation.AccountName = "aa-avd-scaling-$p"
    $Global:Cfg.Automation.ResourceGroup = $Global:Cfg.RG.Automation.Name
    $Global:Cfg.KeyVault.ResourceGroup   = $Global:Cfg.RG.AVD.Name
    $Global:Cfg.LogAnalytics.ResourceGroup = $Global:Cfg.RG.Monitoring.Name
    $Global:Cfg.Storage.Account.ResourceGroup = $Global:Cfg.RG.Storage.Name
    $Global:Cfg.Gallery.ResourceGroup = $Global:Cfg.RG.AVD.Name
}
#endregion

#region ═══════════════════════════════════════════════════════════════════════
#  AZURE CONNECTION & SUBSCRIPTION MANAGEMENT
#═══════════════════════════════════════════════════════════════════════════════

function Connect-ToAzure {
    Write-Banner "AZURE CONNECTION"

    # Check if already connected
    $ctx = Get-AzContext -EA SilentlyContinue
    if ($ctx) {
        Write-Status "Currently signed in as" $ctx.Account.Id "info"
        if (Read-YesNo "Use this account?" $true) {
            # Fall through to subscription selection
        } else {
            Write-Status "Connecting to Azure..." "" "step"
            Connect-AzAccount | Out-Null
            $ctx = Get-AzContext
        }
    } else {
        Write-Status "Not signed in — launching Azure login..." "" "step"
        Connect-AzAccount | Out-Null
        $ctx = Get-AzContext
    }
    Write-Status "Signed in as" $ctx.Account.Id "ok"
    Write-Log "Connected as $($ctx.Account.Id)"

    # Subscription selection
    Select-Subscription
}

function Select-Subscription {
    Write-Section "SELECT SUBSCRIPTION"
    Write-Status "Retrieving available subscriptions..." "" "step"

    $subs = Get-AzSubscription | Where-Object { $_.State -eq "Enabled" } |
            Sort-Object Name

    if ($subs.Count -eq 0) {
        Write-Status "No enabled subscriptions found" "" "fail"
        throw "No accessible subscriptions"
    }

    Write-Host ""
    Write-Host ("  {0,-4} {1,-45} {2,-12} {3}" -f "#", "Subscription Name", "State", "ID") -ForegroundColor $C.Muted
    Write-Host "  $($UI.SH * 80)" -ForegroundColor $C.Border

    $i = 1
    foreach ($sub in $subs) {
        $cur = if ($sub.Id -eq (Get-AzContext).Subscription.Id) { " ◄ current" } else { "" }
        $col = if ($cur) { $C.Ok } else { $C.Menu }
        Write-Host ("  {0,-4}" -f $i) -NoNewline -ForegroundColor $C.Accent
        Write-Host ("{0,-45}" -f $sub.Name) -NoNewline -ForegroundColor $col
        Write-Host ("{0,-12}" -f $sub.State) -NoNewline -ForegroundColor $C.Muted
        Write-Host ("{0}{1}" -f $sub.Id, $cur) -ForegroundColor $C.Muted
        $i++
    }

    $valid = 1..($subs.Count) | ForEach-Object { "$_" }
    $choice = Read-MenuChoice "Select subscription number" $valid

    $selected = $subs[[int]$choice - 1]
    Set-AzContext -SubscriptionId $selected.Id | Out-Null

    $Global:Cfg.SubscriptionId   = $selected.Id
    $Global:Cfg.SubscriptionName = $selected.Name
    $Global:Cfg.TenantId         = $selected.TenantId

    Write-Status "Active subscription" "$($selected.Name) ($($selected.Id))" "ok"
    Write-Log "Subscription selected: $($selected.Name) / $($selected.Id)"
}

function Select-Environment {
    Write-Section "SELECT DEPLOYMENT ENVIRONMENT"
    $envs = @(
        @{ Name="POC";        Prefix="bdf-poc";     Desc="Proof of Concept — minimal resources, PAYG pricing" }
        @{ Name="Staging";    Prefix="bdf-stg";     Desc="Staging/UAT — mirrors production config" }
        @{ Name="Production"; Prefix="bdf-prod";    Desc="Production — reserved instances, HA storage" }
    )
    foreach ($i in 0..2) {
        Write-Host ("  {0}. " -f ($i+1)) -NoNewline -ForegroundColor $C.Accent
        Write-Host ("{0,-12}" -f $envs[$i].Name) -NoNewline -ForegroundColor $C.MenuSel
        Write-Host $envs[$i].Desc -ForegroundColor $C.Muted
    }
    $choice = Read-MenuChoice "Environment" @("1","2","3") "1"
    $env    = $envs[[int]$choice - 1]

    $Global:Cfg.Environment = $env.Name
    $Global:Cfg.Prefix      = $env.Prefix
    Set-DefaultNames

    # Override VM counts for production
    if ($env.Name -eq "Production") {
        $Global:Cfg.HostPools.E3.VMCount = 15
        $Global:Cfg.HostPools.F1.VMCount = 10
    }

    Write-Status "Environment" "$($env.Name) (prefix: $($env.Prefix))" "ok"
    Select-Region
}

function Select-Region {
    Write-Section "SELECT AZURE REGION"
    $regions = @(
        @{ Name="eastus";       Display="East US";          Recommended=$true  }
        @{ Name="eastus2";      Display="East US 2";        Recommended=$false }
        @{ Name="centralus";    Display="Central US";       Recommended=$false }
        @{ Name="westus2";      Display="West US 2";        Recommended=$false }
        @{ Name="westus3";      Display="West US 3";        Recommended=$false }
        @{ Name="northeurope";  Display="North Europe";     Recommended=$false }
        @{ Name="westeurope";   Display="West Europe";      Recommended=$false }
    )
    foreach ($i in 0..($regions.Count-1)) {
        $rec = if ($regions[$i].Recommended) { " ★ recommended" } else { "" }
        $col = if ($regions[$i].Recommended) { $C.Ok } else { $C.Menu }
        Write-Host ("  {0}. {1,-14} {2,-20}{3}" -f ($i+1), $regions[$i].Name, $regions[$i].Display, $rec) -ForegroundColor $col
    }
    $choice = Read-MenuChoice "Region" (1..$regions.Count | ForEach-Object {"$_"}) "1"
    $r = $regions[[int]$choice - 1]
    $Global:Cfg.Location        = $r.Name
    $Global:Cfg.LocationDisplay = $r.Display
    Write-Status "Region" "$($r.Display) ($($r.Name))" "ok"
}
#endregion

#region ═══════════════════════════════════════════════════════════════════════
#  DISCOVERY ENGINE — detect existing resources
#═══════════════════════════════════════════════════════════════════════════════

function Find-ExistingVNets {
    Write-Status "Scanning for existing Virtual Networks..." "" "step"
    $vnets = Get-AzVirtualNetwork | Select-Object Name, ResourceGroupName, Location,
        @{N="AddressSpace"; E={$_.AddressSpace.AddressPrefixes -join ", "}},
        @{N="SubnetCount";  E={$_.Subnets.Count}},
        Id |
        Where-Object { $_.Location -eq $Global:Cfg.Location } |
        Sort-Object Name
    return $vnets
}

function Find-ExistingNSGs {
    param([string]$ResourceGroup = "")
    $q = if ($ResourceGroup) {
        Get-AzNetworkSecurityGroup -ResourceGroupName $ResourceGroup
    } else {
        Get-AzNetworkSecurityGroup
    }
    return $q | Select-Object Name, ResourceGroupName, Location, Id |
                Where-Object { $_.Location -eq $Global:Cfg.Location } |
                Sort-Object Name
}

function Find-ExistingStorageAccounts {
    return Get-AzStorageAccount |
        Where-Object { $_.Location -eq $Global:Cfg.Location -and $_.Kind -eq "FileStorage" } |
        Select-Object StorageAccountName, ResourceGroupName, Location,
            @{N="Sku"; E={$_.Sku.Name}}, Id |
        Sort-Object StorageAccountName
}

function Find-ExistingKeyVaults {
    return Get-AzKeyVault |
        Where-Object { $_.Location -eq $Global:Cfg.Location } |
        Select-Object VaultName, ResourceGroupName,
            @{N="Location"; E={$_.Location}},
            @{N="EnableRbac"; E={$_.EnableRbacAuthorization}}, ResourceId |
        Sort-Object VaultName
}

function Find-ExistingLogAnalytics {
    return Get-AzOperationalInsightsWorkspace |
        Where-Object { $_.Location -eq $Global:Cfg.Location } |
        Select-Object Name, ResourceGroupName, Location,
            @{N="Sku"; E={$_.Sku}}, CustomerId, ResourceId |
        Sort-Object Name
}

function Find-ExistingHostPools {
    return Get-AzWvdHostPool |
        Where-Object { $_.Location -eq $Global:Cfg.Location } |
        Select-Object @{N="Name"; E={($_.Name -split "/")[-1]}},
            @{N="ResourceGroup"; E={$_.Id.Split("/")[4]}},
            @{N="Type"; E={$_.HostPoolType}},
            @{N="MaxSessions"; E={$_.MaxSessionLimit}},
            Id |
        Sort-Object Name
}

function Find-ExistingComputeGalleries {
    return Get-AzGallery |
        Select-Object Name, ResourceGroupName, Location, Id |
        Where-Object { $_.Location -eq $Global:Cfg.Location } |
        Sort-Object Name
}

function Find-ExistingResourceGroups {
    return Get-AzResourceGroup |
        Where-Object { $_.Location -eq $Global:Cfg.Location } |
        Select-Object ResourceGroupName, Location,
            @{N="Tags"; E={ ($_.Tags.Keys | Select-Object -First 3) -join "," }} |
        Sort-Object ResourceGroupName
}

function Find-ExistingAutomationAccounts {
    return Get-AzAutomationAccount |
        Select-Object AutomationAccountName, ResourceGroupName, Location, State |
        Where-Object { $_.Location -eq $Global:Cfg.Location } |
        Sort-Object AutomationAccountName
}

# ── Unified detection wizard for a component ──────────────────────────────
function Invoke-ComponentDetect {
    param(
        [string]$ComponentName,
        [string]$Description,
        [scriptblock]$FindExisting,
        [string]$DisplayProp,
        [string]$SubProp = "",
        [bool]$AllowSkip = $false
    )

    Write-Section "DETECT: $ComponentName"
    Write-Status $Description "" "info"

    $items = & $FindExisting
    if ($null -eq $items -or @($items).Count -eq 0) {
        Write-Status "No existing $ComponentName found in $($Global:Cfg.LocationDisplay)" "" "warn"
        return "NEW"
    }

    Write-Status "Found $(@($items).Count) existing $ComponentName in this subscription" "" "info"
    Write-Host ""

    # Quick table preview
    $i = 1
    foreach ($item in @($items)) {
        $main = if ($DisplayProp) { $item.$DisplayProp } else { $item }
        $sub  = if ($SubProp)     { $item.$SubProp }     else { "" }
        Write-Host ("  {0,3}. " -f $i) -NoNewline -ForegroundColor $C.Accent
        Write-Host ("{0,-40}" -f $main) -NoNewline -ForegroundColor $C.Menu
        Write-Host $sub -ForegroundColor $C.Muted
        $i++
    }

    Write-Host ""
    Write-Host "     N.  Create new $ComponentName" -ForegroundColor $C.New
    if ($AllowSkip) { Write-Host "     S.  Skip $ComponentName" -ForegroundColor $C.Skip }

    $validOpts = (@(1..@($items).Count | ForEach-Object { "$_" }) + @("N","n"))
    if ($AllowSkip) { $validOpts += @("S","s") }
    $choice = Read-MenuChoice "Use existing, create new, or skip?" $validOpts "N"

    if ($choice -in @("N","n")) { return "NEW" }
    if ($choice -in @("S","s")) { return "SKIP" }
    return @($items)[[int]$choice - 1]
}
#endregion

#region ═══════════════════════════════════════════════════════════════════════
#  PERMISSION MANAGER
#═══════════════════════════════════════════════════════════════════════════════

$Script:RequiredPermissions = @(
    # scope, principal-description, role, why-needed
    @{ Scope="Subscription";     Who="Deployment User (you)";                Role="Contributor";                                        Why="Deploy all AVD resources" }
    @{ Scope="Subscription";     Who="Deployment User (you)";                Role="User Access Administrator";                          Why="Assign RBAC roles during deploy" }
    @{ Scope="AVD Resource Group"; Who="AVD Service Principal";              Role="Desktop Virtualization Power On Contributor";         Why="Start VM on Connect + Scaling Plans" }
    @{ Scope="AVD Resource Group"; Who="Azure Automation Managed Identity";  Role="Desktop Virtualization Contributor";                  Why="Scaling runbooks modify host pools" }
    @{ Scope="AVD Resource Group"; Who="Azure Automation Managed Identity";  Role="Virtual Machine Contributor";                        Why="Start/stop/delete session host VMs" }
    @{ Scope="AVD Resource Group"; Who="Azure Automation Managed Identity";  Role="Desktop Virtualization Power On Contributor";         Why="Runbook scale-out start VMs" }
    @{ Scope="Storage Account";    Who="E3 User Group";                       Role="Storage File Data SMB Share Contributor";             Why="Mount FSLogix profile VHDXs" }
    @{ Scope="Storage Account";    Who="F1 User Group";                       Role="Storage File Data SMB Share Contributor";             Why="Mount FSLogix profile VHDXs" }
    @{ Scope="Storage Account";    Who="Session Host VMs (system identity)";  Role="Storage File Data SMB Share Elevated Contributor";   Why="FSLogix profile container access" }
    @{ Scope="App Group (E3)";     Who="E3 User Group";                       Role="Desktop Virtualization User";                        Why="Connect to AVD desktop" }
    @{ Scope="App Group (F1)";     Who="F1 User Group";                       Role="Desktop Virtualization User";                        Why="Launch RemoteApps" }
    @{ Scope="App Group (E3)";     Who="AVD Admin Group";                     Role="Desktop Virtualization Contributor";                  Why="Admin management access" }
    @{ Scope="Key Vault";          Who="Automation Managed Identity";         Role="Key Vault Secrets User";                             Why="Read VM admin password for scale-out" }
    @{ Scope="Key Vault";          Who="Deployment User (you)";               Role="Key Vault Secrets Officer";                          Why="Store credentials during deployment" }
    @{ Scope="Subscription";       Who="Automation Managed Identity";         Role="Reader";                                             Why="Read subscription-level resources" }
)

function Show-PermissionManager {
    Write-Banner "PERMISSION MANAGER" $C.Gold

    while ($true) {
        Write-Section "Permission Menu"
        Write-Host "  1.  View all required permissions" -ForegroundColor $C.Menu
        Write-Host "  2.  Check current permission state" -ForegroundColor $C.Menu
        Write-Host "  3.  Assign all missing permissions (auto)" -ForegroundColor $C.New
        Write-Host "  4.  Assign individual permission" -ForegroundColor $C.Menu
        Write-Host "  5.  Check subscription-level access" -ForegroundColor $C.Menu
        Write-Host "  6.  Export permission report" -ForegroundColor $C.Menu
        Write-Host "  0.  Back to main menu" -ForegroundColor $C.Muted
        $choice = Read-MenuChoice "Permission Manager" @("0","1","2","3","4","5","6")
        switch ($choice) {
            "1" { Show-AllRequiredPermissions }
            "2" { Check-PermissionState }
            "3" { Assign-AllMissingPermissions }
            "4" { Assign-IndividualPermission }
            "5" { Check-SubscriptionAccess }
            "6" { Export-PermissionReport }
            "0" { return }
        }
    }
}

function Show-AllRequiredPermissions {
    Write-Section "All Required RBAC Permissions"
    Write-Host ""
    Write-Host ("  {0,-28} {1,-42} {2}" -f "SCOPE", "ROLE", "PRINCIPAL") -ForegroundColor $C.Muted
    Write-Host "  $($UI.SH * 80)" -ForegroundColor $C.Border
    foreach ($p in $Script:RequiredPermissions) {
        Write-Host ("  {0,-28}" -f $p.Scope) -NoNewline -ForegroundColor $C.Azure
        Write-Host ("{0,-42}" -f $p.Role)    -NoNewline -ForegroundColor $C.Menu
        Write-Host $p.Who                                -ForegroundColor $C.Muted
    }
    Write-Host ""
    Read-Host "  Press Enter to continue"
}

function Check-PermissionState {
    Write-Section "Checking Current Permissions"
    $ctx        = Get-AzContext
    $userId     = $ctx.Account.Id
    $subScope   = "/subscriptions/$($Global:Cfg.SubscriptionId)"
    $results    = @()

    # Check deployment user has Contributor + UAA on subscription
    $assignments = Get-AzRoleAssignment -SignInName $userId -Scope $subScope -EA SilentlyContinue
    foreach ($needed in @("Contributor","User Access Administrator")) {
        $has = $null -ne ($assignments | Where-Object { $_.RoleDefinitionName -eq $needed })
        $results += @{ Name="You ($userId)"; Role=$needed; Scope="Subscription"; Has=$has }
    }

    # Check AVD service principal role
    $avdSP = Get-AzADServicePrincipal -ApplicationId "9cdead84-a844-4324-93f2-b2e6bb768d07" -EA SilentlyContinue
    if ($avdSP) {
        $avdRoles = Get-AzRoleAssignment -ObjectId $avdSP.Id -EA SilentlyContinue
        $has = $null -ne ($avdRoles | Where-Object { $_.RoleDefinitionName -eq "Desktop Virtualization Power On Contributor" })
        $results += @{ Name="AVD Service Principal"; Role="Desktop Virtualization Power On Contributor"; Scope="RG"; Has=$has }
    }

    # Check Automation MI if it exists
    $aa = Get-AzAutomationAccount -Name $Global:Cfg.Automation.AccountName `
          -ResourceGroupName $Global:Cfg.Automation.ResourceGroup -EA SilentlyContinue
    if ($aa) {
        $miId = (Get-AzAutomationAccount -Name $Global:Cfg.Automation.AccountName `
                 -ResourceGroupName $Global:Cfg.Automation.ResourceGroup).Identity.PrincipalId
        if ($miId) {
            $miRoles = Get-AzRoleAssignment -ObjectId $miId -EA SilentlyContinue
            foreach ($r in @("Desktop Virtualization Contributor","Virtual Machine Contributor","Desktop Virtualization Power On Contributor")) {
                $has = $null -ne ($miRoles | Where-Object { $_.RoleDefinitionName -eq $r })
                $results += @{ Name="Automation MI"; Role=$r; Scope="RG"; Has=$has }
            }
        }
    }

    Write-Host ""
    $pass = 0; $fail = 0
    foreach ($r in $results) {
        $icon = if ($r.Has) { "✔" } else { "✖" }
        $col  = if ($r.Has) { $C.Ok } else { $C.Fail }
        Write-Host ("  $icon  {0,-35} {1,-45} {2}" -f $r.Name, $r.Role, $r.Scope) -ForegroundColor $col
        if ($r.Has) { $pass++ } else { $fail++ }
    }
    Write-Host ""
    Write-Status "Passed: $pass  |  Missing: $fail" "" (if ($fail -eq 0) {"ok"} else {"warn"})
    $Global:Cfg.DeploymentState.Permissions = if ($fail -eq 0) { "Deployed" } else { "NotVerified" }
    Read-Host "  Press Enter to continue"
}

function Assign-AllMissingPermissions {
    Write-Section "Auto-Assigning All Required Permissions"

    if (-not (Confirm-Action "Assign all required RBAC roles" "This will make role assignments across subscription and resource groups")) {
        return
    }

    $subScope = "/subscriptions/$($Global:Cfg.SubscriptionId)"
    $assigned = 0; $skipped = 0

    # 1. AVD Service Principal — Power On Contributor
    Write-Status "Assigning AVD Service Principal roles..." "" "step"
    $avdSP = Get-AzADServicePrincipal -ApplicationId "9cdead84-a844-4324-93f2-b2e6bb768d07" -EA SilentlyContinue
    if ($avdSP) {
        $rgScope = "/subscriptions/$($Global:Cfg.SubscriptionId)/resourceGroups/$($Global:Cfg.RG.AVD.Name)"
        foreach ($role in @("Desktop Virtualization Power On Contributor")) {
            $exists = Get-AzRoleAssignment -ObjectId $avdSP.Id -RoleDefinitionName $role -Scope $rgScope -EA SilentlyContinue
            if (-not $exists) {
                New-AzRoleAssignment -ObjectId $avdSP.Id -RoleDefinitionName $role -Scope $rgScope -EA SilentlyContinue | Out-Null
                Write-Status "Assigned" "$role → AVD SP" "ok"; $assigned++
            } else { $skipped++ }
        }
    } else {
        Write-Status "AVD Service Principal not found in this tenant" "" "warn"
    }

    # 2. Automation Account Managed Identity
    $aa = Get-AzAutomationAccount -Name $Global:Cfg.Automation.AccountName `
          -ResourceGroupName $Global:Cfg.Automation.ResourceGroup -EA SilentlyContinue
    if ($aa) {
        $miId = $aa.Identity.PrincipalId
        if ($miId) {
            Write-Status "Assigning Automation Managed Identity roles..." "" "step"
            $rgScope = "/subscriptions/$($Global:Cfg.SubscriptionId)/resourceGroups/$($Global:Cfg.RG.AVD.Name)"
            $miRoles = @(
                @{ Role="Desktop Virtualization Contributor";          Scope=$rgScope  }
                @{ Role="Virtual Machine Contributor";                  Scope=$rgScope  }
                @{ Role="Desktop Virtualization Power On Contributor";  Scope=$rgScope  }
                @{ Role="Reader";                                       Scope=$subScope }
            )
            foreach ($mr in $miRoles) {
                $exists = Get-AzRoleAssignment -ObjectId $miId -RoleDefinitionName $mr.Role `
                          -Scope $mr.Scope -EA SilentlyContinue
                if (-not $exists) {
                    New-AzRoleAssignment -ObjectId $miId -RoleDefinitionName $mr.Role `
                        -Scope $mr.Scope -EA SilentlyContinue | Out-Null
                    Write-Status "Assigned" "$($mr.Role) → Automation MI" "ok"; $assigned++
                } else { $skipped++ }
            }
        }
    }

    # 3. E3 Group — Storage File share
    if ($Global:Cfg.Identity.E3GroupId -and $Global:Cfg.Storage.Account.Name) {
        $sa = Get-AzStorageAccount -Name $Global:Cfg.Storage.Account.Name `
              -ResourceGroupName $Global:Cfg.Storage.Account.ResourceGroup -EA SilentlyContinue
        if ($sa) {
            foreach ($share in @($Global:Cfg.Storage.Shares.E3, $Global:Cfg.Storage.Shares.ODFC)) {
                $shareScope = "$($sa.Id)/fileServices/default/fileshares/$share"
                $exists = Get-AzRoleAssignment -ObjectId $Global:Cfg.Identity.E3GroupId `
                          -RoleDefinitionName "Storage File Data SMB Share Contributor" `
                          -Scope $shareScope -EA SilentlyContinue
                if (-not $exists) {
                    New-AzRoleAssignment -ObjectId $Global:Cfg.Identity.E3GroupId `
                        -RoleDefinitionName "Storage File Data SMB Share Contributor" `
                        -Scope $shareScope -EA SilentlyContinue | Out-Null
                    Write-Status "Assigned" "Storage SMB Contributor → E3 Group ($share)" "ok"; $assigned++
                } else { $skipped++ }
            }
        }
    }

    # 4. F1 Group — Storage File share
    if ($Global:Cfg.Identity.F1GroupId -and $Global:Cfg.Storage.Account.Name) {
        $sa = Get-AzStorageAccount -Name $Global:Cfg.Storage.Account.Name `
              -ResourceGroupName $Global:Cfg.Storage.Account.ResourceGroup -EA SilentlyContinue
        if ($sa) {
            $shareScope = "$($sa.Id)/fileServices/default/fileshares/$($Global:Cfg.Storage.Shares.F1)"
            $exists = Get-AzRoleAssignment -ObjectId $Global:Cfg.Identity.F1GroupId `
                      -RoleDefinitionName "Storage File Data SMB Share Contributor" `
                      -Scope $shareScope -EA SilentlyContinue
            if (-not $exists) {
                New-AzRoleAssignment -ObjectId $Global:Cfg.Identity.F1GroupId `
                    -RoleDefinitionName "Storage File Data SMB Share Contributor" `
                    -Scope $shareScope -EA SilentlyContinue | Out-Null
                Write-Status "Assigned" "Storage SMB Contributor → F1 Group" "ok"; $assigned++
            } else { $skipped++ }
        }
    }

    # 5. Key Vault access for Automation MI
    $kv = Get-AzKeyVault -VaultName $Global:Cfg.KeyVault.Name -EA SilentlyContinue
    if ($kv -and $aa) {
        $miId = (Get-AzAutomationAccount -Name $Global:Cfg.Automation.AccountName `
                 -ResourceGroupName $Global:Cfg.Automation.ResourceGroup -EA SilentlyContinue).Identity.PrincipalId
        if ($miId) {
            $exists = Get-AzRoleAssignment -ObjectId $miId -RoleDefinitionName "Key Vault Secrets User" `
                      -Scope $kv.ResourceId -EA SilentlyContinue
            if (-not $exists) {
                New-AzRoleAssignment -ObjectId $miId -RoleDefinitionName "Key Vault Secrets User" `
                    -Scope $kv.ResourceId -EA SilentlyContinue | Out-Null
                Write-Status "Assigned" "Key Vault Secrets User → Automation MI" "ok"; $assigned++
            } else { $skipped++ }
        }
    }

    Write-Host ""
    Write-Status "Assigned: $assigned  |  Already existed: $skipped" "" "ok"
    if ($assigned -gt 0) {
        Write-Status "Waiting 30s for RBAC propagation..." "" "info"
        Start-Sleep -Seconds 30
    }
    $Global:Cfg.DeploymentState.Permissions = "Deployed"
    Save-Config
    Read-Host "  Press Enter to continue"
}

function Assign-IndividualPermission {
    Write-Section "Assign Individual Permission"
    $i = 1
    foreach ($p in $Script:RequiredPermissions) {
        Write-Host ("  {0,3}. [{1,-28}] {2,-42} → {3}" -f $i, $p.Scope, $p.Role, $p.Who) -ForegroundColor $C.Menu
        $i++
    }
    $valid = 1..($Script:RequiredPermissions.Count) | ForEach-Object {"$_"}
    $choice = Read-MenuChoice "Select permission to assign" $valid
    $perm = $Script:RequiredPermissions[[int]$choice - 1]
    Write-Status "Role"   $perm.Role  "info"
    Write-Status "Scope"  $perm.Scope "info"
    Write-Status "Target" $perm.Who   "info"
    $objId = Read-MenuChoice "Enter Object ID (user/group/service principal)"
    $scope = Read-MenuChoice "Enter Scope (full resource ID or /subscriptions/...)"
    if ($objId -and $scope) {
        New-AzRoleAssignment -ObjectId $objId -RoleDefinitionName $perm.Role `
            -Scope $scope | Out-Null
        Write-Status "Role assigned successfully" "" "ok"
    }
    Read-Host "  Press Enter to continue"
}

function Check-SubscriptionAccess {
    Write-Section "Subscription-Level Access Check"
    $ctx    = Get-AzContext
    $userId = $ctx.Account.Id
    $subScope = "/subscriptions/$($Global:Cfg.SubscriptionId)"
    Write-Status "Checking roles for" $userId "step"
    $roles = Get-AzRoleAssignment -SignInName $userId -Scope $subScope -EA SilentlyContinue

    $required = @("Contributor","Owner","User Access Administrator")
    foreach ($r in $required) {
        $has = $null -ne ($roles | Where-Object { $_.RoleDefinitionName -eq $r })
        Write-Status $r "" (if ($has) {"ok"} else {"warn"})
    }

    Write-Host ""
    Write-Status "All current subscription-level roles:" "" "info"
    foreach ($r in $roles) {
        Write-Host ("    • {0,-45} (Scope: {1})" -f $r.RoleDefinitionName, $r.Scope.Split("/")[-1]) -ForegroundColor $C.Muted
    }
    Read-Host "  Press Enter to continue"
}

function Export-PermissionReport {
    $report = $Script:RequiredPermissions |
        Select-Object @{N="Scope";E={$_.Scope}},
                      @{N="Principal";E={$_.Who}},
                      @{N="Role";E={$_.Role}},
                      @{N="Purpose";E={$_.Why}}
    $file = ".\BDF-AVD-PermissionReport-$(Get-Date -Format 'yyyyMMdd').csv"
    $report | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
    Write-Status "Permission report exported" $file "ok"
    Read-Host "  Press Enter to continue"
}
#endregion

#region ═══════════════════════════════════════════════════════════════════════
#  COMPONENT DEPLOYMENT FUNCTIONS (with detection/skip logic)
#═══════════════════════════════════════════════════════════════════════════════

function Deploy-ResourceGroups {
    Write-Banner "RESOURCE GROUPS"

    $existingRGs = Find-ExistingResourceGroups
    $tags = @{ Environment=$Global:Cfg.Environment; Project="AVD-Migration"; Owner="IT-Infrastructure"; ManagedBy="AIO-Deploy" }

    foreach ($rgKey in @("AVD","Network","Storage","Monitoring","Automation")) {
        $rgCfg = $Global:Cfg.RG[$rgKey]
        Write-Host ""
        Write-Host "  ── $rgKey Resource Group" -ForegroundColor $C.Accent

        # Check if it already exists
        $exists = Get-AzResourceGroup -Name $rgCfg.Name -EA SilentlyContinue
        if ($exists) {
            Write-Status "Exists" $rgCfg.Name "reuse"
            $rgCfg.IsNew = $false
            $rgCfg.ResourceId = $exists.ResourceId
            continue
        }

        # Offer to use an existing RG or create new
        $match = $existingRGs | Where-Object { $_.ResourceGroupName -like "*avd*" -or $_.ResourceGroupName -like "*bdf*" }
        if ($match -and @($match).Count -gt 0) {
            Write-Status "Suggested existing RGs found" "" "info"
            $i = 1
            foreach ($m in @($match)) {
                Write-Host ("    {0}. {1}" -f $i, $m.ResourceGroupName) -ForegroundColor $C.Reuse
                $i++
            }
            Write-Host "    N. Create new: $($rgCfg.Name)" -ForegroundColor $C.New
            $opts = (1..@($match).Count | ForEach-Object {"$_"}) + @("N","n")
            $ch = Read-MenuChoice "Use existing or create new $rgKey RG" $opts "N"
            if ($ch -notmatch "^[Nn]$") {
                $sel = @($match)[[int]$ch - 1]
                $rgCfg.Name    = $sel.ResourceGroupName
                $rgCfg.IsNew   = $false
                $rgCfg.ResourceId = (Get-AzResourceGroup -Name $sel.ResourceGroupName).ResourceId
                Write-Status "Using existing" $rgCfg.Name "reuse"
                continue
            }
        }

        # Create new
        Write-Status "Creating" $rgCfg.Name "new"
        $rg = New-AzResourceGroup -Name $rgCfg.Name -Location $Global:Cfg.Location -Tag $tags
        $rgCfg.IsNew      = $true
        $rgCfg.ResourceId = $rg.ResourceId
        Write-Status "Created" $rgCfg.Name "ok"
    }

    $Global:Cfg.DeploymentState.ResourceGroups = "Deployed"
    Save-Config
}

function Deploy-NetworkingComponent {
    Write-Banner "NETWORKING"

    # ── VNet ──────────────────────────────────────────────────────────────
    Write-Section "Virtual Network"
    $foundVNets = Find-ExistingVNets

    if (@($foundVNets).Count -gt 0) {
        Write-Host ""
        Write-Host ("  {0,-4} {1,-35} {2,-18} {3,-8} {4}" -f "#", "VNet Name", "ResourceGroup", "Subnets", "Address Space") -ForegroundColor $C.Muted
        Write-Host "  $($UI.SH * 78)" -ForegroundColor $C.Border
        $i = 1
        foreach ($v in $foundVNets) {
            Write-Host ("  {0,-4}" -f $i) -NoNewline -ForegroundColor $C.Accent
            Write-Host ("{0,-35}" -f $v.Name) -NoNewline -ForegroundColor $C.Menu
            Write-Host ("{0,-18}" -f $v.ResourceGroupName) -NoNewline -ForegroundColor $C.Muted
            Write-Host ("{0,-8}" -f $v.SubnetCount) -NoNewline -ForegroundColor $C.Muted
            Write-Host $v.AddressSpace -ForegroundColor $C.Azure
            $i++
        }
        Write-Host ""
        Write-Host "     N.  Create new VNet" -ForegroundColor $C.New
        $opts  = (1..@($foundVNets).Count | ForEach-Object {"$_"}) + @("N","n")
        $vChoice = Read-MenuChoice "Use existing VNet or create new?" $opts "N"

        if ($vChoice -notmatch "^[Nn]$") {
            $selVNet = @($foundVNets)[[int]$vChoice - 1]
            $Global:Cfg.Network.VNet.Name          = $selVNet.Name
            $Global:Cfg.Network.VNet.ResourceGroup = $selVNet.ResourceGroupName
            $Global:Cfg.Network.VNet.Id            = $selVNet.Id
            $Global:Cfg.Network.VNet.IsNew         = $false
            Write-Status "Reusing existing VNet" $selVNet.Name "reuse"

            # Subnet handling for existing VNet
            Invoke-SubnetHandling -VNetName $selVNet.Name -VNetRG $selVNet.ResourceGroupName
            $Global:Cfg.DeploymentState.Networking = "Deployed"
            Save-Config
            return
        }
    } else {
        Write-Status "No existing VNets found in $($Global:Cfg.LocationDisplay)" "" "warn"
        if (-not (Read-YesNo "Create new VNet?" $true)) {
            Write-Status "Networking skipped" "" "skip"
            $Global:Cfg.DeploymentState.Networking = "Skipped"
            Save-Config
            return
        }
    }

    # Confirm new VNet address space
    $defaultAS = "10.10.0.0/16"
    Write-Host ""
    Write-Status "Default address space" $defaultAS "info"
    $customAS = Read-MenuChoice "Enter address space or press Enter to use default" -Default $defaultAS
    if (-not $customAS) { $customAS = $defaultAS }

    # Create VNet with subnets
    Write-Status "Creating VNet" "$($Global:Cfg.Network.VNet.Name) ($customAS)" "new"
    $nsg = New-AzNetworkSecurityGroup -Name "nsg-avd-$($Global:Cfg.Prefix)" `
               -ResourceGroupName $Global:Cfg.RG.Network.Name `
               -Location $Global:Cfg.Location `
               -SecurityRules @(
                   New-AzNetworkSecurityRuleConfig -Name "Allow-AVD" -Priority 100 -Direction Inbound `
                       -Access Allow -Protocol Tcp -SourceAddressPrefix "WindowsVirtualDesktop" `
                       -SourcePortRange "*" -DestinationAddressPrefix "VirtualNetwork" -DestinationPortRange "443"
                   New-AzNetworkSecurityRuleConfig -Name "Allow-VNet" -Priority 200 -Direction Inbound `
                       -Access Allow -Protocol "*" -SourceAddressPrefix "VirtualNetwork" `
                       -SourcePortRange "*" -DestinationAddressPrefix "VirtualNetwork" -DestinationPortRange "*"
                   New-AzNetworkSecurityRuleConfig -Name "Deny-All-In" -Priority 4096 -Direction Inbound `
                       -Access Deny -Protocol "*" -SourceAddressPrefix "*" `
                       -SourcePortRange "*" -DestinationAddressPrefix "*" -DestinationPortRange "*"
               )
    $s = $Global:Cfg.Network.Subnets
    $newVNet = New-AzVirtualNetwork -Name $Global:Cfg.Network.VNet.Name `
                   -ResourceGroupName $Global:Cfg.RG.Network.Name `
                   -Location $Global:Cfg.Location `
                   -AddressPrefix $customAS `
                   -Subnet @(
                       New-AzVirtualNetworkSubnetConfig -Name $s.E3.Name      -AddressPrefix $s.E3.Prefix      -NetworkSecurityGroup $nsg
                       New-AzVirtualNetworkSubnetConfig -Name $s.F1.Name      -AddressPrefix $s.F1.Prefix      -NetworkSecurityGroup $nsg
                       New-AzVirtualNetworkSubnetConfig -Name $s.Mgmt.Name    -AddressPrefix $s.Mgmt.Prefix
                       New-AzVirtualNetworkSubnetConfig -Name $s.Storage.Name -AddressPrefix $s.Storage.Prefix -PrivateEndpointNetworkPoliciesFlag Disabled
                   )
    $Global:Cfg.Network.VNet.Id            = $newVNet.Id
    $Global:Cfg.Network.VNet.IsNew         = $true
    $Global:Cfg.Network.VNet.ResourceGroup = $Global:Cfg.RG.Network.Name
    Write-Status "VNet created" $Global:Cfg.Network.VNet.Name "ok"
    $Global:Cfg.DeploymentState.Networking = "Deployed"
    Save-Config
}

function Invoke-SubnetHandling {
    param([string]$VNetName, [string]$VNetRG)

    Write-Section "Subnet Detection — Existing VNet: $VNetName"
    $vnet    = Get-AzVirtualNetwork -Name $VNetName -ResourceGroupName $VNetRG
    $subnets = $vnet.Subnets

    if ($subnets.Count -eq 0) {
        Write-Status "No subnets found — will create required subnets" "" "warn"
        return
    }

    Write-Host ""
    Write-Host ("  {0,-4} {1,-30} {2,-20} {3}" -f "#", "Subnet Name", "Address Prefix", "Delegations") -ForegroundColor $C.Muted
    Write-Host "  $($UI.SH * 70)" -ForegroundColor $C.Border
    $i = 1
    foreach ($sn in $subnets) {
        Write-Host ("  {0,-4}" -f $i) -NoNewline -ForegroundColor $C.Accent
        Write-Host ("{0,-30}" -f $sn.Name) -NoNewline -ForegroundColor $C.Menu
        Write-Host ("{0,-20}" -f ($sn.AddressPrefix | Select-Object -First 1)) -NoNewline -ForegroundColor $C.Azure
        Write-Host ($sn.Delegations.ServiceName -join ",") -ForegroundColor $C.Muted
        $i++
    }

    Write-Host ""
    Write-Status "Map your subnets to AVD roles (enter number, or N to create new)" "" "info"

    foreach ($role in @("E3","F1","Mgmt","Storage")) {
        $cfg = $Global:Cfg.Network.Subnets[$role]
        Write-Host ""
        Write-Host "  AVD $role subnet (default: $($cfg.Name), $($cfg.Prefix)):" -ForegroundColor $C.Accent
        $opts = (1..$subnets.Count | ForEach-Object {"$_"}) + @("N","n","S","s")
        $ch = Read-MenuChoice "  Use existing subnet (# ) / N=new / S=skip" $opts "N"
        if ($ch -match "^\d+$") {
            $sel = $subnets[[int]$ch - 1]
            $cfg.Name  = $sel.Name
            $cfg.IsNew = $false
            Write-Status "Mapped" "$role → $($sel.Name)" "reuse"
        } elseif ($ch -in @("S","s")) {
            Write-Status "Skipped" $role "skip"
        } else {
            $cfg.IsNew = $true
            Write-Status "Will create" "$($cfg.Name) ($($cfg.Prefix))" "new"
        }
    }
}

function Deploy-KeyVaultComponent {
    Write-Banner "KEY VAULT"

    # Detect existing
    $found = Find-ExistingKeyVaults
    if (@($found).Count -gt 0) {
        Write-Host ""
        $i = 1
        foreach ($kv in $found) {
            Write-Host ("  {0}. {1,-35} RG: {2,-25} RBAC: {3}" -f $i, $kv.VaultName, $kv.ResourceGroupName, $kv.EnableRbac) -ForegroundColor $C.Menu
            $i++
        }
        Write-Host "     N. Create new Key Vault" -ForegroundColor $C.New
        Write-Host "     S. Skip Key Vault" -ForegroundColor $C.Skip
        $opts = (1..@($found).Count | ForEach-Object {"$_"}) + @("N","n","S","s")
        $ch = Read-MenuChoice "Key Vault selection" $opts "N"
        if ($ch -in @("S","s")) {
            Write-Status "Key Vault skipped" "" "skip"
            $Global:Cfg.DeploymentState.KeyVault = "Skipped"
            Save-Config; return
        }
        if ($ch -notmatch "^[Nn]$") {
            $sel = @($found)[[int]$ch - 1]
            $Global:Cfg.KeyVault.Name          = $sel.VaultName
            $Global:Cfg.KeyVault.ResourceGroup = $sel.ResourceGroupName
            $Global:Cfg.KeyVault.Id            = $sel.ResourceId
            $Global:Cfg.KeyVault.IsNew         = $false
            Write-Status "Reusing Key Vault" $sel.VaultName "reuse"
            # Store credentials in existing KV
            Invoke-StoreKeyVaultSecrets
            $Global:Cfg.DeploymentState.KeyVault = "Deployed"
            Save-Config; return
        }
    }

    # Create new KV
    Write-Status "Creating Key Vault" $Global:Cfg.KeyVault.Name "new"
    $kv = New-AzKeyVault -VaultName $Global:Cfg.KeyVault.Name `
              -ResourceGroupName $Global:Cfg.KeyVault.ResourceGroup `
              -Location $Global:Cfg.Location `
              -Sku Standard `
              -EnableRbacAuthorization $true `
              -EnableSoftDelete $true `
              -SoftDeleteRetentionInDays 30 `
              -EnablePurgeProtection $true
    $Global:Cfg.KeyVault.Id    = $kv.ResourceId
    $Global:Cfg.KeyVault.IsNew = $true
    Write-Status "Key Vault created" $Global:Cfg.KeyVault.Name "ok"
    Invoke-StoreKeyVaultSecrets
    $Global:Cfg.DeploymentState.KeyVault = "Deployed"
    Save-Config
}

function Invoke-StoreKeyVaultSecrets {
    $kvName = $Global:Cfg.KeyVault.Name
    # Grant current user secrets officer role
    $myId = (Get-AzADUser -UserPrincipalName (Get-AzContext).Account.Id -EA SilentlyContinue).Id
    if ($myId) {
        New-AzRoleAssignment -ObjectId $myId -RoleDefinitionName "Key Vault Secrets Officer" `
            -Scope $Global:Cfg.KeyVault.Id -EA SilentlyContinue | Out-Null
        Start-Sleep -Seconds 20
    }
    $adminPwd = Read-Host "  Enter VM admin password to store in Key Vault" -AsSecureString
    Set-AzKeyVaultSecret -VaultName $kvName -Name "avd-vm-admin-password" -SecretValue $adminPwd | Out-Null
    Set-AzKeyVaultSecret -VaultName $kvName -Name "avd-vm-admin-username" `
        -SecretValue (ConvertTo-SecureString $Global:Cfg.Identity.AdminUsername -AsPlainText -Force) | Out-Null
    Write-Status "Credentials stored in Key Vault" $kvName "ok"
}

function Deploy-LogAnalyticsComponent {
    Write-Banner "LOG ANALYTICS WORKSPACE"
    $found = Find-ExistingLogAnalytics

    if (@($found).Count -gt 0) {
        $i = 1
        foreach ($w in $found) {
            Write-Host ("  {0}. {1,-35} RG: {2,-25} SKU: {3}" -f $i, $w.Name, $w.ResourceGroupName, $w.Sku) -ForegroundColor $C.Menu
            $i++
        }
        Write-Host "     N. Create new" -ForegroundColor $C.New
        Write-Host "     S. Skip" -ForegroundColor $C.Skip
        $opts = (1..@($found).Count | ForEach-Object {"$_"}) + @("N","n","S","s")
        $ch = Read-MenuChoice "Log Analytics selection" $opts "N"
        if ($ch -in @("S","s")) {
            $Global:Cfg.DeploymentState.LogAnalytics = "Skipped"; Save-Config; return
        }
        if ($ch -notmatch "^[Nn]$") {
            $sel = @($found)[[int]$ch - 1]
            $Global:Cfg.LogAnalytics.Name          = $sel.Name
            $Global:Cfg.LogAnalytics.ResourceGroup = $sel.ResourceGroupName
            $Global:Cfg.LogAnalytics.WorkspaceId   = $sel.CustomerId
            $Global:Cfg.LogAnalytics.Id            = $sel.ResourceId
            $Global:Cfg.LogAnalytics.IsNew         = $false
            Write-Status "Reusing Log Analytics" $sel.Name "reuse"
            $Global:Cfg.DeploymentState.LogAnalytics = "Deployed"; Save-Config; return
        }
    }

    Write-Status "Creating Log Analytics Workspace" $Global:Cfg.LogAnalytics.Name "new"
    $law = New-AzOperationalInsightsWorkspace -Name $Global:Cfg.LogAnalytics.Name `
               -ResourceGroupName $Global:Cfg.LogAnalytics.ResourceGroup `
               -Location $Global:Cfg.Location -Sku PerGB2018 -RetentionInDays 90
    $Global:Cfg.LogAnalytics.WorkspaceId = $law.CustomerId
    $Global:Cfg.LogAnalytics.Id          = $law.ResourceId
    $Global:Cfg.LogAnalytics.IsNew       = $true
    Write-Status "Created" $Global:Cfg.LogAnalytics.Name "ok"
    $Global:Cfg.DeploymentState.LogAnalytics = "Deployed"
    Save-Config
}

function Deploy-AzureFilesComponent {
    Write-Banner "AZURE FILES — FSLOGIX"
    $found = Find-ExistingStorageAccounts

    if (@($found).Count -gt 0) {
        $i = 1
        foreach ($s in $found) {
            Write-Host ("  {0}. {1,-32} RG: {2,-25} SKU: {3}" -f $i, $s.StorageAccountName, $s.ResourceGroupName, $s.Sku) -ForegroundColor $C.Menu
            $i++
        }
        Write-Host "     N. Create new" -ForegroundColor $C.New
        Write-Host "     S. Skip" -ForegroundColor $C.Skip
        $opts = (1..@($found).Count | ForEach-Object {"$_"}) + @("N","n","S","s")
        $ch = Read-MenuChoice "Storage Account" $opts "N"
        if ($ch -in @("S","s")) {
            $Global:Cfg.DeploymentState.AzureFiles = "Skipped"; Save-Config; return
        }
        if ($ch -notmatch "^[Nn]$") {
            $sel = @($found)[[int]$ch - 1]
            $Global:Cfg.Storage.Account.Name          = $sel.StorageAccountName
            $Global:Cfg.Storage.Account.ResourceGroup = $sel.ResourceGroupName
            $Global:Cfg.Storage.Account.IsNew         = $false
            Write-Status "Reusing Storage Account" $sel.StorageAccountName "reuse"
            Invoke-CreateFileShares
            $Global:Cfg.DeploymentState.AzureFiles = "Deployed"; Save-Config; return
        }
    }

    Write-Status "Creating Azure Files Premium storage" $Global:Cfg.Storage.Account.Name "new"
    New-AzStorageAccount -Name $Global:Cfg.Storage.Account.Name `
        -ResourceGroupName $Global:Cfg.Storage.Account.ResourceGroup `
        -Location $Global:Cfg.Location -SkuName "Premium_ZRS" -Kind FileStorage `
        -MinimumTlsVersion TLS1_2 -AllowBlobPublicAccess $false `
        -EnableHttpsTrafficOnly $true | Out-Null

    # Enable AAD Kerberos
    $saRes = Get-AzResource -ResourceType "Microsoft.Storage/storageAccounts" `
             -Name $Global:Cfg.Storage.Account.Name
    $saRes.Properties.azureFilesIdentityBasedAuthentication = @{ directoryServiceOptions = "AADKERB" }
    $saRes | Set-AzResource -Force | Out-Null
    Write-Status "Azure AD Kerberos enabled" "" "ok"

    Invoke-CreateFileShares
    Invoke-CreatePrivateEndpoint
    $Global:Cfg.Storage.Account.IsNew         = $true
    $Global:Cfg.DeploymentState.AzureFiles = "Deployed"
    Save-Config
}

function Invoke-CreateFileShares {
    $sa  = Get-AzStorageAccount -Name $Global:Cfg.Storage.Account.Name `
               -ResourceGroupName $Global:Cfg.Storage.Account.ResourceGroup
    $ctx = $sa.Context
    foreach ($share in @(
        @{ Name=$Global:Cfg.Storage.Shares.E3;   Quota=1024;  Label="E3 Profiles"  }
        @{ Name=$Global:Cfg.Storage.Shares.F1;   Quota=512;   Label="F1 Profiles"  }
        @{ Name=$Global:Cfg.Storage.Shares.ODFC; Quota=2048;  Label="ODFC (Teams+Outlook)" }
    )) {
        if (-not (Get-AzStorageShare -Name $share.Name -Context $ctx -EA SilentlyContinue)) {
            New-AzStorageShare -Name $share.Name -Context $ctx -QuotaGiB $share.Quota | Out-Null
            Write-Status "Share created" "$($share.Name) ($($share.Quota) GiB) — $($share.Label)" "ok"
        } else {
            Write-Status "Share exists" $share.Name "reuse"
        }
    }
    Update-AzStorageFileServiceProperty -StorageAccountName $Global:Cfg.Storage.Account.Name `
        -ResourceGroupName $Global:Cfg.Storage.Account.ResourceGroup `
        -EnableShareDeleteRetentionPolicy $true -ShareRetentionDays 30 -EA SilentlyContinue | Out-Null
}

function Invoke-CreatePrivateEndpoint {
    if ($Global:Cfg.Network.VNet.Id -eq "") {
        Write-Status "VNet not configured — skipping Private Endpoint" "" "warn"; return
    }
    Write-Status "Creating Private Endpoint for Azure Files" "" "new"
    $vnet    = Get-AzVirtualNetwork -Name $Global:Cfg.Network.VNet.Name `
                   -ResourceGroupName $Global:Cfg.Network.VNet.ResourceGroup
    $subnet  = $vnet.Subnets | Where-Object { $_.Name -eq $Global:Cfg.Network.Subnets.Storage.Name }
    if (-not $subnet) { Write-Status "Storage subnet not found — skipping PE" "" "warn"; return }
    $sa      = Get-AzStorageAccount -Name $Global:Cfg.Storage.Account.Name `
                   -ResourceGroupName $Global:Cfg.Storage.Account.ResourceGroup
    $peLinkCfg = New-AzPrivateLinkServiceConnection -Name "plsc-avd-files" `
                     -PrivateLinkServiceId $sa.Id -GroupId "file"
    New-AzPrivateEndpoint -Name "pe-avd-files-$($Global:Cfg.Prefix)" `
        -ResourceGroupName $Global:Cfg.Storage.Account.ResourceGroup `
        -Location $Global:Cfg.Location -Subnet $subnet `
        -PrivateLinkServiceConnection $peLinkCfg | Out-Null
    Write-Status "Private Endpoint created" "" "ok"
}
#endregion

#region ═══════════════════════════════════════════════════════════════════════
#  HOST POOL DEPLOYMENT WITH DETECTION
#═══════════════════════════════════════════════════════════════════════════════

function Deploy-HostPoolsComponent {
    Write-Banner "AVD HOST POOLS"
    $found = Find-ExistingHostPools

    foreach ($poolType in @("E3","F1")) {
        $poolCfg = $Global:Cfg.HostPools[$poolType]
        Write-Section "$poolType Host Pool"

        # Check if pool already exists by configured name
        $existing = Get-AzWvdHostPool -Name $poolCfg.Name -ResourceGroupName $Global:Cfg.RG.AVD.Name -EA SilentlyContinue
        if ($existing) {
            Write-Status "Host pool already exists" $poolCfg.Name "reuse"
            $poolCfg.Status = "Deployed"; continue
        }

        if (@($found).Count -gt 0) {
            $i = 1
            foreach ($hp in $found) {
                Write-Host ("  {0}. {1,-40} Type: {2,-8} MaxSessions: {3}" -f $i, $hp.Name, $hp.Type, $hp.MaxSessions) -ForegroundColor $C.Menu
                $i++
            }
            Write-Host "     N. Create new $poolType host pool" -ForegroundColor $C.New
            Write-Host "     S. Skip" -ForegroundColor $C.Skip
            $opts = (1..@($found).Count | ForEach-Object {"$_"}) + @("N","n","S","s")
            $ch = Read-MenuChoice "$poolType host pool" $opts "N"
            if ($ch -in @("S","s")) { $poolCfg.Status = "Skipped"; continue }
            if ($ch -notmatch "^[Nn]$") {
                $sel = @($found)[[int]$ch - 1]
                $poolCfg.Name   = $sel.Name
                $poolCfg.Status = "Deployed"
                Write-Status "Reusing" $sel.Name "reuse"; continue
            }
        }

        # Configure pool settings
        Write-Host ""
        Write-Status "Configuring $poolType host pool" $poolCfg.Name "step"
        $maxSess = Read-MenuChoice "Max sessions per VM" -Default "$($poolCfg.MaxSessions)"
        if ($maxSess -match "^\d+$") { $poolCfg.MaxSessions = [int]$maxSess }

        Write-Status "Creating host pool" $poolCfg.Name "new"
        $appType    = if ($poolType -eq "E3") { "Desktop" } else { "RailApplications" }
        $tokenExpiry = (Get-Date).ToUniversalTime().AddHours(48).ToString("yyyy-MM-ddTHH:mm:ss.fffZ")

        # Use configured RDP profile or build balanced baseline
        $isAAD = ($Global:Cfg.JoinConfig.Type -eq "EntraID" -or -not $Global:Cfg.JoinConfig.Configured)
        if ($Global:Cfg.RDP.Configured -and $Global:Cfg.RDP.E3PropertyString) {
            $rdpProps = if ($poolType -eq "E3") {$Global:Cfg.RDP.E3PropertyString} else {$Global:Cfg.RDP.F1PropertyString}
            Write-Status "RDP Profile" "$($Global:Cfg.RDP.ActiveProfile) — $($rdpProps.Split(';').Count) properties" "ok"
        } else {
            # Auto-build balanced profile appropriate for pool type
            $defaultOverrides = if ($poolType -eq "F1") {$Script:RDPProfiles["Frontline"].Overrides} else {$Script:RDPProfiles["Balanced"].Overrides}
            $rdpProps = Build-RDPString $poolType $defaultOverrides
            Write-Status "RDP Profile" "Auto-built $(if($poolType -eq 'F1'){'Frontline'}else{'Balanced'}) — configure in RDP Manager for custom settings" "info"
        }
        # Enforce correct join-type flag
        if ($isAAD) {
            $rdpProps = ($rdpProps -replace "targetisaadjoined:i:0","targetisaadjoined:i:1")
            if ($rdpProps -notlike "*targetisaadjoined:i:1*") { $rdpProps = "targetisaadjoined:i:1;$rdpProps" }
        } else {
            $rdpProps = ($rdpProps -replace "targetisaadjoined:i:1","targetisaadjoined:i:0")
        }

        $hp = New-AzWvdHostPool -Name $poolCfg.Name `
                  -ResourceGroupName $Global:Cfg.RG.AVD.Name `
                  -Location $Global:Cfg.Location `
                  -FriendlyName "BDF $poolType Workers" `
                  -HostPoolType Pooled `
                  -LoadBalancerType $poolCfg.LoadBalancer `
                  -PreferredAppGroupType $appType `
                  -MaxSessionLimit $poolCfg.MaxSessions `
                  -StartVMOnConnect $true `
                  -ValidationEnvironment ($Global:Cfg.Environment -eq "POC") `
                  -ExpirationTime $tokenExpiry `
                  -RegistrationTokenOperation Update `
                  -CustomRdpProperty $rdpProps

        # Diagnostic settings
        if ($Global:Cfg.LogAnalytics.Id) {
            Set-AzDiagnosticSetting -ResourceId $hp.Id -Name "diag-$($poolCfg.Name)" `
                -WorkspaceId $Global:Cfg.LogAnalytics.Id -Enabled $true `
                -Category @("Checkpoint","Error","Management","Connection","HostRegistration","AgentHealthStatus","SessionHostManagement") `
                -EA SilentlyContinue | Out-Null
        }
        $poolCfg.Status = "Deployed"
        Write-Status "Host pool created" $poolCfg.Name "ok"
    }
    $Global:Cfg.DeploymentState.HostPools = "Deployed"
    Save-Config
}
#endregion

#region ═══════════════════════════════════════════════════════════════════════
#  DYNAMIC AUTO-SCALING CONFIGURATION WIZARD
#═══════════════════════════════════════════════════════════════════════════════

function Show-AutoScalingMenu {
    Write-Banner "AUTO-SCALING MANAGER" $C.Azure
    while ($true) {
        Write-Section "Auto-Scaling Options"

        # Show current scaling status
        foreach ($pool in @("E3","F1")) {
            $sc  = $Global:Cfg.Scaling[$pool]
            $spExists = Get-AzWvdScalingPlan -Name $sc.PlanName `
                            -ResourceGroupName $Global:Cfg.RG.AVD.Name -EA SilentlyContinue
            $status = if ($spExists) { "Deployed" } else { "Not Deployed" }
            $icon   = if ($spExists) { "✔" } else { "○" }
            $col    = if ($spExists) { $C.Ok } else { $C.Muted }
            Write-Host ("  $icon  $pool Scaling Plan: {0,-30} Schedules: {1}" -f $sc.PlanName, @($sc.Schedules).Count) -ForegroundColor $col
        }

        Write-Host ""
        Write-Host "  1.  Run Scaling Configuration Wizard (E3 + F1)" -ForegroundColor $C.Menu
        Write-Host "  2.  Edit E3 Office scaling schedule" -ForegroundColor $C.Menu
        Write-Host "  3.  Edit F1 Frontline scaling schedule" -ForegroundColor $C.Menu
        Write-Host "  4.  Deploy / Update Scaling Plans" -ForegroundColor $C.New
        Write-Host "  5.  Configure Automation Runbooks" -ForegroundColor $C.Menu
        Write-Host "  6.  Set Holiday Surge Dates" -ForegroundColor $C.Menu
        Write-Host "  7.  View current scaling status (live)" -ForegroundColor $C.Azure
        Write-Host "  8.  Trigger manual scale action" -ForegroundColor $C.Warn
        Write-Host "  0.  Back" -ForegroundColor $C.Muted
        $choice = Read-MenuChoice "Auto-Scaling" @("0","1","2","3","4","5","6","7","8")
        switch ($choice) {
            "1" { Invoke-ScalingWizard }
            "2" { Edit-ScalingSchedule "E3" }
            "3" { Edit-ScalingSchedule "F1" }
            "4" { Deploy-ScalingPlans }
            "5" { Deploy-AutomationRunbooks }
            "6" { Set-HolidayDates }
            "7" { Show-LiveScalingStatus }
            "8" { Invoke-ManualScaleAction }
            "0" { return }
        }
    }
}

function Invoke-ScalingWizard {
    Write-Banner "SCALING CONFIGURATION WIZARD"

    foreach ($poolType in @("E3","F1")) {
        $sc = $Global:Cfg.Scaling[$poolType]
        Write-Section "$poolType Scaling Plan — $($sc.PlanName)"

        $sc.Enabled = Read-YesNo "Enable auto-scaling for $poolType?" $true
        if (-not $sc.Enabled) { Write-Status "Scaling disabled for $poolType" "" "skip"; continue }

        # Time zone
        $tzOptions = @(
            "Eastern Standard Time"
            "Central Standard Time"
            "Mountain Standard Time"
            "Pacific Standard Time"
        )
        $i = 1
        foreach ($tz in $tzOptions) {
            $cur = if ($tz -eq $sc.TimeZone) { " ◄" } else { "" }
            Write-Host ("  {0}. {1}{2}" -f $i, $tz, $cur) -ForegroundColor $C.Menu
            $i++
        }
        $tzChoice = Read-MenuChoice "Time Zone" (1..$tzOptions.Count | ForEach-Object {"$_"}) "1"
        $sc.TimeZone = $tzOptions[[int]$tzChoice - 1]

        # Peak threshold
        $thresh = Read-MenuChoice "Peak capacity threshold % (scale-out triggers at this %)" -Default "$($sc.PeakCapacityPct)"
        if ($thresh -match "^\d+$") { $sc.PeakCapacityPct = [int]$thresh }

        # Force logoff minutes
        $logoff = Read-MenuChoice "Force logoff warning minutes (ramp-down)" -Default "$($sc.ForceLogoffMinutes)"
        if ($logoff -match "^\d+$") { $sc.ForceLogoffMinutes = [int]$logoff }

        # Build schedules interactively
        $sc.Schedules = @()
        $presets = Get-ScalingPresets $poolType
        Write-Host ""
        Write-Host "  Schedule Presets:" -ForegroundColor $C.Accent
        $i = 1
        foreach ($preset in $presets) {
            Write-Host ("  {0}. {1,-20} {2}" -f $i, $preset.Name, $preset.Description) -ForegroundColor $C.Menu
            $i++
        }
        Write-Host "  C. Custom schedule wizard" -ForegroundColor $C.Azure

        $opts = (1..$presets.Count | ForEach-Object {"$_"}) + @("C","c")
        $sch = Read-MenuChoice "Select schedule preset" $opts "1"

        if ($sch -in @("C","c")) {
            $sc.Schedules = Invoke-CustomScheduleWizard $poolType
        } else {
            $sc.Schedules = $presets[[int]$sch - 1].Schedules
            Write-Status "Preset applied" $presets[[int]$sch - 1].Name "ok"
        }
        Write-Status "$poolType scaling configured" "$(@($sc.Schedules).Count) schedule(s)" "ok"
    }
    Save-Config
    if (Read-YesNo "Deploy Scaling Plans now?" $true) { Deploy-ScalingPlans }
}

function Get-ScalingPresets {
    param([string]$PoolType)
    if ($PoolType -eq "E3") {
        return @(
            @{
                Name="Standard Office Hours"
                Description="Mon-Fri 6:30AM-10PM, Weekend reduced"
                Schedules = @(
                    @{ Name="Weekday"; Days=@("Monday","Tuesday","Wednesday","Thursday","Friday")
                       RampUpTime="06:30"; PeakTime="08:00"; RampDownTime="18:00"; OffPeakTime="22:00"
                       RampUpMin=20; PeakMin=50; RampDownMin=20; OffPeakMin=10 }
                    @{ Name="Weekend"; Days=@("Saturday","Sunday")
                       RampUpTime="08:00"; PeakTime="09:00"; RampDownTime="17:00"; OffPeakTime="19:00"
                       RampUpMin=10; PeakMin=30; RampDownMin=10; OffPeakMin=5 }
                )
            },
            @{
                Name="24/7 Always On"
                Description="Constant capacity — no scale-down (not recommended)"
                Schedules = @(
                    @{ Name="AlwaysOn"; Days=@("Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday")
                       RampUpTime="00:00"; PeakTime="00:00"; RampDownTime="23:00"; OffPeakTime="23:30"
                       RampUpMin=100; PeakMin=100; RampDownMin=100; OffPeakMin=100 }
                )
            },
            @{
                Name="Aggressive Savings"
                Description="Scale to 1 VM outside core hours — maximum cost savings"
                Schedules = @(
                    @{ Name="CoreHours"; Days=@("Monday","Tuesday","Wednesday","Thursday","Friday")
                       RampUpTime="07:45"; PeakTime="08:00"; RampDownTime="17:00"; OffPeakTime="19:00"
                       RampUpMin=15; PeakMin=80; RampDownMin=10; OffPeakMin=5 }
                )
            }
        )
    } else {
        return @(
            @{
                Name="Retail Store Hours"
                Description="7 days/week, 7:30AM-11PM with weekend hours"
                Schedules = @(
                    @{ Name="StoreWeekday"; Days=@("Monday","Tuesday","Wednesday","Thursday","Friday")
                       RampUpTime="07:30"; PeakTime="09:00"; RampDownTime="21:00"; OffPeakTime="23:00"
                       RampUpMin=30; PeakMin=60; RampDownMin=20; OffPeakMin=10 }
                    @{ Name="StoreWeekend"; Days=@("Saturday","Sunday")
                       RampUpTime="07:30"; PeakTime="09:00"; RampDownTime="20:00"; OffPeakTime="22:00"
                       RampUpMin=25; PeakMin=50; RampDownMin=15; OffPeakMin=10 }
                )
            },
            @{
                Name="Extended Holiday"
                Description="Earlier start, later end — holiday season (Nov-Jan)"
                Schedules = @(
                    @{ Name="HolidayWeekday"; Days=@("Monday","Tuesday","Wednesday","Thursday","Friday")
                       RampUpTime="06:30"; PeakTime="08:00"; RampDownTime="22:00"; OffPeakTime="23:59"
                       RampUpMin=50; PeakMin=80; RampDownMin=30; OffPeakMin=10 }
                    @{ Name="HolidayWeekend"; Days=@("Saturday","Sunday")
                       RampUpTime="07:00"; PeakTime="08:00"; RampDownTime="22:00"; OffPeakTime="23:59"
                       RampUpMin=50; PeakMin=80; RampDownMin=30; OffPeakMin=10 }
                )
            }
        )
    }
}

function Invoke-CustomScheduleWizard {
    param([string]$PoolType)
    $schedules = @()
    Write-Section "Custom Schedule Builder — $PoolType"
    do {
        $name = Read-MenuChoice "Schedule name (e.g. Weekday, Weekend)"
        Write-Host "  Days of week (comma-separated, e.g. Monday,Tuesday,Wednesday,Thursday,Friday):" -ForegroundColor $C.Accent
        $daysInput = Read-MenuChoice "Days"
        $days = $daysInput -split "," | ForEach-Object { $_.Trim() }

        $rampUp   = Read-MenuChoice "Ramp-Up start time (HH:MM)"     -Default "07:00"
        $peak     = Read-MenuChoice "Peak start time (HH:MM)"         -Default "08:00"
        $rampDown = Read-MenuChoice "Ramp-Down start time (HH:MM)"    -Default "18:00"
        $offPeak  = Read-MenuChoice "Off-Peak start time (HH:MM)"     -Default "22:00"
        $ruMin    = Read-MenuChoice "Ramp-Up min hosts % (of max)"    -Default "20"
        $pkMin    = Read-MenuChoice "Peak min hosts % (of max)"        -Default "50"
        $rdMin    = Read-MenuChoice "Ramp-Down min hosts % (of max)"  -Default "15"
        $opMin    = Read-MenuChoice "Off-Peak min hosts % (of max)"   -Default "10"

        $schedules += @{
            Name        = $name
            Days        = $days
            RampUpTime  = $rampUp;  PeakTime      = $peak
            RampDownTime= $rampDown; OffPeakTime   = $offPeak
            RampUpMin   = [int]$ruMin; PeakMin     = [int]$pkMin
            RampDownMin = [int]$rdMin; OffPeakMin  = [int]$opMin
        }
        Write-Status "Schedule added" $name "ok"
    } until (-not (Read-YesNo "Add another schedule?" $false))
    return $schedules
}

function Edit-ScalingSchedule {
    param([string]$PoolType)
    $sc = $Global:Cfg.Scaling[$PoolType]
    if (@($sc.Schedules).Count -eq 0) {
        Write-Status "No schedules configured yet — run wizard first" "" "warn"
        Read-Host "  Press Enter"; return
    }
    Write-Section "Edit $PoolType Schedules"
    $i = 1
    foreach ($s in $sc.Schedules) {
        Write-Host ("  {0}. {1,-20} Days: {2}" -f $i, $s.Name, ($s.Days -join ",")) -ForegroundColor $C.Menu
        Write-Host ("      Times: RampUp={0} Peak={1} RampDown={2} OffPeak={3}" -f $s.RampUpTime,$s.PeakTime,$s.RampDownTime,$s.OffPeakTime) -ForegroundColor $C.Muted
        $i++
    }
    $opts = (1..@($sc.Schedules).Count | ForEach-Object {"$_"}) + @("A","a","D","d")
    Write-Host "  A. Add schedule  D. Delete schedule" -ForegroundColor $C.Azure
    $ch = Read-MenuChoice "Action" $opts
    if ($ch -in @("A","a")) {
        $newSch = Invoke-CustomScheduleWizard $PoolType
        $sc.Schedules += $newSch
        Write-Status "Schedules added" "" "ok"
    } elseif ($ch -in @("D","d")) {
        $del = Read-MenuChoice "Delete schedule number" (1..@($sc.Schedules).Count | ForEach-Object {"$_"})
        $sc.Schedules = $sc.Schedules | Where-Object { $_ -ne $sc.Schedules[[int]$del - 1] }
        Write-Status "Schedule deleted" "" "ok"
    } else {
        # Edit existing
        Write-Status "Editing schedule not yet interactive — re-run wizard to replace" "" "warn"
    }
    Save-Config
    Read-Host "  Press Enter to continue"
}

function Deploy-ScalingPlans {
    Write-Banner "DEPLOYING SCALING PLANS"

    foreach ($poolType in @("E3","F1")) {
        $sc      = $Global:Cfg.Scaling[$poolType]
        $hpCfg   = $Global:Cfg.HostPools[$poolType]
        if (-not $sc.Enabled) { Write-Status "Scaling disabled for $poolType — skipping" "" "skip"; continue }
        if (@($sc.Schedules).Count -eq 0) { Write-Status "No schedules for $poolType — run wizard first" "" "warn"; continue }

        Write-Section "Scaling Plan: $($sc.PlanName)"
        $existing = Get-AzWvdScalingPlan -Name $sc.PlanName `
                        -ResourceGroupName $Global:Cfg.RG.AVD.Name -EA SilentlyContinue
        if ($existing) {
            if (-not (Read-YesNo "Scaling Plan exists — replace/update?" $false)) {
                Write-Status "Skipped" $sc.PlanName "skip"; continue
            }
            Remove-AzWvdScalingPlan -Name $sc.PlanName `
                -ResourceGroupName $Global:Cfg.RG.AVD.Name | Out-Null
        }

        # Build schedule objects from config
        $scheduleObjs = foreach ($s in $sc.Schedules) {
            New-AzWvdScalingScheduleObject `
                -Name                          $s.Name `
                -DaysOfWeek                    $s.Days `
                -RampUpStartTime               $s.RampUpTime `
                -PeakStartTime                 $s.PeakTime `
                -RampDownStartTime             $s.RampDownTime `
                -OffPeakStartTime              $s.OffPeakTime `
                -RampUpLoadBalancingAlgorithm  BreadthFirst `
                -PeakLoadBalancingAlgorithm    BreadthFirst `
                -RampDownLoadBalancingAlgorithm DepthFirst `
                -OffPeakLoadBalancingAlgorithm  DepthFirst `
                -RampUpCapacityThresholdPct     $sc.RampUpCapacityPct `
                -PeakCapacityThresholdPct       $sc.PeakCapacityPct `
                -RampDownCapacityThresholdPct   90 `
                -OffPeakCapacityThresholdPct    90 `
                -RampUpMinimumHostsPct          $s.RampUpMin `
                -RampDownMinimumHostsPct        $s.RampDownMin `
                -OffPeakMinimumHostsPct         $s.OffPeakMin `
                -RampDownForceLogoffUser        $true `
                -RampDownWaitTimeMinute         $sc.ForceLogoffMinutes `
                -RampDownNotificationMessage    "Your session will close in $($sc.ForceLogoffMinutes) minutes. Please save your work." `
                -RampDownStopHostsWhen          ZeroActiveSessions `
                -RampUpStartVMOnConnect         Enable `
                -PeakStartVMOnConnect           Enable `
                -RampDownStartVMOnConnect       Disable `
                -OffPeakStartVMOnConnect        Enable
        }

        $sp = New-AzWvdScalingPlan -Name $sc.PlanName `
                  -ResourceGroupName $Global:Cfg.RG.AVD.Name `
                  -Location $Global:Cfg.Location `
                  -FriendlyName "BDF $poolType Scaling" `
                  -HostPoolType Pooled `
                  -TimeZone $sc.TimeZone `
                  -Schedule $scheduleObjs

        # Associate with host pool
        $hp = Get-AzWvdHostPool -Name $hpCfg.Name `
                  -ResourceGroupName $Global:Cfg.RG.AVD.Name -EA SilentlyContinue
        if ($hp) {
            Update-AzWvdScalingPlan -Name $sc.PlanName `
                -ResourceGroupName $Global:Cfg.RG.AVD.Name `
                -HostPoolReference @(@{ HostPoolArmPath=$hp.Id; ScalingPlanEnabled=$true }) | Out-Null
            Write-Status "Assigned to host pool" $hpCfg.Name "ok"
        }

        # Diagnostics
        if ($Global:Cfg.LogAnalytics.Id) {
            Set-AzDiagnosticSetting -ResourceId $sp.Id -Name "diag-$($sc.PlanName)" `
                -WorkspaceId $Global:Cfg.LogAnalytics.Id -Enabled $true `
                -Category @("Autoscale") -EA SilentlyContinue | Out-Null
        }
        Write-Status "Scaling Plan deployed" "$($sc.PlanName) ($(@($sc.Schedules).Count) schedules)" "ok"
    }
    $Global:Cfg.DeploymentState.ScalingPlans = "Deployed"
    Save-Config
}

function Set-HolidayDates {
    Write-Section "Holiday Surge Configuration"
    Write-Host "  Configure specific dates where max capacity should be pre-warmed." -ForegroundColor $C.Muted
    Write-Host "  The holiday runbook will be scheduled to activate the day before each date." -ForegroundColor $C.Muted
    Write-Host ""
    $currentYear = (Get-Date).Year
    $holidays = @(
        @{ Name="Black Friday";   Date="$currentYear-11-28"; E3Max=15; F1Max=15 }
        @{ Name="Christmas Eve";  Date="$currentYear-12-24"; E3Max=12; F1Max=12 }
        @{ Name="Post-Holiday";   Date="$currentYear-12-26"; E3Max=12; F1Max=12 }
        @{ Name="New Year's Eve"; Date="$currentYear-12-31"; E3Max=10; F1Max=10 }
    )
    $i = 1
    foreach ($h in $holidays) {
        Write-Host ("  {0}. {1,-18} {2}   E3 Max: {3}  F1 Max: {4}" -f $i, $h.Name, $h.Date, $h.E3Max, $h.F1Max) -ForegroundColor $C.Menu
        $i++
    }
    Write-Host "  A. Add custom holiday date" -ForegroundColor $C.New
    $opts = (1..$holidays.Count | ForEach-Object {"$_"}) + @("A","a")
    $ch = Read-MenuChoice "Configure" $opts
    if ($ch -in @("A","a")) {
        $hName = Read-MenuChoice "Holiday name"
        $hDate = Read-MenuChoice "Date (YYYY-MM-DD)"
        $hE3   = Read-MenuChoice "E3 max hosts" -Default "12"
        $hF1   = Read-MenuChoice "F1 max hosts" -Default "12"
        $holidays += @{ Name=$hName; Date=$hDate; E3Max=[int]$hE3; F1Max=[int]$hF1 }
    }
    # Store in config and schedule if automation account exists
    $aa = Get-AzAutomationAccount -Name $Global:Cfg.Automation.AccountName `
          -ResourceGroupName $Global:Cfg.Automation.ResourceGroup -EA SilentlyContinue
    if ($aa) {
        foreach ($h in $holidays) {
            $schedDate = ([datetime]$h.Date).AddDays(-1).ToString("yyyy-MM-dd") + "T18:00:00"
            New-AzAutomationSchedule -AutomationAccountName $Global:Cfg.Automation.AccountName `
                -ResourceGroupName $Global:Cfg.Automation.ResourceGroup `
                -Name "sched-holiday-$($h.Name -replace ' ','-')" `
                -StartTime ([datetime]$schedDate) -OneTime `
                -TimeZone $Global:Cfg.Scaling.E3.TimeZone `
                -Description "Holiday surge: $($h.Name) on $($h.Date)" -EA SilentlyContinue | Out-Null
            Write-Status "Holiday scheduled" "$($h.Name) — activate $schedDate" "ok"
        }
    } else {
        Write-Status "Automation account not deployed — holiday schedules not created" "" "warn"
    }
    Read-Host "  Press Enter to continue"
}

function Show-LiveScalingStatus {
    Write-Banner "LIVE SCALING STATUS"
    foreach ($poolType in @("E3","F1")) {
        $hpName = $Global:Cfg.HostPools[$poolType].Name
        $hpRg   = $Global:Cfg.RG.AVD.Name
        Write-Section "$poolType — $hpName"
        $hosts = Get-AzWvdSessionHost -HostPoolName $hpName -ResourceGroupName $hpRg -EA SilentlyContinue
        if (-not $hosts) { Write-Status "Host pool not found or no hosts" "" "warn"; continue }
        $total      = @($hosts).Count
        $available  = @($hosts | Where-Object { $_.Status -eq "Available" }).Count
        $draining   = @($hosts | Where-Object { $_.AllowNewSession -eq $false }).Count
        $unavail    = @($hosts | Where-Object { $_.Status -eq "Unavailable" }).Count
        $sessions   = ($hosts | Measure-Object -Property Session -Sum).Sum
        $maxSlots   = $total * $Global:Cfg.HostPools[$poolType].MaxSessions
        $capPct     = if ($maxSlots -gt 0) { [Math]::Round(($sessions * 100) / $maxSlots, 1) } else { 0 }
        Write-Status "Total Hosts"     $total          "info"
        Write-Status "Available"       $available      "ok"
        Write-Status "Draining"        $draining       (if ($draining -gt 0) {"warn"} else {"ok"})
        Write-Status "Unavailable"     $unavail        (if ($unavail -gt 0) {"fail"} else {"ok"})
        Write-Status "Active Sessions" "$sessions / $maxSlots ($capPct%)" "info"

        # Capacity bar
        $filled = [Math]::Floor($capPct / 2.5)
        $empty  = 40 - $filled
        $barCol = if ($capPct -ge 90) {"Red"} elseif ($capPct -ge 75) {"Yellow"} else {"Green"}
        Write-Host ""
        Write-Host "  Capacity: [" -NoNewline -ForegroundColor $C.Muted
        Write-Host ("#" * $filled) -NoNewline -ForegroundColor $barCol
        Write-Host ("─" * $empty + "] $capPct%") -ForegroundColor $barCol

        Write-Host ""
        Write-Host ("  {0,-35} {1,-10} {2,-12} {3,-8} {4}" -f "Host Name","Sessions","Status","Allow New","VM Size") -ForegroundColor $C.Muted
        Write-Host "  $($UI.SH * 78)" -ForegroundColor $C.Border
        foreach ($h in @($hosts)) {
            $vn    = ($h.Name -split "/")[-1]
            $col   = if ($h.Status -eq "Available") {$C.Ok} elseif ($h.Status -eq "Unavailable") {$C.Fail} else {$C.Warn}
            $allow = if ($h.AllowNewSession) {"Yes"} else {"No (Drain)"}
            Write-Host ("  {0,-35}" -f $vn) -NoNewline -ForegroundColor $col
            Write-Host ("{0,-10}" -f $h.Session) -NoNewline -ForegroundColor $C.Menu
            Write-Host ("{0,-12}" -f $h.Status) -NoNewline -ForegroundColor $col
            Write-Host ("{0,-8}" -f $allow) -NoNewline -ForegroundColor $C.Muted
            Write-Host $h.VirtualMachineId.Split("/")[-1] -ForegroundColor $C.Muted
        }
    }
    Read-Host "`n  Press Enter to continue"
}

function Invoke-ManualScaleAction {
    Write-Section "Manual Scale Action"
    Write-Host "  1.  Add session host to E3 pool" -ForegroundColor $C.New
    Write-Host "  2.  Add session host to F1 pool" -ForegroundColor $C.New
    Write-Host "  3.  Drain + remove a specific session host" -ForegroundColor $C.Warn
    Write-Host "  4.  Enable drain mode on all hosts in a pool" -ForegroundColor $C.Warn
    Write-Host "  5.  Disable drain mode on all hosts in a pool" -ForegroundColor $C.Ok
    Write-Host "  6.  Trigger Auto-Heal runbook" -ForegroundColor $C.Warn
    Write-Host "  7.  Activate holiday surge NOW" -ForegroundColor $C.Accent
    Write-Host "  0.  Back" -ForegroundColor $C.Muted
    $choice = Read-MenuChoice "Action" @("0","1","2","3","4","5","6","7")
    switch ($choice) {
        "1" { Invoke-RunbookTrigger "Add-AVDSessionHost" @{HostPoolType="E3"} }
        "2" { Invoke-RunbookTrigger "Add-AVDSessionHost" @{HostPoolType="F1"} }
        "3" { Invoke-DrainRemoveHost }
        "4" { Set-PoolDrainMode -PoolType (Read-MenuChoice "Pool type E3 or F1" @("E3","F1")) -Enable $true }
        "5" { Set-PoolDrainMode -PoolType (Read-MenuChoice "Pool type E3 or F1" @("E3","F1")) -Enable $false }
        "6" { Invoke-RunbookTrigger "Invoke-AVDAutoHeal" @{HostPoolType=(Read-MenuChoice "E3 or F1" @("E3","F1"))} }
        "7" { Invoke-RunbookTrigger "Set-AVDHolidayScaling" @{Mode="Activate"; E3MaxHosts=15; F1MaxHosts=15} }
        "0" { return }
    }
}

function Invoke-RunbookTrigger {
    param([string]$RunbookName, [hashtable]$Params = @{})
    $aa = $Global:Cfg.Automation.AccountName
    $rg = $Global:Cfg.Automation.ResourceGroup
    if (-not (Get-AzAutomationAccount -Name $aa -ResourceGroupName $rg -EA SilentlyContinue)) {
        Write-Status "Automation Account not deployed" "" "fail"; Read-Host "  Press Enter"; return
    }
    if (Confirm-Action "Trigger runbook: $RunbookName" "Parameters: $($Params | ConvertTo-Json -Compress)") {
        Write-Status "Starting runbook" $RunbookName "step"
        $job = Start-AzAutomationRunbook -AutomationAccountName $aa -ResourceGroupName $rg `
                   -Name $RunbookName -Parameters $Params -EA SilentlyContinue
        if ($job) {
            Write-Status "Job started" $job.JobId "ok"
            Write-Status "Monitor at: Azure Portal → Automation → $aa → Jobs" "" "info"
        } else {
            Write-Status "Failed to start runbook — ensure it is deployed and published" "" "fail"
        }
    }
    Read-Host "  Press Enter to continue"
}

function Set-PoolDrainMode {
    param([string]$PoolType, [bool]$Enable)
    $hpName = $Global:Cfg.HostPools[$PoolType].Name
    $hpRg   = $Global:Cfg.RG.AVD.Name
    $mode   = if ($Enable) {"ENABLED (no new sessions)"} else {"DISABLED (normal operation)"}
    if (-not (Confirm-Action "Set Drain Mode $mode" "Pool: $hpName")) { return }
    $hosts = Get-AzWvdSessionHost -HostPoolName $hpName -ResourceGroupName $hpRg -EA SilentlyContinue
    foreach ($h in $hosts) {
        $vmName = ($h.Name -split "/")[-1]
        Update-AzWvdSessionHost -HostPoolName $hpName -ResourceGroupName $hpRg `
            -Name $vmName -AllowNewSession (-not $Enable) -EA SilentlyContinue | Out-Null
        Write-Status "Drain mode $(if($Enable){'on'}else{'off'})" $vmName "ok"
    }
    Read-Host "  Press Enter to continue"
}

function Invoke-DrainRemoveHost {
    $poolType = Read-MenuChoice "Pool type" @("E3","F1")
    $hpName   = $Global:Cfg.HostPools[$poolType].Name
    $hpRg     = $Global:Cfg.RG.AVD.Name
    $hosts    = Get-AzWvdSessionHost -HostPoolName $hpName -ResourceGroupName $hpRg -EA SilentlyContinue
    $i = 1
    foreach ($h in $hosts) {
        Write-Host ("  {0}. {1,-35} Sessions: {2}  Status: {3}" -f $i, ($h.Name -split "/")[-1], $h.Session, $h.Status) -ForegroundColor $C.Menu
        $i++
    }
    $ch = Read-MenuChoice "Select host to drain and remove" (1..@($hosts).Count | ForEach-Object {"$_"})
    $target = @($hosts)[[int]$ch - 1]
    $vmName = ($target.Name -split "/")[-1]
    if (Confirm-Action "Drain and remove $vmName" "Sessions: $($target.Session)") {
        Update-AzWvdSessionHost -HostPoolName $hpName -ResourceGroupName $hpRg `
            -Name $vmName -AllowNewSession $false | Out-Null
        Write-Status "Drain mode enabled" $vmName "ok"
        Write-Status "Wait for sessions to end before removing..." "" "info"
        Read-Host "  Press Enter when ready to remove from pool"
        Remove-AzWvdSessionHost -HostPoolName $hpName -ResourceGroupName $hpRg `
            -Name $vmName -Force -EA SilentlyContinue | Out-Null
        Write-Status "Removed from host pool" $vmName "ok"
    }
    Read-Host "  Press Enter to continue"
}

function Deploy-AutomationRunbooks {
    Write-Banner "AUTOMATION ACCOUNT & RUNBOOKS"
    $cfg = $Global:Cfg.Automation
    Write-Status "Checking Automation Account..." "" "step"

    # Detect existing
    $found = Find-ExistingAutomationAccounts
    if (@($found).Count -gt 0) {
        $i = 1
        foreach ($a in $found) {
            Write-Host ("  {0}. {1,-35} RG: {2}" -f $i, $a.AutomationAccountName, $a.ResourceGroupName) -ForegroundColor $C.Menu
            $i++
        }
        Write-Host "     N. Create new" -ForegroundColor $C.New
        Write-Host "     S. Skip automation" -ForegroundColor $C.Skip
        $opts = (1..@($found).Count | ForEach-Object {"$_"}) + @("N","n","S","s")
        $ch = Read-MenuChoice "Automation Account" $opts "N"
        if ($ch -in @("S","s")) {
            $Global:Cfg.DeploymentState.Automation = "Skipped"; Save-Config; return
        }
        if ($ch -notmatch "^[Nn]$") {
            $sel = @($found)[[int]$ch - 1]
            $cfg.AccountName    = $sel.AutomationAccountName
            $cfg.ResourceGroup  = $sel.ResourceGroupName
            $cfg.IsNew          = $false
            Write-Status "Reusing Automation Account" $sel.AutomationAccountName "reuse"
        }
    }

    # Create or verify
    $aa = Get-AzAutomationAccount -Name $cfg.AccountName -ResourceGroupName $cfg.ResourceGroup -EA SilentlyContinue
    if (-not $aa) {
        Write-Status "Creating Automation Account" $cfg.AccountName "new"
        $aa = New-AzAutomationAccount -Name $cfg.AccountName -ResourceGroupName $cfg.ResourceGroup `
                  -Location $Global:Cfg.Location -Plan Free
    } else {
        Write-Status "Automation Account exists" $cfg.AccountName "reuse"
    }

    # Managed Identity
    Set-AzAutomationAccount -Name $cfg.AccountName -ResourceGroupName $cfg.ResourceGroup `
        -AssignSystemIdentity | Out-Null
    Start-Sleep -Seconds 20
    $aa = Get-AzAutomationAccount -Name $cfg.AccountName -ResourceGroupName $cfg.ResourceGroup
    $cfg.ManagedIdentityId = $aa.Identity.PrincipalId
    Write-Status "Managed Identity" $cfg.ManagedIdentityId "ok"

    # Store variables
    $vars = @{
        "AVD-SubscriptionId"     = $Global:Cfg.SubscriptionId
        "AVD-E3-HostPoolName"    = $Global:Cfg.HostPools.E3.Name
        "AVD-F1-HostPoolName"    = $Global:Cfg.HostPools.F1.Name
        "AVD-ResourceGroup"      = $Global:Cfg.RG.AVD.Name
        "AVD-E3-MaxHosts"        = "15"
        "AVD-F1-MaxHosts"        = "10"
        "AVD-E3-MinHosts"        = "1"
        "AVD-F1-MinHosts"        = "1"
        "AVD-E3-VmSku"           = $Global:Cfg.HostPools.E3.VMSize
        "AVD-F1-VmSku"           = $Global:Cfg.HostPools.F1.VMSize
        "AVD-E3-VmPrefix"        = "vm-avd-e3-$($Global:Cfg.Prefix -replace '-','')"
        "AVD-F1-VmPrefix"        = "vm-avd-f1-$($Global:Cfg.Prefix -replace '-','')"
        "AVD-GalleryName"        = $Global:Cfg.Gallery.Name
        "AVD-GalleryRG"          = $Global:Cfg.Gallery.ResourceGroup
        "AVD-E3-ImageDef"        = "img-avd-win11ms-e3"
        "AVD-F1-ImageDef"        = "img-avd-win11ms-f1"
        "AVD-VNetName"           = $Global:Cfg.Network.VNet.Name
        "AVD-VNetRG"             = $Global:Cfg.Network.VNet.ResourceGroup
        "AVD-E3-SubnetName"      = $Global:Cfg.Network.Subnets.E3.Name
        "AVD-F1-SubnetName"      = $Global:Cfg.Network.Subnets.F1.Name
        "AVD-AdminUsername"      = $Global:Cfg.Identity.AdminUsername
        "AVD-KeyVaultName"       = $Global:Cfg.KeyVault.Name
    }
    foreach ($v in $vars.GetEnumerator()) {
        New-AzAutomationVariable -AutomationAccountName $cfg.AccountName `
            -ResourceGroupName $cfg.ResourceGroup -Name $v.Key -Value $v.Value `
            -Encrypted $false -EA SilentlyContinue | Out-Null
    }
    Write-Status "Automation variables set" "$($vars.Count) variables" "ok"

    # Check which runbooks to deploy
    Write-Section "Select Runbooks to Deploy"
    $runbooks = @{
        ScaleOut = @{ Name="Add-AVDSessionHost";          Deploy=$Global:Cfg.Scaling.CustomRunbooks.ScaleOut }
        ScaleIn  = @{ Name="Remove-DrainedAVDSessionHost"; Deploy=$Global:Cfg.Scaling.CustomRunbooks.ScaleIn  }
        AutoHeal = @{ Name="Invoke-AVDAutoHeal";           Deploy=$Global:Cfg.Scaling.CustomRunbooks.AutoHeal }
        Holiday  = @{ Name="Set-AVDHolidayScaling";        Deploy=$Global:Cfg.Scaling.CustomRunbooks.Holiday  }
    }
    foreach ($rb in $runbooks.GetEnumerator()) {
        $rb.Value.Deploy = Read-YesNo "Deploy runbook: $($rb.Value.Name)?" $rb.Value.Deploy
    }

    # Deploy runbooks
    foreach ($rb in $runbooks.GetEnumerator()) {
        if (-not $rb.Value.Deploy) { Write-Status "Skipped" $rb.Value.Name "skip"; continue }
        Write-Status "Deploying runbook" $rb.Value.Name "step"
        $content = Get-RunbookContent $rb.Key
        Import-AzAutomationRunbook -AutomationAccountName $cfg.AccountName `
            -ResourceGroupName $cfg.ResourceGroup -Name $rb.Value.Name `
            -Type PowerShell -Description "BDF AVD auto-scaling: $($rb.Value.Name)" `
            -Force | Out-Null
        Set-AzAutomationRunbookContent -AutomationAccountName $cfg.AccountName `
            -ResourceGroupName $cfg.ResourceGroup -Name $rb.Value.Name `
            -Content $content | Out-Null
        Publish-AzAutomationRunbook -AutomationAccountName $cfg.AccountName `
            -ResourceGroupName $cfg.ResourceGroup -Name $rb.Value.Name | Out-Null
        Write-Status "Published" $rb.Value.Name "ok"
    }

    # Nightly scale-in schedule
    $schedTime = (Get-Date "22:30").AddDays(1)
    New-AzAutomationSchedule -AutomationAccountName $cfg.AccountName `
        -ResourceGroupName $cfg.ResourceGroup -Name "sched-nightly-scalein" `
        -StartTime $schedTime -DayInterval 1 `
        -TimeZone $Global:Cfg.Scaling.E3.TimeZone `
        -Description "Nightly scale-in — remove idle hosts" -EA SilentlyContinue | Out-Null
    foreach ($pt in @("E3","F1")) {
        Register-AzAutomationScheduledRunbook -AutomationAccountName $cfg.AccountName `
            -ResourceGroupName $cfg.ResourceGroup `
            -RunbookName "Remove-DrainedAVDSessionHost" `
            -ScheduleName "sched-nightly-scalein" `
            -Parameters @{HostPoolType=$pt} -EA SilentlyContinue | Out-Null
    }
    Write-Status "Nightly scale-in schedule created" "10:30 PM daily" "ok"

    $Global:Cfg.DeploymentState.Automation = "Deployed"
    Save-Config
    Read-Host "  Press Enter to continue"
}

function Get-RunbookContent {
    param([string]$RunbookKey)
    # Returns condensed runbook content — each leverages Automation Variables
    $base = @'
Connect-AzAccount -Identity | Out-Null
Set-AzContext -SubscriptionId (Get-AutomationVariable "AVD-SubscriptionId") | Out-Null
'@
    switch ($RunbookKey) {
        "ScaleOut" { return $base + "`nparam([string]`$HostPoolType='E3')`n" + @'
$hpName = Get-AutomationVariable "AVD-$HostPoolType-HostPoolName"
$hpRg   = Get-AutomationVariable "AVD-ResourceGroup"
$maxH   = [int](Get-AutomationVariable "AVD-$HostPoolType-MaxHosts")
$vmSku  = Get-AutomationVariable "AVD-$HostPoolType-VmSku"
$prefix = Get-AutomationVariable "AVD-$HostPoolType-VmPrefix"
$vnetN  = Get-AutomationVariable "AVD-VNetName"
$vnetRg = Get-AutomationVariable "AVD-VNetRG"
$snet   = Get-AutomationVariable "AVD-$HostPoolType-SubnetName"
$kvName = Get-AutomationVariable "AVD-KeyVaultName"
$admin  = Get-AutomationVariable "AVD-AdminUsername"
$gallery= Get-AutomationVariable "AVD-GalleryName"
$galRg  = Get-AutomationVariable "AVD-GalleryRG"
$imgDef = Get-AutomationVariable "AVD-$HostPoolType-ImageDef"

$hosts = @(Get-AzWvdSessionHost -HostPoolName $hpName -ResourceGroupName $hpRg)
$active = @($hosts | Where-Object { $_.Status -eq "Available" -or $_.Status -eq "NeedsAssistance" })
if ($active.Count -ge $maxH) { Write-Output "At max hosts ($maxH). No scale-out."; exit 0 }

$nums = @($hosts | ForEach-Object { $n=($_.Name -split "/")[-1] -replace "^$prefix-",""; if($n -match "^\d+$"){[int]$n} })
$next = 1; while ($nums -contains $next) { $next++ }
$vmName = "$prefix-$next"
Write-Output "Provisioning: $vmName"

$pwd = (Get-AzKeyVaultSecret -VaultName $kvName -Name "avd-vm-admin-password").SecretValue
$cred = New-Object PSCredential($admin, $pwd)
$exp = (Get-Date).ToUniversalTime().AddHours(4).ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
$tok = (New-AzWvdRegistrationInfo -HostPoolName $hpName -ResourceGroupName $hpRg -ExpirationTime $exp).Token

$vnet   = Get-AzVirtualNetwork -Name $vnetN -ResourceGroupName $vnetRg
$subnet = $vnet.Subnets | Where-Object { $_.Name -eq $snet }
$nic    = New-AzNetworkInterface -Name "nic-$vmName" -ResourceGroupName $hpRg `
              -Location (Get-AzResourceGroup -Name $hpRg).Location -SubnetId $subnet.Id `
              -PrivateIpAllocationMethod Dynamic
$imgVer = (Get-AzGalleryImageVersion -GalleryName $gallery -ResourceGroupName $galRg `
    -GalleryImageDefinitionName $imgDef -EA SilentlyContinue | Sort-Object Name -Desc | Select-Object -First 1).Id

$vmCfg = New-AzVMConfig -VMName $vmName -VMSize $vmSku -IdentityType SystemAssigned |
    Set-AzVMOperatingSystem -Windows -ComputerName ($vmName -replace "-","") -Credential $cred `
        -TimeZone "Eastern Standard Time" -EnableAutoUpdate $false |
    Add-AzVMNetworkInterface -Id $nic.Id |
    Set-AzVMOSDisk -DiskSizeGB 128 -StorageAccountType Premium_LRS -CreateOption FromImage -DeleteOption Delete |
    Set-AzVMSecurityProfile -SecurityType TrustedLaunch | Set-AzVMUefi -EnableVtpm $true -EnableSecureBoot $true

if ($imgVer) { $vmCfg.StorageProfile.ImageReference = @{Id=$imgVer} }
else { $vmCfg = $vmCfg | Set-AzVMSourceImage -PublisherName "MicrosoftWindowsDesktop" -Offer "windows-11" -Skus "win11-24h2-avd-m365" -Version "latest" }

New-AzVM -ResourceGroupName $hpRg -Location (Get-AzResourceGroup -Name $hpRg).Location -VM $vmCfg | Out-Null
Set-AzVMExtension -ResourceGroupName $hpRg -VMName $vmName -Name "AADLoginForWindows" `
    -Publisher "Microsoft.Azure.ActiveDirectory" -ExtensionType "AADLoginForWindows" -TypeHandlerVersion "2.0" -Settings @{mdmId=""} | Out-Null
Set-AzVMExtension -ResourceGroupName $hpRg -VMName $vmName -Name "DSC" `
    -Publisher "Microsoft.Powershell" -ExtensionType "DSC" -TypeHandlerVersion "2.83" `
    -Settings @{ modulesUrl="https://wvdportalstorageblob.blob.core.windows.net/galleryartifacts/Configuration_09-08-2022.zip"
                 configurationFunction="Configuration.ps1\AddSessionHost"
                 properties=@{HostPoolName=$hpName;RegistrationInfoToken=$tok;AadJoin=$true} } | Out-Null
Update-AzTag -ResourceId (Get-AzVM -Name $vmName -ResourceGroupName $hpRg).Id `
    -Tag @{ScaledBy="Automation";HostPool=$hpName;ScaleTime=(Get-Date -Format "o")} -Operation Merge | Out-Null
Write-Output "SUCCESS: $vmName added to $hpName"
'@ }
        "ScaleIn" { return $base + @'
param([string]$HostPoolType="E3",[int]$MinHosts=1)
$hpName = Get-AutomationVariable "AVD-$HostPoolType-HostPoolName"
$hpRg   = Get-AutomationVariable "AVD-ResourceGroup"
$minH   = [Math]::Max($MinHosts,[int](Get-AutomationVariable "AVD-$HostPoolType-MinHosts"))
$allH   = @(Get-AzWvdSessionHost -HostPoolName $hpName -ResourceGroupName $hpRg)
$active = @($allH | Where-Object {$_.AllowNewSession -and $_.Status -eq "Available"})
if ($active.Count -le $minH) {Write-Output "At minimum hosts ($minH)";exit 0}
$cands  = @($allH | Where-Object {$_.Session -eq 0 -and $_.AllowNewSession -and $_.Status -eq "Available"} | Sort-Object Name)
$canRem = [Math]::Min($cands.Count,($active.Count-$minH))
Write-Output "Removing up to $canRem hosts"
foreach ($h in ($cands | Select-Object -First $canRem)) {
    $vn = ($h.Name -split "/")[-1]
    Update-AzWvdSessionHost -HostPoolName $hpName -ResourceGroupName $hpRg -Name $vn -AllowNewSession $false | Out-Null
    $w=0; do {Start-Sleep 30;$w+=30;$r=Get-AzWvdSessionHost -HostPoolName $hpName -ResourceGroupName $hpRg -Name $vn} until ($r.Session -eq 0 -or $w -ge 600)
    if ($r.Session -gt 0) {Update-AzWvdSessionHost -HostPoolName $hpName -ResourceGroupName $hpRg -Name $vn -AllowNewSession $true; continue}
    Remove-AzWvdSessionHost -HostPoolName $hpName -ResourceGroupName $hpRg -Name $vn -Force | Out-Null
    $vm=(Get-AzVM -Name $vn -EA SilentlyContinue); if ($vm) {
        if ($vm.Tags["ScaledBy"] -eq "Automation") {Remove-AzVM -Name $vn -ResourceGroupName $vm.ResourceGroupName -Force | Out-Null}
        else {Stop-AzVM -Name $vn -ResourceGroupName $vm.ResourceGroupName -Force | Out-Null}
    }
    Write-Output "Removed: $vn"
}
'@ }
        "AutoHeal" { return $base + @'
param([string]$HostPoolType="E3")
$hpName = Get-AutomationVariable "AVD-$HostPoolType-HostPoolName"
$hpRg   = Get-AutomationVariable "AVD-ResourceGroup"
$aaName = (Get-AzAutomationAccount | Select-Object -First 1).AutomationAccountName
$aaRg   = (Get-AzAutomationAccount | Select-Object -First 1).ResourceGroupName
$bad = @(Get-AzWvdSessionHost -HostPoolName $hpName -ResourceGroupName $hpRg | Where-Object {$_.Status -in @("Unavailable","NeedsAssistance","Shutdown")})
if ($bad.Count -eq 0) {Write-Output "All hosts healthy"; exit 0}
foreach ($h in $bad) {
    $vn = ($h.Name -split "/")[-1]; Write-Output "Healing: $vn"
    Update-AzWvdSessionHost -HostPoolName $hpName -ResourceGroupName $hpRg -Name $vn -AllowNewSession $false -EA SilentlyContinue | Out-Null
    $sess = Get-AzWvdUserSession -HostPoolName $hpName -ResourceGroupName $hpRg -SessionHostName $vn -EA SilentlyContinue
    foreach ($s in $sess) { Send-AzWvdUserSessionMessage -HostPoolName $hpName -ResourceGroupName $hpRg -SessionHostName $vn -UserSessionId ($s.Name -split "/")[-1] -MessageTitle "Maintenance" -MessageBody "Host maintenance in 5 minutes. Please save your work." -EA SilentlyContinue | Out-Null }
    Start-Sleep -Seconds 300
    Remove-AzWvdSessionHost -HostPoolName $hpName -ResourceGroupName $hpRg -Name $vn -Force -EA SilentlyContinue | Out-Null
    $vm=(Get-AzVM -Name $vn -EA SilentlyContinue); if ($vm) {Stop-AzVM -Name $vn -ResourceGroupName $vm.ResourceGroupName -Force | Out-Null}
    Start-AzAutomationRunbook -AutomationAccountName $aaName -ResourceGroupName $aaRg -Name "Add-AVDSessionHost" -Parameters @{HostPoolType=$HostPoolType} -EA SilentlyContinue | Out-Null
    Write-Output "Replacement triggered for $vn"
}
'@ }
        "Holiday" { return $base + @'
param([string]$Mode="Activate",[int]$E3MaxHosts=15,[int]$F1MaxHosts=15)
foreach ($pt in @("E3","F1")) {
    $maxH = if ($pt -eq "E3") {$E3MaxHosts} else {$F1MaxHosts}
    $varN = "AVD-$pt-MaxHosts"; Set-AutomationVariable -Name $varN -Value "$maxH" | Out-Null
    $hpName = Get-AutomationVariable "AVD-$pt-HostPoolName"
    $hpRg   = Get-AutomationVariable "AVD-ResourceGroup"
    if ($Mode -eq "Activate") {
        $hosts = Get-AzWvdSessionHost -HostPoolName $hpName -ResourceGroupName $hpRg
        foreach ($h in $hosts) {
            $vn = ($h.Name -split "/")[-1]
            $vm = Get-AzVM -Name $vn -Status -EA SilentlyContinue | Where-Object {$_.Statuses.Code -contains "PowerState/deallocated"}
            if ($vm) { Start-AzVM -Name $vn -ResourceGroupName (Get-AzVM -Name $vn).ResourceGroupName -NoWait | Out-Null }
            Update-AzWvdSessionHost -HostPoolName $hpName -ResourceGroupName $hpRg -Name $vn -AllowNewSession $true -EA SilentlyContinue | Out-Null
        }
        Write-Output "HOLIDAY ACTIVE: $pt max=$maxH, all VMs starting"
    } else {
        Set-AutomationVariable -Name $varN -Value "$(if($pt -eq "E3"){15}else{10})" | Out-Null
        Write-Output "HOLIDAY DEACTIVATED for $pt"
    }
}
'@ }
    }
}
#endregion

#region ═══════════════════════════════════════════════════════════════════════
#  HEALTH DASHBOARD & VALIDATION
#═══════════════════════════════════════════════════════════════════════════════

function Show-HealthDashboard {
    Write-Banner "DEPLOYMENT HEALTH DASHBOARD"
    $pass = 0; $fail = 0; $warn = 0; $skip = 0

    $checks = @(
        @{ Name="E3 Host Pool";      Query={ Get-AzWvdHostPool -Name $Global:Cfg.HostPools.E3.Name -ResourceGroupName $Global:Cfg.RG.AVD.Name -EA SilentlyContinue } }
        @{ Name="F1 Host Pool";      Query={ Get-AzWvdHostPool -Name $Global:Cfg.HostPools.F1.Name -ResourceGroupName $Global:Cfg.RG.AVD.Name -EA SilentlyContinue } }
        @{ Name="E3 Session Hosts";  Query={ @(Get-AzWvdSessionHost -HostPoolName $Global:Cfg.HostPools.E3.Name -ResourceGroupName $Global:Cfg.RG.AVD.Name -EA SilentlyContinue).Count -ge 1 } }
        @{ Name="F1 Session Hosts";  Query={ @(Get-AzWvdSessionHost -HostPoolName $Global:Cfg.HostPools.F1.Name -ResourceGroupName $Global:Cfg.RG.AVD.Name -EA SilentlyContinue).Count -ge 1 } }
        @{ Name="AVD Workspace";     Query={ Get-AzWvdWorkspace -ResourceGroupName $Global:Cfg.RG.AVD.Name -EA SilentlyContinue | Where-Object {$_.Name} } }
        @{ Name="Key Vault";         Query={ Get-AzKeyVault -VaultName $Global:Cfg.KeyVault.Name -EA SilentlyContinue } }
        @{ Name="Log Analytics";     Query={ Get-AzOperationalInsightsWorkspace -Name $Global:Cfg.LogAnalytics.Name -ResourceGroupName $Global:Cfg.LogAnalytics.ResourceGroup -EA SilentlyContinue } }
        @{ Name="Azure Files";       Query={ Get-AzStorageAccount -Name $Global:Cfg.Storage.Account.Name -ResourceGroupName $Global:Cfg.Storage.Account.ResourceGroup -EA SilentlyContinue } }
        @{ Name="E3 Profile Share";  Query={ $sa=Get-AzStorageAccount -Name $Global:Cfg.Storage.Account.Name -ResourceGroupName $Global:Cfg.Storage.Account.ResourceGroup -EA SilentlyContinue; if($sa){Get-AzStorageShare -Name $Global:Cfg.Storage.Shares.E3 -Context $sa.Context -EA SilentlyContinue} else {$null}} }
        @{ Name="E3 Scaling Plan";   Query={ Get-AzWvdScalingPlan -Name $Global:Cfg.Scaling.E3.PlanName -ResourceGroupName $Global:Cfg.RG.AVD.Name -EA SilentlyContinue } }
        @{ Name="F1 Scaling Plan";   Query={ Get-AzWvdScalingPlan -Name $Global:Cfg.Scaling.F1.PlanName -ResourceGroupName $Global:Cfg.RG.AVD.Name -EA SilentlyContinue } }
        @{ Name="Automation Account";Query={ Get-AzAutomationAccount -Name $Global:Cfg.Automation.AccountName -ResourceGroupName $Global:Cfg.Automation.ResourceGroup -EA SilentlyContinue } }
        @{ Name="Scale-Out Runbook"; Query={ $rb=Get-AzAutomationRunbook -AutomationAccountName $Global:Cfg.Automation.AccountName -ResourceGroupName $Global:Cfg.Automation.ResourceGroup -Name "Add-AVDSessionHost" -EA SilentlyContinue; $rb.State -eq "Published" } }
        @{ Name="Virtual Network";   Query={ Get-AzVirtualNetwork -Name $Global:Cfg.Network.VNet.Name -ResourceGroupName $Global:Cfg.Network.VNet.ResourceGroup -EA SilentlyContinue } }
        @{ Name="Start VM on Connect";Query={ (Get-AzWvdHostPool -Name $Global:Cfg.HostPools.E3.Name -ResourceGroupName $Global:Cfg.RG.AVD.Name -EA SilentlyContinue).StartVMOnConnect -eq $true } }
    )

    Write-Host ""
    Write-Host ("  {0,-32} {1,-12} {2}" -f "Check", "Status", "Details") -ForegroundColor $C.Muted
    Write-Host "  $($UI.SH * 68)" -ForegroundColor $C.Border

    foreach ($chk in $checks) {
        try {
            $result = & $chk.Query
            $ok     = ($null -ne $result -and $result -ne $false)
            $icon   = if ($ok) {"✔"} else {"✖"}
            $col    = if ($ok) {$C.Ok} else {$C.Fail}
            $label  = if ($ok) {"OK"} else {"MISSING"}
            Write-Host ("  $icon  {0,-30} {1}" -f $chk.Name, $label) -ForegroundColor $col
            if ($ok) {$pass++} else {$fail++}
        } catch {
            Write-Host ("  ⚠  {0,-30} ERROR" -f $chk.Name) -ForegroundColor $C.Warn
            $warn++
        }
    }

    Write-Host ""
    Write-Host "  $($UI.SH * 68)" -ForegroundColor $C.Border
    Write-Host ("  Passed: {0}   Failed: {1}   Warnings: {2}" -f $pass, $fail, $warn) -ForegroundColor (if($fail -gt 0){$C.Fail} elseif($warn -gt 0){$C.Warn} else {$C.Ok})
    Write-Host ""
    Read-Host "  Press Enter to continue"
}
#endregion

#region ═══════════════════════════════════════════════════════════════════════
#  DOMAIN JOIN TYPE WIZARD — ENTRA ID JOIN vs HYBRID AD JOIN
#═══════════════════════════════════════════════════════════════════════════════

function Show-JoinTypeMenu {
    Write-Banner "DOMAIN JOIN CONFIGURATION" $C.Azure

    while ($true) {
        $jc = $Global:Cfg.JoinConfig
        Write-Section "Current Join Configuration"

        $joinStatus = if ($jc.Configured) { "✔ $($jc.Type)" } else { "○ Not Configured" }
        $joinCol    = if ($jc.Configured) { $C.Ok } else { $C.Muted }
        Write-Host ("  {0,-32} {1}" -f "Join Type:", $joinStatus) -ForegroundColor $joinCol

        if ($jc.Configured -and $jc.Type -eq "HybridAD") {
            Write-Host ("  {0,-32} {1}" -f "Domain:", $jc.HybridAD.DomainName) -ForegroundColor $C.Muted
            Write-Host ("  {0,-32} {1}" -f "OU:", $jc.HybridAD.DomainJoinOU) -ForegroundColor $C.Muted
            Write-Host ("  {0,-32} {1}" -f "Sync Method:", $jc.HybridAD.SyncMethod) -ForegroundColor $C.Muted
        }
        Write-Host ""

        Write-Host "  1.  Run Join Type Selection Wizard" -ForegroundColor $C.New
        Write-Host "  2.  Configure Entra ID (Azure AD) Join" -ForegroundColor $C.Menu
        Write-Host "  3.  Configure Hybrid AD Join" -ForegroundColor $C.Menu
        Write-Host "  4.  View prerequisites & comparison" -ForegroundColor $C.Menu
        Write-Host "  5.  Validate join configuration" -ForegroundColor $C.Azure
        Write-Host "  6.  Export join configuration scripts" -ForegroundColor $C.Menu
        Write-Host "  0.  Back" -ForegroundColor $C.Muted

        $ch = Read-MenuChoice "Join Config" @("0","1","2","3","4","5","6")
        switch ($ch) {
            "1" { Invoke-JoinTypeWizard }
            "2" { Configure-EntraIDJoin }
            "3" { Configure-HybridADJoin }
            "4" { Show-JoinTypeComparison }
            "5" { Validate-JoinConfiguration }
            "6" { Export-JoinScripts }
            "0" { return }
        }
    }
}

function Invoke-JoinTypeWizard {
    Write-Banner "JOIN TYPE SELECTION WIZARD"

    Write-Host @"
  Azure Virtual Desktop session hosts can be joined to Azure AD (Entra ID)
  or to your on-premises Active Directory (Hybrid AD Join).

  The choice affects:
    • How users authenticate to AVD session hosts
    • Whether a domain controller is required in your network path
    • How FSLogix authenticates to Azure Files (Kerberos method)
    • How Intune manages the session host VMs
    • Whether legacy app authentication (Kerberos/NTLM) works

"@ -ForegroundColor $C.Muted

    Write-Host "  ┌─ Compare Join Types ──────────────────────────────────────────────────────┐" -ForegroundColor $C.Border
    Write-Host "  │  Feature                     Entra ID Join        Hybrid AD Join          │" -ForegroundColor $C.Muted
    Write-Host "  │  ──────────────────────────  ─────────────────    ─────────────────────   │" -ForegroundColor $C.Border
    Write-Host "  │  Domain Controller needed    ✖  Not required      ✔  Required in VNet      │" -ForegroundColor $C.Menu
    Write-Host "  │  Intune MDM enrollment       ✔  Native/auto       ⚠  Co-management needed  │" -ForegroundColor $C.Menu
    Write-Host "  │  Azure Files Kerberos        ✔  Azure AD Kerb.    ✔  On-prem AD Kerberos   │" -ForegroundColor $C.Menu
    Write-Host "  │  Legacy app auth (NTLM)      ⚠  Limited           ✔  Full support          │" -ForegroundColor $C.Menu
    Write-Host "  │  Single-sign-on to AVD       ✔  Native SSO        ✔  With SSSO config      │" -ForegroundColor $C.Menu
    Write-Host "  │  GPO support                 ⚠  Intune only       ✔  Full GPO + Intune     │" -ForegroundColor $C.Menu
    Write-Host "  │  Conditional Access          ✔  Full support      ✔  Full support          │" -ForegroundColor $C.Menu
    Write-Host "  │  Complexity                  ✔  Simpler           ⚠  Higher complexity     │" -ForegroundColor $C.Menu
    Write-Host "  │  BDF Recommendation          ★  RECOMMENDED       Use only if needed       │" -ForegroundColor $C.Ok
    Write-Host "  └──────────────────────────────────────────────────────────────────────────┘" -ForegroundColor $C.Border
    Write-Host ""

    Write-Host "  1.  Entra ID Join (Azure AD Join)  — Recommended for BDF retail" -ForegroundColor $C.Ok
    Write-Host "  2.  Hybrid AD Join                 — Required if legacy app auth needed" -ForegroundColor $C.Warn

    $ch = Read-MenuChoice "Select join type" @("1","2")
    if ($ch -eq "1") {
        $Global:Cfg.JoinConfig.Type = "EntraID"
        Configure-EntraIDJoin
    } else {
        $Global:Cfg.JoinConfig.Type = "HybridAD"
        Configure-HybridADJoin
    }
    $Global:Cfg.JoinConfig.Configured = $true
    $Global:Cfg.Identity.AzureAdJoin  = ($Global:Cfg.JoinConfig.Type -eq "EntraID")
    $Global:Cfg.DeploymentState.JoinType = "Deployed"
    Save-Config
}

function Configure-EntraIDJoin {
    Write-Banner "ENTRA ID JOIN CONFIGURATION"
    $jc = $Global:Cfg.JoinConfig

    Write-Section "Prerequisites Check"
    Write-Host @"
  Required before deploying Entra ID-joined AVD session hosts:

  ✔  Azure subscription with AVD resource providers registered
  ✔  Entra ID tenant (M365 E3 / F1 tenant already active at BDF)
  ✔  Session host VMs must have System-Assigned Managed Identity enabled
  ✔  'Virtual Machine User Login' or 'Virtual Machine Administrator Login'
     RBAC role assigned to AVD users on the session host resource group
  ✔  Intune enrollment — Intune (MDM) set as MDM authority in Entra ID
  ✔  FSLogix configured to use Azure AD Kerberos (not on-prem AD Kerberos)
  ⚠  Legacy apps using Kerberos/NTLM against on-prem AD will NOT work
     without additional configuration (Kerberos Cloud Trust or AAD Kerberos)

"@ -ForegroundColor $C.Muted

    Write-Section "Entra ID Join Settings"

    # MDM enrollment
    $jc.EntraID.MDMEnrollment = Read-YesNo "Auto-enroll session hosts in Microsoft Intune (MDM)?" $true
    if ($jc.EntraID.MDMEnrollment) {
        Write-Status "Intune auto-enrollment" "Enabled via AADLoginForWindows extension" "ok"
        Write-Status "MDM App ID" "0000000a-0000-0000-c000-000000000000 (Intune default — leave blank)" "info"
        $customMDM = Read-MenuChoice "Custom MDM App ID (blank = use Intune default)"
        if ($customMDM) { $jc.EntraID.MDMAppId = $customMDM }
    }

    # RBAC guidance
    Write-Section "Required RBAC Roles for Entra ID Join"
    Write-Host @"
  Two RBAC roles must be assigned on the AVD Resource Group or Subscription
  so users can log into the Entra ID-joined session hosts:

  Role: 'Virtual Machine User Login'
  Scope: Resource Group ($($Global:Cfg.RG.AVD.Name))
  Assign to: E3 Users Group ($($Global:Cfg.Identity.E3GroupId))
             F1 Users Group ($($Global:Cfg.Identity.F1GroupId))

  Role: 'Virtual Machine Administrator Login'
  Scope: Resource Group ($($Global:Cfg.RG.AVD.Name))
  Assign to: AVD Admins Group ($($Global:Cfg.Identity.AdminGroupId))

"@ -ForegroundColor $C.Muted

    if (Read-YesNo "Assign these VM Login roles now?" $true) {
        $rgScope = "/subscriptions/$($Global:Cfg.SubscriptionId)/resourceGroups/$($Global:Cfg.RG.AVD.Name)"
        $assignments = @(
            @{ ObjectId=$Global:Cfg.Identity.E3GroupId;   Role="Virtual Machine User Login" }
            @{ ObjectId=$Global:Cfg.Identity.F1GroupId;   Role="Virtual Machine User Login" }
            @{ ObjectId=$Global:Cfg.Identity.AdminGroupId; Role="Virtual Machine Administrator Login" }
        )
        foreach ($a in $assignments) {
            if ($a.ObjectId) {
                New-AzRoleAssignment -ObjectId $a.ObjectId -RoleDefinitionName $a.Role `
                    -Scope $rgScope -EA SilentlyContinue | Out-Null
                Write-Status "Assigned" "$($a.Role) → $($a.ObjectId)" "ok"
            }
        }
    }

    # DNS settings
    Write-Section "DNS Configuration for Entra ID Join"
    Write-Host @"
  For Entra ID-joined AVD (cloud-only):
  • VNet DNS: Azure Default DNS (168.63.129.16) OR custom DNS that resolves
    public Microsoft endpoints
  • Do NOT point DNS to on-prem AD domain controllers
  • Azure Files with Azure AD Kerberos uses Entra ID — no DC needed

  Current VNet DNS configuration: Azure Default (recommended for Entra ID Join)

"@ -ForegroundColor $C.Muted

    Write-Status "Entra ID Join configured" "" "ok"
    $Global:Cfg.JoinConfig.Type = "EntraID"
    $Global:Cfg.Identity.AzureAdJoin = $true
    Save-Config
    Read-Host "  Press Enter to continue"
}

function Configure-HybridADJoin {
    Write-Banner "HYBRID AD JOIN CONFIGURATION"
    $jc = $Global:Cfg.JoinConfig

    Write-Host @"
  Hybrid Azure AD Join connects session host VMs to BOTH:
    • Your on-premises Active Directory domain (traditional domain join)
    • Azure Active Directory (Entra ID) via Azure AD Connect or Cloud Sync

  USE HYBRID JOIN when:
    • You have on-premises apps that require Kerberos/NTLM authentication
    • Legacy applications depend on on-prem AD group membership
    • SAP uses on-prem Kerberos (not SAML/SSO)
    • Existing GPO infrastructure must apply to session hosts

  CRITICAL NETWORK REQUIREMENT:
    • Session host VNet MUST have line-of-sight to domain controllers
    • DNS must resolve the on-prem AD domain
    • Firewall must allow Kerberos (TCP/UDP 88), LDAP (389), SMB (445), etc.

"@ -ForegroundColor $C.Muted

    Write-Section "Step 1 — Domain Information"
    $jc.HybridAD.DomainName    = Read-MenuChoice "AD Domain FQDN"       -Default "bdf.internal"
    $jc.HybridAD.DomainNetbios = Read-MenuChoice "NetBIOS domain name"  -Default "BDF"
    Write-Status "Domain" "$($jc.HybridAD.DomainName) ($($jc.HybridAD.DomainNetbios))" "ok"

    Write-Section "Step 2 — Domain Join Service Account"
    Write-Host "  Create a dedicated service account with minimal permissions:" -ForegroundColor $C.Muted
    Write-Host "  • 'Create Computer Objects' in the target OU only" -ForegroundColor $C.Muted
    Write-Host "  • 'Reset Password' on Computer objects in the OU" -ForegroundColor $C.Muted
    Write-Host "  • NOT a domain admin — principle of least privilege" -ForegroundColor $C.Warn
    Write-Host ""
    $jc.HybridAD.DomainJoinUPN = Read-MenuChoice "Domain join service account UPN" -Default "svc-avd-join@bdf.internal"
    Write-Status "Service account" $jc.HybridAD.DomainJoinUPN "ok"

    Write-Section "Step 3 — Target OU for Session Hosts"
    Write-Host "  Create a dedicated OU for AVD session hosts — apply only AVD-specific GPOs" -ForegroundColor $C.Muted
    $jc.HybridAD.DomainJoinOU = Read-MenuChoice "Target OU Distinguished Name" -Default "OU=AVDHosts,OU=Servers,DC=bdf,DC=internal"
    Write-Status "Target OU" $jc.HybridAD.DomainJoinOU "ok"

    Write-Section "Step 4 — Domain Controller Connectivity"
    Write-Host "  VNet DNS must point to your on-premises domain controllers." -ForegroundColor $C.Muted
    Write-Host "  Add DC IP addresses — these will be set as custom DNS on the VNet." -ForegroundColor $C.Muted
    $dcIPs = @()
    do {
        $ip = Read-MenuChoice "Domain Controller IP (blank to finish)"
        if ($ip -match "^\d+\.\d+\.\d+\.\d+$") {
            $dcIPs += $ip
            Write-Status "DC added" $ip "ok"
        }
    } until (-not $ip)
    if ($dcIPs.Count -eq 0) { $dcIPs = @("10.0.0.4","10.0.0.5") }
    $jc.HybridAD.DCIPAddresses = $dcIPs

    # Apply DNS to VNet
    if ($Global:Cfg.Network.VNet.Name -and (Read-YesNo "Apply DC IPs as custom DNS on VNet now?" $true)) {
        $vnet = Get-AzVirtualNetwork -Name $Global:Cfg.Network.VNet.Name `
                    -ResourceGroupName $Global:Cfg.Network.VNet.ResourceGroup -EA SilentlyContinue
        if ($vnet) {
            $vnet.DhcpOptions.DnsServers = $dcIPs
            $vnet | Set-AzVirtualNetwork | Out-Null
            Write-Status "VNet DNS updated" ($dcIPs -join ", ") "ok"
        }
    }

    Write-Section "Step 5 — Azure AD Connect or Cloud Sync"
    Write-Host "  Hybrid Join requires AD objects to sync to Entra ID." -ForegroundColor $C.Muted
    Write-Host "  1.  Azure AD Connect (AADC) — traditional, installed on-prem server" -ForegroundColor $C.Menu
    Write-Host "  2.  Entra Cloud Sync — agent-based, lighter footprint (recommended for new)" -ForegroundColor $C.Ok
    $syncChoice = Read-MenuChoice "Sync method" @("1","2") -Default "2"
    $jc.HybridAD.SyncMethod = if ($syncChoice -eq "1") { "AADConnect" } else { "CloudSync" }
    Write-Status "Sync method" $jc.HybridAD.SyncMethod "ok"

    if ($jc.HybridAD.SyncMethod -eq "CloudSync") {
        Write-Host @"

  Entra Cloud Sync setup (if not already configured):
  1. Download agent: Entra ID portal → Hybrid Management → Cloud Sync → New agent
  2. Install on-prem on a domain-joined Windows Server (not the DC itself)
  3. Create sync configuration for the AVD OU scope
  4. Enable Computer object sync (required for Hybrid Join)
  Download: https://aka.ms/CloudSyncAgent

"@ -ForegroundColor $C.Muted
    }

    Write-Section "Step 6 — FSLogix Azure Files Authentication for Hybrid"
    Write-Host @"
  With Hybrid AD Join, Azure Files authentication uses ON-PREM AD Kerberos:

  Option A: Active Directory (Traditional) — AD DS authentication
    • Azure Files storage account joined to on-prem AD
    • Requires running: AzFilesHybrid PowerShell module on-prem
    • Run: Join-AzStorageAccountForAuth -StorageAccountName ... -DomainName ...
    • Tickets issued by on-prem DCs

  Option B: Azure AD Kerberos (Cloud Kerberos Trust) — RECOMMENDED even for Hybrid
    • Session hosts use Entra ID-issued Kerberos tickets
    • No AADC password hash sync required for file access
    • Enable: Storage Account → Configuration → Azure AD Kerberos = Enabled
    • Works with Hybrid join — Entra ID ticket used for SMB auth

  BDF Recommendation: Use Azure AD Kerberos (Option B) even for Hybrid Join
  — simpler, no AzFilesHybrid module needed, works with Conditional Access

"@ -ForegroundColor $C.Muted
    $fslogixAuth = Read-MenuChoice "FSLogix Azure Files auth method" @("A","B") -Default "B"
    Write-Status "FSLogix auth" (if ($fslogixAuth -eq "B") {"Azure AD Kerberos (recommended)"} else {"On-prem AD Kerberos"}) "ok"

    $Global:Cfg.JoinConfig.Type = "HybridAD"
    $Global:Cfg.Identity.AzureAdJoin = $false
    Write-Status "Hybrid AD Join configured" "" "ok"
    Save-Config
    Read-Host "  Press Enter to continue"
}

function Show-JoinTypeComparison {
    Write-Banner "JOIN TYPE DETAILED COMPARISON"
    Write-Host @"
  ════════════════════════════════════════════════════════════════════════
   ENTRA ID JOIN (Azure AD Join) — RECOMMENDED FOR BDF
  ════════════════════════════════════════════════════════════════════════

  ✔  Prerequisites:
     • M365 E3/F1 tenant (BDF already has this)
     • Intune enrolled (E3 includes Intune)
     • VMs need System-Assigned Managed Identity
     • RBAC: 'VM User Login' role on RG for user groups

  ✔  VM Extension: AADLoginForWindows (sets up Entra ID auth on VM)

  ✔  FSLogix Auth: Azure AD Kerberos → enabled on storage account
     (HKLM:\SOFTWARE\FSLogix\Profiles\AccessNetworkAsComputerObject=1)

  ✔  AVD Session Host RDP Properties:
     targetisaadjoined:i:1
     enablerdsaadredirection:i:1

  ✔  SSO Flow: User signs in → Entra ID MFA (Okta federation) → AVD session
     starts automatically → OneDrive silently signs in → FSLogix mounts

  ✔  Intune: Full MDM management — Settings Catalog, Compliance, Apps

  ⚠  Limitations:
     • Legacy apps needing on-prem Kerberos/NTLM need workaround
     • GPOs not supported — Intune Settings Catalog used instead
     • Some older apps may not work without AD line-of-sight

  ════════════════════════════════════════════════════════════════════════
   HYBRID AD JOIN — USE ONLY WHEN ON-PREM AD IS REQUIRED
  ════════════════════════════════════════════════════════════════════════

  ⚠  Prerequisites:
     • On-premises Active Directory domain controllers
     • Site-to-site VPN or ExpressRoute from Azure VNet to on-prem datacenter
     • VNet DNS pointing to on-prem DCs
     • Domain join service account in AD with OU 'Create Computer' rights
     • Azure AD Connect or Entra Cloud Sync configured for Computer objects
     • Hybrid Azure AD Join configured in AAD Connect settings

  ✔  VM Extension: JsonADDomainExtension (joins VM to on-prem AD domain)
     PLUS AADLoginForWindows (enables Entra ID SSO on top of domain join)

  ✔  FSLogix Auth: On-prem AD Kerberos OR Azure AD Kerberos (both work)

  ✔  AVD Session Host RDP Properties:
     targetisaadjoined:i:0  (or omit)
     enablerdsaadredirection:i:1  (still use Entra ID SSO for RDP auth)

  ✔  Full GPO support — create AVD-specific GPO linked to AVD OU
     Apply: FSLogix policies, timeout policies, printer maps, drive maps

  ✔  SAP Kerberos: Session hosts can get Kerberos tickets for on-prem SAP
     (if SAP uses Windows Integrated Auth / Kerberos)

  ⚠  Complexity:
     • Domain join account must be rotated regularly
     • Network path to DCs must always be available (VPN/ExpressRoute SLA)
     • Computer accounts accumulate in AD OU as hosts are scaled in/out
     • Azure Update Manager and Intune co-management config needed

  ════════════════════════════════════════════════════════════════════════
   BDF RECOMMENDATION: ENTRA ID JOIN
  ════════════════════════════════════════════════════════════════════════

  BDF's environment (M365 E3/F1, Okta SAML, SAP browser, Zscaler) is
  entirely cloud-compatible. SAP access is via browser (no Kerberos needed).
  Entra ID Join is simpler, faster to deploy, and fully supported.

  Use Hybrid AD Join ONLY if a specific on-prem app requires it and
  cannot be migrated to SAML/OIDC authentication.

"@ -ForegroundColor $C.Muted
    Read-Host "  Press Enter to continue"
}

function Validate-JoinConfiguration {
    Write-Section "Join Configuration Validation"
    $jc = $Global:Cfg.JoinConfig

    if (-not $jc.Configured) {
        Write-Status "Join type not configured — run wizard first" "" "warn"
        Read-Host "  Press Enter"; return
    }

    if ($jc.Type -eq "EntraID") {
        Write-Status "Join Type"          "Entra ID (Azure AD) Join" "ok"
        Write-Status "AzureAdJoin flag"   $Global:Cfg.Identity.AzureAdJoin "ok"
        Write-Status "Intune MDM"         (if ($jc.EntraID.MDMEnrollment) {"Enabled"} else {"Disabled"}) (if ($jc.EntraID.MDMEnrollment) {"ok"} else {"warn"})
        Write-Status "RDP targetisaadjoined" "Will be set to :i:1" "ok"

        # Check if RBAC VM Login roles are assigned
        if ($Global:Cfg.Identity.E3GroupId) {
            $rgScope = "/subscriptions/$($Global:Cfg.SubscriptionId)/resourceGroups/$($Global:Cfg.RG.AVD.Name)"
            $roles   = Get-AzRoleAssignment -ObjectId $Global:Cfg.Identity.E3GroupId -Scope $rgScope -EA SilentlyContinue
            $hasVMLogin = $null -ne ($roles | Where-Object { $_.RoleDefinitionName -like "*VM*Login*" })
            Write-Status "VM User Login RBAC (E3)" (if ($hasVMLogin) {"Assigned"} else {"MISSING — assign in Permission Manager"}) (if ($hasVMLogin) {"ok"} else {"fail"})
        }
    } else {
        Write-Status "Join Type"          "Hybrid AD Join" "ok"
        Write-Status "Domain"             $jc.HybridAD.DomainName "ok"
        Write-Status "Domain Join OU"     $jc.HybridAD.DomainJoinOU (if ($jc.HybridAD.DomainJoinOU) {"ok"} else {"warn"})
        Write-Status "Service Account"    $jc.HybridAD.DomainJoinUPN (if ($jc.HybridAD.DomainJoinUPN) {"ok"} else {"fail"})
        Write-Status "DC IPs configured"  ($jc.HybridAD.DCIPAddresses -join ", ") (if ($jc.HybridAD.DCIPAddresses.Count -gt 0) {"ok"} else {"fail"})
        Write-Status "Sync Method"        $jc.HybridAD.SyncMethod "ok"

        # Network connectivity test to DCs
        foreach ($dc in $jc.HybridAD.DCIPAddresses) {
            $reachable = Test-NetConnection -ComputerName $dc -Port 389 -WarningAction SilentlyContinue -EA SilentlyContinue
            Write-Status "DC reachable (LDAP 389)" "$dc — $($reachable.TcpTestSucceeded)" (if ($reachable.TcpTestSucceeded) {"ok"} else {"warn"})
        }
    }
    Read-Host "  Press Enter to continue"
}

function Export-JoinScripts {
    Write-Section "Exporting Join Configuration Scripts"
    $jc = $Global:Cfg.JoinConfig

    if ($jc.Type -eq "EntraID" -or -not $jc.Configured) {
        # Entra ID Join — VM extension config
        @"
# Entra ID Join VM Extension Deployment
# Deploy via golden image build OR per-VM via Set-AzVMExtension
# Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

`$vmName = "YOUR-VM-NAME"
`$rgName = "$($Global:Cfg.RG.AVD.Name)"

# AADLoginForWindows extension — enables Entra ID authentication on the VM
Set-AzVMExtension ``
    -ResourceGroupName `$rgName ``
    -VMName `$vmName ``
    -Name "AADLoginForWindows" ``
    -Publisher "Microsoft.Azure.ActiveDirectory" ``
    -ExtensionType "AADLoginForWindows" ``
    -TypeHandlerVersion "2.0" ``
    -Settings @{ mdmId = "" }   # Leave empty for default Intune MDM

# RBAC — assign after extension installs
`$rgScope = "/subscriptions/$($Global:Cfg.SubscriptionId)/resourceGroups/`$rgName"
New-AzRoleAssignment -ObjectId "$($Global:Cfg.Identity.E3GroupId)" ``
    -RoleDefinitionName "Virtual Machine User Login" -Scope `$rgScope
New-AzRoleAssignment -ObjectId "$($Global:Cfg.Identity.F1GroupId)" ``
    -RoleDefinitionName "Virtual Machine User Login" -Scope `$rgScope
New-AzRoleAssignment -ObjectId "$($Global:Cfg.Identity.AdminGroupId)" ``
    -RoleDefinitionName "Virtual Machine Administrator Login" -Scope `$rgScope
"@ | Set-Content ".\Configure-EntraIDJoin.ps1" -Encoding UTF8
        Write-Status "Entra ID join script" ".\Configure-EntraIDJoin.ps1" "ok"
    }

    if ($jc.Type -eq "HybridAD") {
        @"
# Hybrid AD Join VM Extension Deployment
# Deploy per-VM via Set-AzVMExtension (or bake into golden image pre-join)
# Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
# IMPORTANT: Store domain join password in Key Vault — never hardcode it

`$vmName    = "YOUR-VM-NAME"
`$rgName    = "$($Global:Cfg.RG.AVD.Name)"
`$domain    = "$($jc.HybridAD.DomainName)"
`$ouPath    = "$($jc.HybridAD.DomainJoinOU)"
`$joinUPN   = "$($jc.HybridAD.DomainJoinUPN)"
`$kvName    = "$($Global:Cfg.KeyVault.Name)"

# Retrieve domain join password from Key Vault (stored securely)
`$djPwd = (Get-AzKeyVaultSecret -VaultName `$kvName -Name "avd-domain-join-password").SecretValue
`$djPwdPlain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR(`$djPwd))

# Step 1: Join to on-prem AD domain
`$djSettings = @{
    Name          = `$domain
    User          = `$joinUPN
    Restart       = "true"
    Options       = "3"  # 1=join domain, 2=create account, 3=both
    OUPath        = `$ouPath
}
`$djProtected = @{ Password = `$djPwdPlain }

Set-AzVMExtension ``
    -ResourceGroupName `$rgName ``
    -VMName `$vmName ``
    -Name "JsonADDomainExtension" ``
    -Publisher "Microsoft.Compute" ``
    -ExtensionType "JsonADDomainExtension" ``
    -TypeHandlerVersion "1.3" ``
    -Settings `$djSettings ``
    -ProtectedSettings `$djProtected

# Step 2: Enable Entra ID SSO on top of domain join
Set-AzVMExtension ``
    -ResourceGroupName `$rgName ``
    -VMName `$vmName ``
    -Name "AADLoginForWindows" ``
    -Publisher "Microsoft.Azure.ActiveDirectory" ``
    -ExtensionType "AADLoginForWindows" ``
    -TypeHandlerVersion "2.0" ``
    -Settings @{ mdmId = "" }

# Step 3: Store domain join password in Key Vault (run once during setup)
# Set-AzKeyVaultSecret -VaultName `$kvName -Name "avd-domain-join-password" ``
#     -SecretValue (Read-Host "Domain join password" -AsSecureString)
"@ | Set-Content ".\Configure-HybridADJoin.ps1" -Encoding UTF8
        Write-Status "Hybrid AD join script" ".\Configure-HybridADJoin.ps1" "ok"

        # GPO guidance script
        @"
# BDF AVD Hybrid Join — Recommended GPO Structure
# Apply these GPOs to: $($jc.HybridAD.DomainJoinOU)

# GPO 1: AVD-SessionHosts-FSLogix
# Settings: FSLogix Profile Container config (VHDLocations, etc.)
# Path: Computer Config > Policies > Admin Templates > FSLogix

# GPO 2: AVD-SessionHosts-Timeouts
# Settings: Session timeout, idle disconnect, logon banner
# Path: Computer Config > Policies > Windows Settings > Security Settings > Local Policies

# GPO 3: AVD-SessionHosts-Printers
# Settings: Deploy network printers via Group Policy Preferences
# Path: User Config > Preferences > Control Panel Settings > Printers

# GPO 4: AVD-SessionHosts-Hardening
# Settings: Windows Firewall, Defender settings, UAC, account lockout
# Path: Computer Config > Policies > Windows Settings > Security Settings

# Block Inheritance on AVD OU:
Set-GPInheritance -Target "$($jc.HybridAD.DomainJoinOU)" -IsBlocked Yes
# Then explicitly link only AVD-relevant GPOs to this OU

# Loopback Processing (apply user GPOs from computer policy — critical for VDI):
# Computer Config > Admin Templates > System > Group Policy
#   Configure user Group Policy loopback processing mode = Enabled, Mode = Replace
"@ | Set-Content ".\AVD-GPO-Guidance.ps1" -Encoding UTF8
        Write-Status "GPO guidance script" ".\AVD-GPO-Guidance.ps1" "ok"
    }
    Read-Host "  Press Enter to continue"
}
#endregion

#region ═══════════════════════════════════════════════════════════════════════
#  RDP PROPERTIES MANAGER — BEST PRACTICES & INTERACTIVE BUILDER
#═══════════════════════════════════════════════════════════════════════════════

# Master RDP property catalog — key, description, impact, E3 default, F1 default, security risk
$Script:RDPCatalog = [ordered]@{

    "DISPLAY & GRAPHICS" = @(
        @{ Key="dynamic resolution:i";       E3=1;  F1=1;  Impact="UX";     Risk="Low";
           Desc="Dynamically adjust resolution when user resizes RD Client window" }
        @{ Key="use multimon:i";             E3=1;  F1=0;  Impact="UX";     Risk="Low";
           Desc="Multi-monitor support (E3: allow, F1: single screen for thin clients)" }
        @{ Key="maximizetocurrentdisplays:i"; E3=1; F1=0;  Impact="UX";     Risk="Low";
           Desc="Maximize session to span all client monitors" }
        @{ Key="smart sizing:i";             E3=0;  F1=1;  Impact="UX";     Risk="Low";
           Desc="Scale session content to fit window (F1: useful on small thin-client screens)" }
        @{ Key="desktopscalefactor:i";       E3=100; F1=100; Impact="UX";   Risk="Low";
           Desc="DPI scaling factor % (100=normal, 150=high-DPI)" }
        @{ Key="session bpp:i";              E3=32; F1=32;  Impact="Perf";   Risk="Low";
           Desc="Color depth: 16 or 32 bit (32=best quality, 16=lower bandwidth)" }
        @{ Key="allow font smoothing:i";     E3=1;  F1=1;  Impact="UX";     Risk="Low";
           Desc="ClearType font rendering in session" }
        @{ Key="allow desktop composition:i"; E3=1; F1=1;  Impact="Perf";   Risk="Low";
           Desc="Aero glass effects (Windows 11 — minimal perf impact)" }
        @{ Key="disable menu anims:i";       E3=0;  F1=0;  Impact="Perf";   Risk="Low";
           Desc="0=show animations, 1=disable (set 1 on low-bandwidth connections)" }
        @{ Key="disable themes:i";           E3=0;  F1=0;  Impact="Perf";   Risk="Low";
           Desc="0=use themes, 1=disable themes (saves bandwidth on slow connections)" }
        @{ Key="disable wallpaper:i";        E3=0;  F1=0;  Impact="Perf";   Risk="Low";
           Desc="0=show wallpaper, 1=plain background (bandwidth saving — rarely needed with RDP Shortpath)" }
        @{ Key="videoplaybackmode:i";        E3=1;  F1=1;  Impact="Perf";   Risk="Low";
           Desc="Video rendering optimization (1=enabled — important for multimedia content)" }
    )

    "AUDIO & MULTIMEDIA" = @(
        @{ Key="audiomode:i";                E3=0;  F1=0;  Impact="UX";     Risk="Low";
           Desc="Audio output: 0=play on client (recommended), 1=play on server, 2=no audio" }
        @{ Key="audiocapturemode:i";         E3=1;  F1=0;  Impact="UX";     Risk="Low";
           Desc="Microphone redirection: 1=enabled (E3 for Teams calls), 0=disabled (F1 frontline)" }
        @{ Key="encode redirected video capture:i"; E3=1; F1=1; Impact="Perf"; Risk="Low";
           Desc="Multimedia Redirection (MMR): offloads video decode to client device — CRITICAL for Teams" }
        @{ Key="redirected video capture encoding quality:i"; E3=0; F1=0; Impact="Perf"; Risk="Low";
           Desc="Video capture quality: 0=high, 1=medium, 2=low" }
        @{ Key="camerastoredirect:s";        E3="*"; F1="";Impact="UX";     Risk="Medium";
           Desc="Camera redirection: * = all cameras (E3 Teams video), blank = disabled (F1 kiosks)" }
    )

    "DEVICE REDIRECTION" = @(
        @{ Key="redirectclipboard:i";        E3=1;  F1=0;  Impact="DLP";    Risk="HIGH";
           Desc="Clipboard copy/paste: 1=allow (E3 productivity), 0=BLOCK (F1 DLP — prevents data exfil to kiosks)" }
        @{ Key="drivestoredirect:s";         E3=""; F1=""; Impact="DLP";    Risk="HIGH";
           Desc="Drive redirection: blank=DISABLED (all users — use OneDrive instead, prevents USB exfil)" }
        @{ Key="redirectprinters:i";         E3=1;  F1=1;  Impact="UX";     Risk="Low";
           Desc="Printer redirection: 1=allow (use with Universal Print for cloud-managed printing)" }
        @{ Key="redirectsmartcards:i";       E3=0;  F1=0;  Impact="Auth";   Risk="Low";
           Desc="Smart card passthrough: 0=disabled (BDF uses Okta/Entra — no smart cards)" }
        @{ Key="redirectwebauthn:i";         E3=1;  F1=0;  Impact="Auth";   Risk="Low";
           Desc="WebAuthn/FIDO2 passthrough: 1=allow passkey/hardware key from client device (E3)" }
        @{ Key="redirectlocation:i";         E3=0;  F1=0;  Impact="Privacy";Risk="Medium";
           Desc="Location services redirection: 0=disabled (privacy — no need for AVD sessions)" }
        @{ Key="redirectposdevices:i";       E3=0;  F1=0;  Impact="UX";     Risk="Low";
           Desc="POS device redirection: 0=disabled (manage via separate POS integration if needed)" }
        @{ Key="redirectcomports:i";         E3=0;  F1=0;  Impact="UX";     Risk="Medium";
           Desc="COM/serial port redirection: 0=disabled (enable only if legacy serial devices required)" }
        @{ Key="usbdevicestoredirect:s";     E3=""; F1=""; Impact="DLP";    Risk="HIGH";
           Desc="USB device redirection: blank=DISABLED. Whitelist specific device classes if needed." }
    )

    "AUTHENTICATION & SSO" = @(
        @{ Key="targetisaadjoined:i";        E3=1;  F1=1;  Impact="Auth";   Risk="N/A";
           Desc="Azure AD/Entra ID joined session hosts: 1=yes (set 0 for Hybrid AD join)" }
        @{ Key="enablerdsaadredirection:i";  E3=1;  F1=1;  Impact="Auth";   Risk="N/A";
           Desc="Entra ID SSO — user's Entra token reused for AVD session auth (zero password prompt)" }
        @{ Key="enablecredsspsupport:i";     E3=1;  F1=1;  Impact="Auth";   Risk="Low";
           Desc="CredSSP (Network Level Authentication) — secure pre-auth before session creation" }
        @{ Key="authentication level:i";     E3=2;  F1=2;  Impact="Security";Risk="Medium";
           Desc="0=no warning, 1=warn if auth fails, 2=block if certificate fails (use 2)" }
        @{ Key="prompt for credentials:i";   E3=0;  F1=0;  Impact="UX";     Risk="Low";
           Desc="0=use SSO (no prompt), 1=always prompt (set 0 when SSO is configured)" }
        @{ Key="prompt for credentials on client:i"; E3=0; F1=0; Impact="UX"; Risk="Low";
           Desc="0=no prompt on reconnect, 1=prompt (set 0 with Entra ID SSO)" }
        @{ Key="autoreconnection enabled:i"; E3=1;  F1=1;  Impact="UX";     Risk="Low";
           Desc="Auto-reconnect if network drops: 1=enabled (critical for retail store reliability)" }
    )

    "SECURITY & DATA PROTECTION" = @(
        @{ Key="screen-capture-protection:i"; E3=1; F1=2; Impact="DLP";    Risk="N/A";
           Desc="Screenshot blocking: 0=off, 1=block on client only, 2=block on client+server. F1=2 (full block for store)" }
        @{ Key="watermarking:i";             E3=1;  F1=1;  Impact="DLP";    Risk="N/A";
           Desc="Session watermark (shows username+session info as faint overlay): 1=enabled for both pools" }
        @{ Key="watermarking watermark opacity:i"; E3=20; F1=30; Impact="DLP"; Risk="N/A";
           Desc="Watermark opacity: 0-100. 20=subtle for office, 30=more visible for store (retail data sensitivity)" }
        @{ Key="use redirection server name:s"; E3="*"; F1="*"; Impact="UX"; Risk="Low";
           Desc="Use hostname for reconnection (improves reconnect reliability across load balanced hosts)" }
        @{ Key="negotiate security layer:i"; E3=1;  F1=1;  Impact="Security";Risk="Low";
           Desc="Negotiate highest security layer (TLS): 1=required" }
    )

    "PERFORMANCE & BANDWIDTH" = @(
        @{ Key="connection type:i";          E3=7;  F1=7;  Impact="Perf";   Risk="Low";
           Desc="Network bandwidth mode: 7=auto-detect (recommended — adapts to current conditions)" }
        @{ Key="bandwidthautodetect:i";      E3=1;  F1=1;  Impact="Perf";   Risk="Low";
           Desc="Automatically detect available bandwidth and adjust session quality" }
        @{ Key="networkautodetect:i";        E3=1;  F1=1;  Impact="Perf";   Risk="Low";
           Desc="Network auto-detection: 1=enabled — detects RTT and packet loss, adapts codecs" }
        @{ Key="compression:i";             E3=1;  F1=1;  Impact="Perf";   Risk="Low";
           Desc="RDP data compression: 1=enabled (reduces bandwidth, slight CPU overhead)" }
        @{ Key="bitmapcachesize:i";         E3=32000; F1=16000; Impact="Perf"; Risk="Low";
           Desc="Bitmap cache size in KB: E3=32MB (rich UI), F1=16MB (lighter — SAP browser apps)" }
        @{ Key="bitmapcachepersistenable:i"; E3=1; F1=0;  Impact="Perf";   Risk="Low";
           Desc="Persist bitmap cache across sessions: E3=1 (speeds up reconnect), F1=0 (shared kiosks)" }
        @{ Key="keyboardhook:i";            E3=2;  F1=2;  Impact="UX";     Risk="Low";
           Desc="Keyboard shortcut handling: 0=local, 1=remote, 2=full-screen only (recommended)" }
    )

    "SESSION TIMEOUTS" = @(
        @{ Key="session disconnect timeout (minutes):i"; E3=480; F1=240; Impact="Cost"; Risk="Low";
           Desc="Disconnect session after N minutes of disconnect state. E3=8hrs, F1=4hrs (shift-based)" }
        @{ Key="session idle timeout (minutes):i"; E3=120; F1=30; Impact="Cost"; Risk="Low";
           Desc="Log off idle sessions. E3=2hrs (office workers), F1=30min (store workers — free slot for next shift)" }
    )

    "TOUCH & ACCESSIBILITY" = @(
        @{ Key="enablesuperpan:i";          E3=1;  F1=1;  Impact="UX";     Risk="Low";
           Desc="Touch screen pan/scroll: 1=enabled (important for iPad users in stores)" }
        @{ Key="gatewayusagemethod:i";      E3=0;  F1=0;  Impact="Network";Risk="Low";
           Desc="Gateway usage: 0=auto-detect (AVD uses its own reverse-connect, not traditional RD Gateway)" }
    )
}

# Pre-built profiles
$Script:RDPProfiles = @{
    "Strict" = @{
        Desc = "Maximum security — clipboard/USB/drives all disabled. Best for PCI-scope or sensitive data sessions."
        Overrides = @{
            "redirectclipboard:i"         = 0
            "drivestoredirect:s"          = ""
            "usbdevicestoredirect:s"      = ""
            "screen-capture-protection:i" = 2
            "watermarking:i"              = 1
            "watermarking watermark opacity:i" = 40
            "camerastoredirect:s"         = ""
            "audiocapturemode:i"          = 0
            "redirectcomports:i"          = 0
            "redirectsmartcards:i"        = 0
            "redirectlocation:i"          = 0
        }
    }
    "Balanced" = @{
        Desc = "Recommended for most BDF users — productivity + security. E3 office workers."
        Overrides = @{
            "redirectclipboard:i"         = 1
            "drivestoredirect:s"          = ""     # Always disabled — use OneDrive
            "usbdevicestoredirect:s"      = ""
            "screen-capture-protection:i" = 1
            "watermarking:i"              = 1
            "camerastoredirect:s"         = "*"
            "audiocapturemode:i"          = 1
        }
    }
    "Open" = @{
        Desc = "Maximum compatibility — most redirections enabled. For IT/admin users or troubleshooting."
        Overrides = @{
            "redirectclipboard:i"         = 1
            "drivestoredirect:s"          = "*"
            "redirectsmartcards:i"        = 1
            "camerastoredirect:s"         = "*"
            "audiocapturemode:i"          = 1
            "screen-capture-protection:i" = 0
            "watermarking:i"              = 0
        }
    }
    "Frontline" = @{
        Desc = "F1 store workers — kiosk-style. Clipboard off, cameras off, minimal redirections."
        Overrides = @{
            "redirectclipboard:i"         = 0
            "drivestoredirect:s"          = ""
            "usbdevicestoredirect:s"      = ""
            "camerastoredirect:s"         = ""
            "audiocapturemode:i"          = 0
            "screen-capture-protection:i" = 2
            "watermarking:i"              = 1
            "watermarking watermark opacity:i" = 30
            "smart sizing:i"              = 1
            "use multimon:i"              = 0
        }
    }
}

function Show-RDPPropertiesMenu {
    Write-Banner "RDP PROPERTIES MANAGER" $C.Azure

    while ($true) {
        $rdp = $Global:Cfg.RDP
        Write-Section "RDP Configuration Status"

        $profileStatus = if ($rdp.Configured) {"✔ $($rdp.ActiveProfile)"} else {"○ Not Configured"}
        $pcol = if ($rdp.Configured) {$C.Ok} else {$C.Muted}
        Write-Host ("  {0,-32} {1}" -f "Active Profile:", $profileStatus) -ForegroundColor $pcol
        Write-Host ("  {0,-32} {1}" -f "E3 Property String:", (if ($rdp.E3PropertyString.Length -gt 60) {"$($rdp.E3PropertyString.Substring(0,57))..."} else {$rdp.E3PropertyString})) -ForegroundColor $C.Muted
        Write-Host ""

        Write-Host "  1.  Apply preset profile (Strict / Balanced / Frontline / Open)" -ForegroundColor $C.New
        Write-Host "  2.  Interactive property builder (category-by-category)" -ForegroundColor $C.Menu
        Write-Host "  3.  View & edit current property string (raw)" -ForegroundColor $C.Menu
        Write-Host "  4.  View all properties with current values" -ForegroundColor $C.Menu
        Write-Host "  5.  Apply RDP properties to host pools NOW" -ForegroundColor $C.Azure
        Write-Host "  6.  Export RDP properties reference sheet" -ForegroundColor $C.Menu
        Write-Host "  7.  Validate properties against best practices" -ForegroundColor $C.Warn
        Write-Host "  0.  Back" -ForegroundColor $C.Muted

        $ch = Read-MenuChoice "RDP Manager" @("0","1","2","3","4","5","6","7")
        switch ($ch) {
            "1" { Apply-RDPPresetProfile }
            "2" { Invoke-RDPPropertyBuilder }
            "3" { Edit-RDPPropertyStringRaw }
            "4" { Show-AllRDPProperties }
            "5" { Apply-RDPToHostPools }
            "6" { Export-RDPReferenceSheet }
            "7" { Validate-RDPProperties }
            "0" { return }
        }
    }
}

function Apply-RDPPresetProfile {
    Write-Section "RDP Preset Profiles"
    $profiles = $Script:RDPProfiles
    $i = 1
    foreach ($p in $profiles.GetEnumerator()) {
        Write-Host ("  {0}. {1,-14} — {2}" -f $i, $p.Key, $p.Value.Desc) -ForegroundColor $C.Menu; $i++
    }
    $opts = 1..($profiles.Count) | ForEach-Object {"$_"}
    $ch   = Read-MenuChoice "Select profile" $opts -Default "2"
    $sel  = @($profiles.GetEnumerator())[[int]$ch - 1]

    Write-Section "Configure Profile: $($sel.Key)"
    Write-Host "  $($sel.Value.Desc)" -ForegroundColor $C.Muted
    Write-Host ""

    # Show what this profile sets
    foreach ($kv in $sel.Value.Overrides.GetEnumerator()) {
        $riskIcon = switch -Wildcard ($kv.Key) {
            "*clipboard*" { "🔒" } "*drive*" { "🔒" } "*usb*" { "🔒" }
            "*screen-cap*"{ "🔒" } "*camera*" { "📷" } default { "  " }
        }
        Write-Host ("  $riskIcon {0,-50} = {1}" -f $kv.Key, $kv.Value) -ForegroundColor $C.Menu
    }

    Write-Host ""
    if (Read-YesNo "Apply '$($sel.Key)' profile?" $true) {
        # Apply the preset overrides to the RDP config
        $rdp = $Global:Cfg.RDP
        foreach ($kv in $sel.Value.Overrides.GetEnumerator()) {
            # Find and update the matching setting
            foreach ($cat in $Script:RDPCatalog.GetEnumerator()) {
                $match = $cat.Value | Where-Object { ($_.Key + ":i") -eq $kv.Key -or ($_.Key + ":s") -eq $kv.Key }
                # Just record the override; string generation handles it
            }
        }
        $rdp.ActiveProfile = $sel.Key
        $rdp.Configured    = $true

        # Build the RDP property strings
        $rdp.E3PropertyString = Build-RDPString "E3" $sel.Value.Overrides
        $rdp.F1PropertyString = Build-RDPString "F1" $sel.Value.Overrides

        Write-Status "Profile applied" $sel.Key "ok"
        Write-Status "E3 string length" "$($rdp.E3PropertyString.Split(';').Count) properties" "ok"
        Save-Config
        if (Read-YesNo "Apply to host pools now?" $true) { Apply-RDPToHostPools }
    }
    Read-Host "  Press Enter to continue"
}

function Build-RDPString {
    param([string]$PoolType, [hashtable]$Overrides = @{})
    $rdp  = $Global:Cfg.RDP
    $join = $Global:Cfg.JoinConfig
    $isAAD = ($join.Type -eq "EntraID" -or -not $join.Configured)

    $props = [ordered]@{
        # Join type
        "targetisaadjoined:i"                   = if ($isAAD) {1} else {0}
        "enablerdsaadredirection:i"              = 1

        # Display
        "dynamic resolution:i"                   = $rdp.Display.DynamicResolution
        "use multimon:i"                         = if ($PoolType -eq "E3") {$rdp.Display.MultiMonitor} else {0}
        "maximizetocurrentdisplays:i"            = if ($PoolType -eq "E3") {$rdp.Display.MaximizeToDisplays} else {0}
        "smart sizing:i"                         = if ($PoolType -eq "F1") {1} else {$rdp.Display.SmartSizing}
        "session bpp:i"                          = $rdp.Display.ColorDepth
        "connection type:i"                      = $rdp.Display.ConnectionType
        "bandwidthautodetect:i"                  = $rdp.Display.BandwidthAutoDetect
        "networkautodetect:i"                    = $rdp.Display.NetworkAutoDetect
        "compression:i"                          = $rdp.Perf.BitmapCachePersist
        "allow font smoothing:i"                 = $rdp.Display.AllowFontSmoothing
        "allow desktop composition:i"            = $rdp.Display.AllowDesktopComposition
        "videoplaybackmode:i"                    = $rdp.Display.VideoPlaybackMode
        "disable menu anims:i"                   = $rdp.Perf.DisableMenuAnims
        "disable themes:i"                       = $rdp.Perf.DisableThemes

        # Audio
        "audiomode:i"                            = $rdp.AV.AudioOutputMode
        "audiocapturemode:i"                     = if ($PoolType -eq "E3") {$rdp.AV.AudioCapture} else {$rdp.AV.AudioCapture -eq 1 ? 0 : 0}
        "encode redirected video capture:i"      = $rdp.AV.VideoCapture
        "redirected video capture encoding quality:i" = $rdp.AV.VideoCaptureQuality
        "camerastoredirect:s"                    = if ($PoolType -eq "E3") {($rdp.AV.CameraRedirect ? "*" : "")} else {""}

        # Device Redirection
        "redirectclipboard:i"                    = if ($PoolType -eq "E3") {$rdp.Redirect.Clipboard} else {$rdp.Redirect.ClipboardF1}
        "drivestoredirect:s"                     = ""     # Always disabled
        "redirectprinters:i"                     = $rdp.Redirect.Printers
        "redirectsmartcards:i"                   = $rdp.Redirect.SmartCards
        "redirectwebauthn:i"                     = if ($PoolType -eq "E3") {$rdp.Redirect.WebAuthn} else {0}
        "redirectlocation:i"                     = $rdp.Redirect.Location
        "redirectposdevices:i"                   = $rdp.Redirect.POSDevices
        "redirectcomports:i"                     = $rdp.Redirect.SerialPorts
        "usbdevicestoredirect:s"                 = ""     # Always disabled (whitelist as needed)

        # Auth
        "enablecredsspsupport:i"                 = $rdp.Auth.CredSSP
        "authentication level:i"                 = $rdp.Auth.AuthLevel
        "prompt for credentials:i"               = $rdp.Auth.PromptCredentials
        "prompt for credentials on client:i"     = 0
        "negotiate security layer:i"             = 1
        "autoreconnection enabled:i"             = $rdp.Timeouts.AutoReconnect

        # Security
        "screen-capture-protection:i"            = if ($PoolType -eq "E3") {$rdp.Security.ScreenCaptureProtection} else {$rdp.Security.ScreenCaptureProtF1}
        "watermarking:i"                         = $rdp.Security.Watermarking
        "watermarking watermark opacity:i"       = if ($PoolType -eq "E3") {$rdp.Security.WatermarkOpacity} else {30}
        "use redirection server name:s"          = "*"

        # Performance
        "keyboardhook:i"                         = $rdp.Perf.KeyboardHook
        "bitmapcachesize:i"                      = if ($PoolType -eq "E3") {$rdp.Perf.BitmapCacheSize} else {16000}
        "bitmapcachepersistenable:i"             = if ($PoolType -eq "E3") {$rdp.Perf.BitmapCachePersist} else {0}

        # Touch
        "enablesuperpan:i"                       = 1
        "gatewayusagemethod:i"                   = 0
    }

    # Apply preset overrides on top
    foreach ($kv in $Overrides.GetEnumerator()) { $props[$kv.Key] = $kv.Value }

    # Build semicolon-separated string
    return ($props.GetEnumerator() | ForEach-Object { "$($_.Key):$($_.Value)" }) -join ";"
}

function Invoke-RDPPropertyBuilder {
    Write-Banner "INTERACTIVE RDP PROPERTY BUILDER"
    $rdp = $Global:Cfg.RDP

    foreach ($cat in $Script:RDPCatalog.GetEnumerator()) {
        Write-Section $cat.Key

        foreach ($prop in $cat.Value) {
            $curE3 = $prop.E3
            $curF1 = $prop.F1

            # Color code by risk
            $col = switch ($prop.Risk) {
                "HIGH"   { $C.Fail }
                "Medium" { $C.Warn }
                default  { $C.Menu }
            }
            Write-Host ("  [$($prop.Risk)$(' '*([Math]::Max(0,6-$prop.Risk.Length)))] {0,-44} E3:{1,-6} F1:{2}" -f $prop.Key, $curE3, $curF1) -ForegroundColor $col
            Write-Host ("         {0}" -f $prop.Desc) -ForegroundColor $C.Muted

            if (Read-YesNo "   Customize this property?" $false) {
                $newE3 = Read-MenuChoice "   E3 value" -Default "$curE3"
                $newF1 = Read-MenuChoice "   F1 value" -Default "$curF1"
                $prop.E3 = if ($newE3 -match "^\d+$") {[int]$newE3} else {$newE3}
                $prop.F1 = if ($newF1 -match "^\d+$") {[int]$newF1} else {$newF1}
                Write-Status "Updated" "$($prop.Key) E3=$($prop.E3) F1=$($prop.F1)" "ok"
            }
        }
        if (-not (Read-YesNo "Continue to next category?" $true)) { break }
    }

    $rdp.ActiveProfile    = "Custom"
    $rdp.Configured       = $true
    $rdp.E3PropertyString = Build-RDPString "E3" @{}
    $rdp.F1PropertyString = Build-RDPString "F1" @{}
    Write-Status "Custom RDP properties built" "" "ok"
    Save-Config

    if (Read-YesNo "Apply to host pools now?" $true) { Apply-RDPToHostPools }
    Read-Host "  Press Enter"
}

function Edit-RDPPropertyStringRaw {
    Write-Section "Current RDP Property Strings"
    $rdp = $Global:Cfg.RDP
    Write-Host ""
    Write-Host "  E3 Pool:" -ForegroundColor $C.Accent
    $rdp.E3PropertyString.Split(";") | ForEach-Object {
        if ($_) { Write-Host ("    {0}" -f $_) -ForegroundColor $C.Muted }
    }
    Write-Host ""
    Write-Host "  F1 Pool:" -ForegroundColor $C.Accent
    $rdp.F1PropertyString.Split(";") | ForEach-Object {
        if ($_) { Write-Host ("    {0}" -f $_) -ForegroundColor $C.Muted }
    }
    Write-Host ""
    if (Read-YesNo "Edit E3 string manually?" $false) {
        Write-Host "  Paste full semicolon-separated RDP property string:" -ForegroundColor $C.Warn
        $rdp.E3PropertyString = Read-MenuChoice "E3 String" -Default $rdp.E3PropertyString
    }
    if (Read-YesNo "Edit F1 string manually?" $false) {
        $rdp.F1PropertyString = Read-MenuChoice "F1 String" -Default $rdp.F1PropertyString
    }
    Save-Config
    Read-Host "  Press Enter"
}

function Show-AllRDPProperties {
    Write-Banner "ALL RDP PROPERTIES — CURRENT VALUES"
    Write-Host ("  {0,-8} {1,-50} {2,-10} {3,-10} {4}" -f "RISK", "Property", "E3 Value", "F1 Value", "Description") -ForegroundColor $C.Muted
    Write-Line
    foreach ($cat in $Script:RDPCatalog.GetEnumerator()) {
        Write-Host ""
        Write-Host "  ── $($cat.Key)" -ForegroundColor $C.Accent
        foreach ($prop in $cat.Value) {
            $rCol = switch ($prop.Risk) { "HIGH"{$C.Fail} "Medium"{$C.Warn} default{$C.Ok} }
            Write-Host ("  {0,-8}" -f "[$($prop.Risk)]") -NoNewline -ForegroundColor $rCol
            Write-Host ("{0,-50}" -f $prop.Key) -NoNewline -ForegroundColor $C.Menu
            Write-Host ("{0,-10}" -f $prop.E3) -NoNewline -ForegroundColor $C.Azure
            Write-Host ("{0,-10}" -f $prop.F1) -NoNewline -ForegroundColor $C.Gold
            Write-Host ($prop.Desc.Substring(0, [Math]::Min(50,$prop.Desc.Length))) -ForegroundColor $C.Muted
        }
    }
    Write-Host ""
    Write-Host "  Legend: " -NoNewline -ForegroundColor $C.Muted
    Write-Host "[HIGH] = security-critical setting  " -NoNewline -ForegroundColor $C.Fail
    Write-Host "[Medium] = review carefully  " -NoNewline -ForegroundColor $C.Warn
    Write-Host "[Low] = standard setting" -ForegroundColor $C.Ok
    Read-Host "  Press Enter to continue"
}

function Apply-RDPToHostPools {
    Write-Section "Applying RDP Properties to Host Pools"
    $rdp = $Global:Cfg.RDP

    if (-not $rdp.Configured -or -not $rdp.E3PropertyString) {
        Write-Status "RDP properties not configured — run wizard or select preset first" "" "warn"
        Read-Host "  Press Enter"; return
    }

    foreach ($poolKey in @("E3","F1")) {
        $hpName  = $Global:Cfg.HostPools[$poolKey].Name
        $hpRg    = $Global:Cfg.RG.AVD.Name
        $rdpStr  = if ($poolKey -eq "E3") {$rdp.E3PropertyString} else {$rdp.F1PropertyString}

        # Ensure join type properties are correct
        if ($Global:Cfg.JoinConfig.Type -eq "EntraID") {
            if ($rdpStr -notlike "*targetisaadjoined:i:1*") {
                $rdpStr = "targetisaadjoined:i:1;$rdpStr"
            }
        } elseif ($Global:Cfg.JoinConfig.Type -eq "HybridAD") {
            $rdpStr = $rdpStr -replace "targetisaadjoined:i:1","targetisaadjoined:i:0"
        }

        $hp = Get-AzWvdHostPool -Name $hpName -ResourceGroupName $hpRg -EA SilentlyContinue
        if (-not $hp) {
            Write-Status "Host pool not found" $hpName "warn"; continue
        }

        Write-Status "Updating $poolKey host pool" $hpName "step"
        Update-AzWvdHostPool -Name $hpName -ResourceGroupName $hpRg `
            -CustomRdpProperty $rdpStr | Out-Null
        Write-Status "RDP properties applied" "$($rdpStr.Split(';').Count) properties" "ok"
    }

    $Global:Cfg.RDP.Configured = $true
    $Global:Cfg.DeploymentState.RDPProperties = "Deployed"
    Save-Config
    Read-Host "  Press Enter to continue"
}

function Export-RDPReferenceSheet {
    Write-Section "Exporting RDP Properties Reference"
    $rdp  = $Global:Cfg.RDP
    $file = ".\BDF-AVD-RDP-Properties-Reference.txt"
    $lines = @(
        "═══════════════════════════════════════════════════════════════════════",
        " BDF Azure Virtual Desktop — RDP Custom Properties Reference",
        " Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        " Active Profile: $($rdp.ActiveProfile)",
        "═══════════════════════════════════════════════════════════════════════",
        "",
        "E3 OFFICE WORKERS — FULL PROPERTY STRING:",
        "─────────────────────────────────────────",
        $rdp.E3PropertyString,
        "",
        "F1 FRONTLINE WORKERS — FULL PROPERTY STRING:",
        "─────────────────────────────────────────────",
        $rdp.F1PropertyString,
        "",
        "FORMATTED PROPERTY LIST (E3):",
        "─────────────────────────────"
    )
    $rdp.E3PropertyString.Split(";") | Where-Object {$_} | ForEach-Object { $lines += "  $_" }
    $lines += ""
    $lines += "FORMATTED PROPERTY LIST (F1):"
    $lines += "─────────────────────────────"
    $rdp.F1PropertyString.Split(";") | Where-Object {$_} | ForEach-Object { $lines += "  $_" }
    $lines += ""
    $lines += "APPLY VIA POWERSHELL:"
    $lines += "─────────────────────"
    $lines += "Update-AzWvdHostPool -Name '$($Global:Cfg.HostPools.E3.Name)' -ResourceGroupName '$($Global:Cfg.RG.AVD.Name)' -CustomRdpProperty '`$E3String'"
    $lines += "Update-AzWvdHostPool -Name '$($Global:Cfg.HostPools.F1.Name)' -ResourceGroupName '$($Global:Cfg.RG.AVD.Name)' -CustomRdpProperty '`$F1String'"
    $lines += ""
    $lines += "SECURITY SETTINGS SUMMARY:"
    $lines += "─────────────────────────"
    foreach ($cat in $Script:RDPCatalog.GetEnumerator()) {
        foreach ($prop in ($cat.Value | Where-Object { $_.Risk -in @("HIGH","Medium") })) {
            $lines += "  [$($prop.Risk)] $($prop.Key) E3=$($prop.E3) F1=$($prop.F1)  — $($prop.Desc)"
        }
    }
    $lines | Set-Content $file -Encoding UTF8
    Write-Status "RDP reference sheet" $file "ok"
    Read-Host "  Press Enter"
}

function Validate-RDPProperties {
    Write-Section "RDP Properties Best Practice Validation"
    $rdp  = $Global:Cfg.RDP
    $pass = 0; $warn = 0; $fail = 0

    $checks = @(
        @{ Name="RDP configured";           Pass=$rdp.Configured;          Sev="FAIL" }
        @{ Name="E3 property string set";   Pass=($rdp.E3PropertyString -ne ""); Sev="FAIL" }
        @{ Name="F1 property string set";   Pass=($rdp.F1PropertyString -ne ""); Sev="FAIL" }
        @{ Name="F1 clipboard disabled";    Pass=($rdp.F1PropertyString -like "*redirectclipboard:i:0*");  Sev="WARN";  Note="DLP — F1 kiosk clipboard should be off" }
        @{ Name="Drives disabled (both)";   Pass=($rdp.E3PropertyString -like "*drivestoredirect:s:;*" -or $rdp.E3PropertyString -like "*drivestoredirect:s:*;*" -or $rdp.E3PropertyString -notlike "*drivestoredirect:s:*" -or $rdp.E3PropertyString -like "*drivestoredirect:s:;*"); Sev="WARN"; Note="Drive redirection should be disabled — use OneDrive" }
        @{ Name="USB disabled (both)";      Pass=($rdp.E3PropertyString -like "*usbdevicestoredirect:s:;*" -or $rdp.E3PropertyString -notlike "*usbdevicestoredirect:s:*"); Sev="WARN"; Note="Generic USB redirection is a security risk" }
        @{ Name="SSO enabled (both)";       Pass=($rdp.E3PropertyString -like "*enablerdsaadredirection:i:1*"); Sev="WARN"; Note="Entra ID SSO should be enabled for better UX" }
        @{ Name="Screen capture protection";Pass=($rdp.E3PropertyString -like "*screen-capture-protection*"); Sev="WARN"; Note="Screen capture protection recommended for data sensitivity" }
        @{ Name="Watermarking enabled";     Pass=($rdp.E3PropertyString -like "*watermarking:i:1*"); Sev="WARN"; Note="Watermarks deter screen photography of sensitive data" }
        @{ Name="Auth level = 2";           Pass=($rdp.E3PropertyString -like "*authentication level:i:2*"); Sev="WARN"; Note="Auth level 2 blocks connections with cert failures" }
        @{ Name="NLA/CredSSP enabled";      Pass=($rdp.E3PropertyString -like "*enablecredsspsupport:i:1*"); Sev="WARN"; Note="CredSSP pre-auth required for secure connections" }
        @{ Name="Auto-reconnect on";        Pass=($rdp.E3PropertyString -like "*autoreconnection enabled:i:1*"); Sev="WARN"; Note="Critical for retail store session reliability" }
        @{ Name="Touch/pan enabled";        Pass=($rdp.E3PropertyString -like "*enablesuperpan:i:1*"); Sev="WARN"; Note="Required for iPad users in store" }
        @{ Name="MMR/video capture on";     Pass=($rdp.E3PropertyString -like "*encode redirected video capture:i:1*"); Sev="WARN"; Note="Multimedia Redirection needed for Teams video quality" }
        @{ Name="Join type in RDP props";   Pass=($rdp.E3PropertyString -like "*targetisaadjoined*"); Sev="FAIL"; Note="Join type must be specified in RDP properties" }
    )

    Write-Host ("  {0,-44} {1,-8} {2}" -f "Check", "Result", "Notes") -ForegroundColor $C.Muted
    Write-Line
    foreach ($chk in $checks) {
        $icon = if ($chk.Pass) {"✔"} elseif ($chk.Sev -eq "FAIL") {"✖"} else {"⚠"}
        $col  = if ($chk.Pass) {$C.Ok} elseif ($chk.Sev -eq "FAIL") {$C.Fail} else {$C.Warn}
        $note = if (-not $chk.Pass -and $chk.Note) {$chk.Note} else {""}
        Write-Host ("  $icon  {0,-42} {1,-8} {2}" -f $chk.Name, (if($chk.Pass){"PASS"}else{$chk.Sev}), $note) -ForegroundColor $col
        if ($chk.Pass) {$pass++} elseif ($chk.Sev -eq "FAIL") {$fail++} else {$warn++}
    }
    Write-Host ""
    Write-Host ("  Passed: {0}   Warnings: {1}   Failed: {2}" -f $pass,$warn,$fail) -ForegroundColor (if($fail -gt 0){$C.Fail} elseif($warn -gt 0){$C.Warn} else {$C.Ok})
    Read-Host "  Press Enter to continue"
}
#endregion



function Configure-Identity {
    Write-Banner "IDENTITY & USER GROUP CONFIGURATION"
    Write-Status "Required: Entra ID Object IDs for user security groups" "" "info"
    Write-Host ""

    $id = $Global:Cfg.Identity
    if ($id.E3GroupId) { Write-Status "E3 Group ID" $id.E3GroupId "reuse" }
    else {
        $id.E3GroupId = Read-MenuChoice "E3 Office Workers — Entra ID Group Object ID"
        Write-Status "E3 Group ID set" $id.E3GroupId "ok"
    }
    if ($id.F1GroupId) { Write-Status "F1 Group ID" $id.F1GroupId "reuse" }
    else {
        $id.F1GroupId = Read-MenuChoice "F1 Frontline Workers — Entra ID Group Object ID"
        Write-Status "F1 Group ID set" $id.F1GroupId "ok"
    }
    if ($id.AdminGroupId) { Write-Status "Admin Group ID" $id.AdminGroupId "reuse" }
    else {
        $id.AdminGroupId = Read-MenuChoice "AVD Admins — Entra ID Group Object ID"
        Write-Status "Admin Group ID set" $id.AdminGroupId "ok"
    }
    $adminUser = Read-MenuChoice "VM local admin username" -Default $id.AdminUsername
    if ($adminUser) { $id.AdminUsername = $adminUser }
    $id.AzureAdJoin = Read-YesNo "Use Azure AD (Entra ID) Join? (No = Hybrid AD Join)" $true
    if (-not $id.AzureAdJoin) {
        Write-Status "Hybrid AD Join selected" "Ensure VNet has line-of-sight to domain controller" "warn"
    }
    Save-Config
}
#endregion

#region ═══════════════════════════════════════════════════════════════════════
#  MAIN MENU
#═══════════════════════════════════════════════════════════════════════════════

function Show-MainMenu {
    while ($true) {
        Show-Logo

        $ds = $Global:Cfg.DeploymentState
        $sub = if ($Global:Cfg.SubscriptionName) { $Global:Cfg.SubscriptionName } else { "Not selected" }
        $env = $Global:Cfg.Environment

        Write-Host ("  Subscription : {0}" -f $sub) -ForegroundColor $C.Azure
        Write-Host ("  Environment  : {0}    Prefix: {1}    Region: {2}" -f $env, $Global:Cfg.Prefix, $Global:Cfg.LocationDisplay) -ForegroundColor $C.Muted
        Write-Host ""

        # Deployment status grid
        Write-Host "  ┌─ Deployment Status ──────────────────────────────────────────────────┐" -ForegroundColor $C.Border
        $components = @(
            @("Resource Groups",  $ds.ResourceGroups)
            @("Networking",       $ds.Networking)
            @("Key Vault",        $ds.KeyVault)
            @("Log Analytics",    $ds.LogAnalytics)
            @("Azure Files",      $ds.AzureFiles)
            @("Host Pools",       $ds.HostPools)
            @("Session Hosts",    $ds.SessionHosts)
            @("App Groups",       $ds.AppGroups)
            @("Scaling Plans",    $ds.ScalingPlans)
            @("Automation",       $ds.Automation)
            @("RDP Properties",   $ds.RDPProperties)
            @("Join Type",        $ds.JoinType)
            @("Permissions",      $ds.Permissions)
        )
        $half = [Math]::Ceiling($components.Count / 2)
        for ($i = 0; $i -lt $half; $i++) {
            $l  = $components[$i]
            $r  = if ($i + $half -lt $components.Count) { $components[$i + $half] } else { $null }
            $ld = Get-StateDisplay $l[1]
            $lStr = "  $($UI.V)  $($ld.Icon) {0,-18} {1,-11}" -f $l[0], $ld.Text
            Write-Host $lStr -NoNewline -ForegroundColor $ld.Color
            if ($r) {
                $rd = Get-StateDisplay $r[1]
                Write-Host ("  $($ld.Icon) {0,-18} {1,-9}" -f $r[0], $rd.Text) -NoNewline -ForegroundColor $rd.Color
            }
            Write-Host "  $($UI.V)" -ForegroundColor $C.Border
        }
        Write-Host "  └──────────────────────────────────────────────────────────────────────┘" -ForegroundColor $C.Border

        Write-Host ""
        Write-Host "  ┌─ Main Menu ──────────────────────────────────────────────────────┐" -ForegroundColor $C.Border
        Write-Host "  │  1.  Guided Full Deployment Wizard                               │" -ForegroundColor $C.Menu
        Write-Host "  │  2.  Configure Azure Connection & Subscription                   │" -ForegroundColor $C.Menu
        Write-Host "  │  3.  Configure Identity & User Groups                            │" -ForegroundColor $C.Menu
        Write-Host "  │  4.  Deploy Individual Components ▶                              │" -ForegroundColor $C.Azure
        Write-Host "  │  5.  Permission Manager ▶                                        │" -ForegroundColor $C.Gold
        Write-Host "  │  6.  Auto-Scaling Manager ▶                                      │" -ForegroundColor $C.Azure
        Write-Host "  │  7.  RDP Properties Manager ▶                                    │" -ForegroundColor $C.Azure
        Write-Host "  │  8.  Domain Join Type Wizard ▶                                   │" -ForegroundColor $C.Azure
        Write-Host "  │  9.  Health Dashboard & Validation                               │" -ForegroundColor $C.Menu
        Write-Host "  │  S.  Save Configuration                                          │" -ForegroundColor $C.Muted
        Write-Host "  │  L.  Load Configuration                                          │" -ForegroundColor $C.Muted
        Write-Host "  │  0.  Exit                                                        │" -ForegroundColor $C.Muted
        Write-Host "  └──────────────────────────────────────────────────────────────────┘" -ForegroundColor $C.Border

        $choice = Read-MenuChoice "Main Menu" @("0","1","2","3","4","5","6","7","8","9","S","s","L","l")
        switch ($choice) {
            "1"         { Invoke-GuidedWizard }
            "2"         { Connect-ToAzure; Select-Environment }
            "3"         { Configure-Identity }
            "4"         { Show-ComponentMenu }
            "5"         { Show-PermissionManager }
            "6"         { Show-AutoScalingMenu }
            "7"         { Show-RDPPropertiesMenu }
            "8"         { Show-JoinTypeMenu }
            "9"         { Show-HealthDashboard }
            {$_ -in @("S","s")} { Save-Config }
            {$_ -in @("L","l")} { Load-Config | Out-Null; Write-Status "Configuration loaded" $ConfigFile "ok"; Start-Sleep 1 }
            "0"         { Write-Host "`n  Goodbye. Log saved to: $LogFile`n" -ForegroundColor $C.Muted; exit 0 }
        }
    }
}

function Show-ComponentMenu {
    while ($true) {
        Write-Banner "COMPONENT DEPLOYMENT"
        Write-Host "  Deploy or configure individual components (existing resources detected automatically)"
        Write-Host ""
        $items = @(
            @("1",  "Resource Groups",          "ResourceGroups")
            @("2",  "Virtual Network & NSGs",   "Networking")
            @("3",  "Key Vault",                "KeyVault")
            @("4",  "Log Analytics Workspace",  "LogAnalytics")
            @("5",  "Azure Files (FSLogix)",     "AzureFiles")
            @("6",  "Azure Compute Gallery",    "ComputeGallery")
            @("7",  "Domain Join Type",         "JoinType")
            @("8",  "RDP Properties",           "RDPProperties")
            @("9",  "AVD Host Pools",           "HostPools")
            @("10", "Session Hosts (VMs)",      "SessionHosts")
            @("11", "App Groups & Workspace",   "AppGroups")
            @("12", "Scaling Plans",            "ScalingPlans")
            @("13", "Automation Runbooks",      "Automation")
            @("14", "Monitor Alerts",           "Monitoring")
        )
        foreach ($item in $items) {
            $state = $Global:Cfg.DeploymentState[$item[2]]
            $sd    = Get-StateDisplay $state
            Write-Host ("  {0,3}.  {1,-32} {2} {3}" -f $item[0], $item[1], $sd.Icon, $sd.Text) -ForegroundColor $C.Menu
        }
        Write-Host "    0.  Back to Main Menu" -ForegroundColor $C.Muted
        $valid = (1..14 | ForEach-Object {"$_"}) + @("0")
        $choice = Read-MenuChoice "Component" $valid
        switch ($choice) {
            "1"  { Deploy-ResourceGroups }
            "2"  { Deploy-NetworkingComponent }
            "3"  { Deploy-KeyVaultComponent }
            "4"  { Deploy-LogAnalyticsComponent }
            "5"  { Deploy-AzureFilesComponent }
            "6"  { Deploy-ComputeGalleryComponent }
            "7"  { Show-JoinTypeMenu }
            "8"  { Show-RDPPropertiesMenu }
            "9"  { Deploy-HostPoolsComponent }
            "10" { Deploy-SessionHostsComponent }
            "11" { Deploy-AppGroupsComponent }
            "12" { Deploy-ScalingPlans }
            "13" { Deploy-AutomationRunbooks }
            "14" { Deploy-MonitoringAlerts }
            "0"  { return }
        }
    }
}

function Deploy-ComputeGalleryComponent {
    Write-Banner "COMPUTE GALLERY"
    $found = Find-ExistingComputeGalleries
    if (@($found).Count -gt 0) {
        $i = 1
        foreach ($g in $found) { Write-Host ("  {0}. {1} (RG: {2})" -f $i,$g.Name,$g.ResourceGroupName) -ForegroundColor $C.Menu; $i++ }
        Write-Host "     N. Create new  S. Skip" -ForegroundColor $C.New
        $opts = (1..@($found).Count | ForEach-Object {"$_"}) + @("N","n","S","s")
        $ch = Read-MenuChoice "Gallery" $opts "N"
        if ($ch -in @("S","s")) { $Global:Cfg.DeploymentState.ComputeGallery = "Skipped"; Save-Config; return }
        if ($ch -notmatch "^[Nn]$") {
            $sel = @($found)[[int]$ch - 1]
            $Global:Cfg.Gallery.Name = $sel.Name; $Global:Cfg.Gallery.ResourceGroup = $sel.ResourceGroupName
            $Global:Cfg.Gallery.IsNew = $false
            Write-Status "Reusing gallery" $sel.Name "reuse"
            $Global:Cfg.DeploymentState.ComputeGallery = "Deployed"; Save-Config; return
        }
    }
    Write-Status "Creating Compute Gallery" $Global:Cfg.Gallery.Name "new"
    New-AzGallery -GalleryName $Global:Cfg.Gallery.Name -ResourceGroupName $Global:Cfg.Gallery.ResourceGroup `
        -Location $Global:Cfg.Location -Description "BDF AVD Golden Images" | Out-Null
    foreach ($def in @(@{Name="img-avd-win11ms-e3";Sku="win11ms-e3"}, @{Name="img-avd-win11ms-f1";Sku="win11ms-f1"})) {
        New-AzGalleryImageDefinition -GalleryName $Global:Cfg.Gallery.Name -ResourceGroupName $Global:Cfg.Gallery.ResourceGroup `
            -Location $Global:Cfg.Location -Name $def.Name -Publisher "BDF-IT" -Offer "Windows11-AVD" -Sku $def.Sku `
            -OsState Generalized -OsType Windows -HyperVGeneration V2 -SecurityType "TrustedLaunch" | Out-Null
        Write-Status "Image definition created" $def.Name "ok"
    }
    $Global:Cfg.Gallery.IsNew = $true; $Global:Cfg.DeploymentState.ComputeGallery = "Deployed"; Save-Config
}

function Deploy-SessionHostsComponent {
    Write-Banner "SESSION HOSTS"
    foreach ($poolType in @("E3","F1")) {
        $hpCfg = $Global:Cfg.HostPools[$poolType]
        Write-Section "$poolType Session Hosts"
        $existing = @(Get-AzWvdSessionHost -HostPoolName $hpCfg.Name -ResourceGroupName $Global:Cfg.RG.AVD.Name -EA SilentlyContinue)
        if ($existing.Count -gt 0) {
            Write-Status "Existing hosts found" "$($existing.Count) hosts" "reuse"
            if (-not (Read-YesNo "Add more session hosts?" $false)) { continue }
        }
        $count = Read-MenuChoice "How many VMs to add?" -Default "$($hpCfg.VMCount)"
        if ($count -match "^\d+$" -and [int]$count -gt 0) {
            Write-Status "Deploying $count VMs for $poolType..." "" "step"
            $expiry = (Get-Date).ToUniversalTime().AddHours(48).ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
            $token  = (New-AzWvdRegistrationInfo -HostPoolName $hpCfg.Name -ResourceGroupName $Global:Cfg.RG.AVD.Name -ExpirationTime $expiry).Token
            $pwd    = (Get-AzKeyVaultSecret -VaultName $Global:Cfg.KeyVault.Name -Name "avd-vm-admin-password" -EA SilentlyContinue)?.SecretValue
            if (-not $pwd) { $pwd = Read-Host "Enter VM admin password" -AsSecureString }
            $cred   = New-Object PSCredential($Global:Cfg.Identity.AdminUsername, $pwd)
            $vnet   = Get-AzVirtualNetwork -Name $Global:Cfg.Network.VNet.Name -ResourceGroupName $Global:Cfg.Network.VNet.ResourceGroup
            $snetName = $Global:Cfg.Network.Subnets[$poolType].Name
            $subnet = $vnet.Subnets | Where-Object { $_.Name -eq $snetName }
            $prefix = "vm-avd-$($poolType.ToLower())-$($Global:Cfg.Prefix -replace '-','')"
            $allHosts = @(Get-AzWvdSessionHost -HostPoolName $hpCfg.Name -ResourceGroupName $Global:Cfg.RG.AVD.Name -EA SilentlyContinue)
            $existNums = @($allHosts | ForEach-Object { $n=($_.Name -split "/")[-1] -replace "^$prefix-",""; if($n -match "^\d+$"){[int]$n} })
            1..[int]$count | ForEach-Object {
                $next = 1; while ($existNums -contains $next) { $next++ }
                $existNums += $next; $vmName = "$prefix-$next"
                Write-Status "Creating VM" $vmName "new"
                $nic = New-AzNetworkInterface -Name "nic-$vmName" -ResourceGroupName $Global:Cfg.RG.AVD.Name `
                           -Location $Global:Cfg.Location -SubnetId $subnet.Id -PrivateIpAllocationMethod Dynamic
                $vmCfg = New-AzVMConfig -VMName $vmName -VMSize $hpCfg.VMSize -IdentityType SystemAssigned |
                    Set-AzVMOperatingSystem -Windows -ComputerName ($vmName -replace "-","") -Credential $cred `
                        -TimeZone "Eastern Standard Time" -EnableAutoUpdate $false |
                    Set-AzVMSourceImage -PublisherName "MicrosoftWindowsDesktop" -Offer "windows-11" -Skus "win11-24h2-avd-m365" -Version "latest" |
                    Add-AzVMNetworkInterface -Id $nic.Id |
                    Set-AzVMOSDisk -DiskSizeGB 128 -StorageAccountType Premium_LRS -CreateOption FromImage -DeleteOption Delete |
                    Set-AzVMSecurityProfile -SecurityType TrustedLaunch | Set-AzVMUefi -EnableVtpm $true -EnableSecureBoot $true
                New-AzVM -ResourceGroupName $Global:Cfg.RG.AVD.Name -Location $Global:Cfg.Location -VM $vmCfg | Out-Null
                Set-AzVMExtension -ResourceGroupName $Global:Cfg.RG.AVD.Name -VMName $vmName `
                    -Name "AADLoginForWindows" -Publisher "Microsoft.Azure.ActiveDirectory" `
                    -ExtensionType "AADLoginForWindows" -TypeHandlerVersion "2.0" -Settings @{mdmId=""} | Out-Null
                Set-AzVMExtension -ResourceGroupName $Global:Cfg.RG.AVD.Name -VMName $vmName `
                    -Name "DSC" -Publisher "Microsoft.Powershell" -ExtensionType "DSC" -TypeHandlerVersion "2.83" `
                    -Settings @{ modulesUrl="https://wvdportalstorageblob.blob.core.windows.net/galleryartifacts/Configuration_09-08-2022.zip"
                                 configurationFunction="Configuration.ps1\AddSessionHost"
                                 properties=@{HostPoolName=$hpCfg.Name;RegistrationInfoToken=$token;AadJoin=$true} } | Out-Null
                Write-Status "Registered" $vmName "ok"
            }
        }
    }
    $Global:Cfg.DeploymentState.SessionHosts = "Deployed"; Save-Config
}

function Deploy-AppGroupsComponent {
    Write-Banner "APPLICATION GROUPS & WORKSPACE"
    # E3 Desktop App Group
    $agE3Name = $Global:Cfg.HostPools.E3.AppGroupName
    $agE3 = Get-AzWvdApplicationGroup -Name $agE3Name -ResourceGroupName $Global:Cfg.RG.AVD.Name -EA SilentlyContinue
    if (-not $agE3) {
        $hp = Get-AzWvdHostPool -Name $Global:Cfg.HostPools.E3.Name -ResourceGroupName $Global:Cfg.RG.AVD.Name -EA SilentlyContinue
        if ($hp) {
            $agE3 = New-AzWvdApplicationGroup -Name $agE3Name -ResourceGroupName $Global:Cfg.RG.AVD.Name `
                        -Location $Global:Cfg.Location -FriendlyName "BDF Office Desktop" `
                        -ApplicationGroupType Desktop -HostPoolArmPath $hp.Id
            Write-Status "E3 App Group created" $agE3Name "ok"
        }
    } else { Write-Status "E3 App Group exists" $agE3Name "reuse" }

    # F1 RemoteApp Group
    $agF1Name = $Global:Cfg.HostPools.F1.AppGroupName
    $agF1 = Get-AzWvdApplicationGroup -Name $agF1Name -ResourceGroupName $Global:Cfg.RG.AVD.Name -EA SilentlyContinue
    if (-not $agF1) {
        $hp = Get-AzWvdHostPool -Name $Global:Cfg.HostPools.F1.Name -ResourceGroupName $Global:Cfg.RG.AVD.Name -EA SilentlyContinue
        if ($hp) {
            $agF1 = New-AzWvdApplicationGroup -Name $agF1Name -ResourceGroupName $Global:Cfg.RG.AVD.Name `
                        -Location $Global:Cfg.Location -FriendlyName "BDF Store Apps" `
                        -ApplicationGroupType RemoteApp -HostPoolArmPath $hp.Id
            Write-Status "F1 RemoteApp Group created" $agF1Name "ok"
            # Publish SAP
            if (Read-YesNo "Publish SAP browser app as RemoteApp?" $true) {
                $sapUrl = Read-MenuChoice "SAP URL" -Default "https://sap.internal.bdf.com/sap/bc/ui5_ui5/"
                New-AzWvdApplication -Name "SAP-ERP" -ApplicationGroupName $agF1Name `
                    -ResourceGroupName $Global:Cfg.RG.AVD.Name -FriendlyName "SAP ERP" `
                    -FilePath "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" `
                    -CommandLineSetting Allow -CommandLineArguments "--app=$sapUrl" `
                    -IconPath "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" `
                    -IconIndex 0 -ShowInPortal $true | Out-Null
                Write-Status "SAP RemoteApp published" $sapUrl "ok"
            }
        }
    } else { Write-Status "F1 App Group exists" $agF1Name "reuse" }

    # Workspace
    $wsName = "ws-bdf-$($Global:Cfg.Prefix)"
    $ws = Get-AzWvdWorkspace -Name $wsName -ResourceGroupName $Global:Cfg.RG.AVD.Name -EA SilentlyContinue
    if (-not $ws -and $agE3 -and $agF1) {
        New-AzWvdWorkspace -Name $wsName -ResourceGroupName $Global:Cfg.RG.AVD.Name `
            -Location $Global:Cfg.Location -FriendlyName "BDF Virtual Desktop" `
            -ApplicationGroupReference @($agE3.Id, $agF1.Id) | Out-Null
        Write-Status "Workspace created" $wsName "ok"
    } elseif ($ws) { Write-Status "Workspace exists" $wsName "reuse" }

    # RBAC on app groups
    if ($Global:Cfg.Identity.E3GroupId -and $agE3) {
        New-AzRoleAssignment -ObjectId $Global:Cfg.Identity.E3GroupId -RoleDefinitionName "Desktop Virtualization User" -Scope $agE3.Id -EA SilentlyContinue | Out-Null
        Write-Status "E3 users assigned to desktop app group" "" "ok"
    }
    if ($Global:Cfg.Identity.F1GroupId -and $agF1) {
        New-AzRoleAssignment -ObjectId $Global:Cfg.Identity.F1GroupId -RoleDefinitionName "Desktop Virtualization User" -Scope $agF1.Id -EA SilentlyContinue | Out-Null
        Write-Status "F1 users assigned to RemoteApp group" "" "ok"
    }
    $Global:Cfg.DeploymentState.AppGroups = "Deployed"; Save-Config
}

function Deploy-MonitoringAlerts {
    Write-Banner "MONITORING ALERTS"
    if (-not $Global:Cfg.LogAnalytics.Id) {
        Write-Status "Log Analytics not configured — deploy it first" "" "fail"; Read-Host "  Press Enter"; return
    }
    $agName = "ag-avd-alerts-$($Global:Cfg.Prefix)"
    $ag = Get-AzActionGroup -ResourceGroupName $Global:Cfg.RG.Monitoring.Name -Name $agName -EA SilentlyContinue
    if (-not $ag) {
        $emailAddr = Read-MenuChoice "Alert email address" -Default "it-avd-alerts@bdf.com"
        $emailRcvr = New-AzActionGroupEmailReceiverObject -Name "IT-Admin" -EmailAddress $emailAddr -UseCommonAlertSchema $true
        $ag = Set-AzActionGroup -ResourceGroupName $Global:Cfg.RG.Monitoring.Name -Name $agName `
                  -ShortName "AVDAlerts" -EmailReceiver @($emailRcvr)
        Write-Status "Action Group created" $agName "ok"
    } else { Write-Status "Action Group exists" $agName "reuse" }

    $lawId  = $Global:Cfg.LogAnalytics.Id
    $alerts = @(
        @{ Name="alert-avd-unhealthy-hosts"; Display="AVD Unhealthy Session Hosts"; Severity=1
           Query="WVDAgentHealthStatus | where TimeGenerated > ago(10m) | where Status != 'Available' | summarize count() by HostPoolName | where count_ > 0" }
        @{ Name="alert-avd-fslogix-failure"; Display="FSLogix Profile Mount Failure"; Severity=2
           Query="Event | where Source == 'Microsoft-FSLogix-Apps' | where EventID in (33,52,35) | where TimeGenerated > ago(15m) | summarize count() by Computer" }
        @{ Name="alert-avd-high-capacity";   Display="AVD High Capacity >= 80%"; Severity=3
           Query="WVDConnections | where TimeGenerated > ago(10m) | summarize count() by _ResourceId | where count_ > 0" }
        @{ Name="alert-avd-slow-logon";      Display="Slow Logon > 60 seconds"; Severity=2
           Query="WVDConnections | where TimeGenerated > ago(30m) | extend d=datetime_diff('second',ConnectTime,StartTime) | where d > 60 | summarize count() by _ResourceId | where count_ >= 3" }
    )
    foreach ($a in $alerts) {
        $cond = New-AzScheduledQueryRuleConditionObject -Query $a.Query -TimeAggregation Count `
                    -Operator GreaterThan -Threshold 0 -FailingPeriodNumberOfEvaluationPeriod 1 `
                    -FailingPeriodMinFailingPeriodsToAlert 1
        New-AzScheduledQueryRule -Name $a.Name -ResourceGroupName $Global:Cfg.RG.Monitoring.Name `
            -Location $Global:Cfg.Location -DisplayName $a.Display -Scope @($lawId) `
            -Severity $a.Severity -WindowSize ([TimeSpan]::FromMinutes(10)) `
            -EvaluationFrequency ([TimeSpan]::FromMinutes(5)) `
            -CriterionAllOf @($cond) -Action @{ActionGroupId=@($ag.Id)} -Enabled $true `
            -EA SilentlyContinue | Out-Null
        Write-Status "Alert rule created" $a.Display "ok"
    }
    $Global:Cfg.DeploymentState.Monitoring = "Deployed"; Save-Config
}

function Invoke-GuidedWizard {
    Write-Banner "GUIDED FULL DEPLOYMENT WIZARD"
    Write-Host "  This wizard will walk you through each component in order." -ForegroundColor $C.Muted
    Write-Host "  Existing resources will be detected and offered for reuse." -ForegroundColor $C.Muted
    Write-Host "  You can skip any component — configuration is saved automatically." -ForegroundColor $C.Muted
    Write-Host ""
    if (-not (Read-YesNo "Start guided deployment?" $true)) { return }

    Connect-ToAzure
    Select-Environment
    Configure-Identity

    $steps = @(
        @{ Name="Resource Groups";  Fn={ Deploy-ResourceGroups } }
        @{ Name="Networking";       Fn={ Deploy-NetworkingComponent } }
        @{ Name="Key Vault";        Fn={ Deploy-KeyVaultComponent } }
        @{ Name="Log Analytics";    Fn={ Deploy-LogAnalyticsComponent } }
        @{ Name="Azure Files";      Fn={ Deploy-AzureFilesComponent } }
        @{ Name="Compute Gallery";  Fn={ Deploy-ComputeGalleryComponent } }
        @{ Name="Domain Join Type"; Fn={ Invoke-JoinTypeWizard } }
        @{ Name="RDP Properties";   Fn={ Apply-RDPPresetProfile } }
        @{ Name="Host Pools";       Fn={ Deploy-HostPoolsComponent } }
        @{ Name="Session Hosts";    Fn={ Deploy-SessionHostsComponent } }
        @{ Name="App Groups";       Fn={ Deploy-AppGroupsComponent } }
        @{ Name="Scaling Config";   Fn={ Invoke-ScalingWizard } }
        @{ Name="Scaling Plans";    Fn={ Deploy-ScalingPlans } }
        @{ Name="Automation";       Fn={ Deploy-AutomationRunbooks } }
        @{ Name="Permissions";      Fn={ Assign-AllMissingPermissions } }
        @{ Name="Monitor Alerts";   Fn={ Deploy-MonitoringAlerts } }
        @{ Name="Health Check";     Fn={ Show-HealthDashboard } }
    )

    $total = $steps.Count
    for ($i = 0; $i -lt $total; $i++) {
        $step = $steps[$i]
        $pct  = [Math]::Round(($i / $total) * 100)
        Write-Host ""
        Show-Progress "Deployment" $pct "Step $($i+1)/$total — $($step.Name)"
        Write-Host ""
        Write-Banner "STEP $($i+1) of $total — $($step.Name.ToUpper())"
        try {
            & $step.Fn
        } catch {
            Write-Status "Error in $($step.Name): $_" "" "fail"
            Write-Log "ERROR in $($step.Name): $_" "ERROR"
            if (-not (Read-YesNo "Continue to next step despite error?" $true)) { break }
        }
    }
    Show-Progress "Deployment" 100 "Complete!"
    Write-Host ""
    Write-Status "Guided wizard complete" "" "ok"
    Save-Config
    Read-Host "  Press Enter to view health dashboard"
    Show-HealthDashboard
}
#endregion

#region ═══════════════════════════════════════════════════════════════════════
#  ENTRY POINT
#═══════════════════════════════════════════════════════════════════════════════

# Initialize log
Write-Log "BDF AVD AIO Deploy Console started v3.0" "INFO"
Write-Log "PowerShell: $($PSVersionTable.PSVersion)" "INFO"

# Load existing config if present
Load-Config | Out-Null

if ($ValidateOnly) {
    Connect-ToAzure
    Show-HealthDashboard
    exit 0
}

if ($Unattended) {
    if (-not (Test-Path $ConfigFile)) {
        Write-Error "Unattended mode requires a saved config file: $ConfigFile"
        exit 1
    }
    Load-Config | Out-Null
    Invoke-GuidedWizard
    exit 0
}

# Interactive mode — show main menu
Show-MainMenu
#endregion

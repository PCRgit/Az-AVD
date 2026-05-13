#Requires -Version 7.0
<#
.SYNOPSIS
    BDF Azure Virtual Desktop — FSLogix, App Masking & App Attach AIO Module
    Standalone or dot-sourced extension for the BDF-AVD-Deploy-AIO.ps1 console.

.DESCRIPTION
    Comprehensive management of:
      ► FSLogix Profile Containers  — Wizard, Intune/GPO export, registry deploy
      ► FSLogix Office Containers   — ODFC config, Teams/Outlook/OneDrive cache
      ► Cloud Cache                 — Multi-site HA profile replication
      ► Redirections.xml Builder    — Interactive exclusion/redirect catalog
      ► FSLogix App Masking         — Rule file builder, group assignments, presets
      ► MSIX App Attach             — Package management, staging, health monitoring
      ► AVD App Attach (Preview)    — Next-gen app attach package lifecycle
      ► Diagnostics & Reporting     — Profile health, size, mount failures, KQL
      ► Remediation Toolkit         — Fix common FSLogix/App Attach issues

.NOTES
    Author  : Jaimin
    Version : 1.0
    Requires: Az PowerShell, Az.DesktopVirtualization 4.0+, PowerShell 7+
    Usage   : .\BDF-AVD-FSLogix-AppAttach.ps1
              OR: . .\BDF-AVD-FSLogix-AppAttach.ps1  (dot-source into AIO script)
#>

[CmdletBinding()]
param(
    [string]$ConfigFile = ".\AVD-Config.json",
    [string]$LogFile    = ".\AVD-FSLogix-$(Get-Date -Format 'yyyyMMdd-HHmmss').log",
    [switch]$NoLogo
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference    = "SilentlyContinue"

#region ════════════════════════════════════════════════════════════════════
# UI HELPERS (self-contained — works standalone or with AIO)
#════════════════════════════════════════════════════════════════════════════
$FX = @{
    TL='╔';TR='╗';BL='╚';BR='╝';H='═';V='║'
    SH='─';SV='│';STL='┌';STR='┐';SBL='└';SBR='┘'
    Arrow='▶';Check='✔';Cross='✖';Warn='⚠';Info='ℹ'
    New='⊕';Reuse='↺';Skip='⊘';Step='▸';Deploy='⚡'
    File='📄';Folder='📁';Lock='🔒';App='📦';Gear='⚙'
}
$FC = @{
    Hdr='Cyan';Border='DarkCyan';Menu='White';Accent='Magenta'
    Ok='Green';Fail='Red';Warn='Yellow';Info='Cyan';Muted='DarkGray'
    New='Green';Reuse='Cyan';Skip='DarkGray';Step='DarkCyan'
    Azure='Blue';Gold='DarkYellow';Purple='Magenta'
}

function fx-log  { param([string]$m,[string]$l="INFO"); Add-Content $LogFile "[$( Get-Date -f 'yyyy-MM-dd HH:mm:ss')][$l] $m" -EA SilentlyContinue }
function fx-hdr  {
    param([string]$T,[string]$Color=$FC.Hdr,[int]$W=72)
    $l=$FX.H*$W; $pad=[Math]::Max(0,($W-$T.Length-4)); $lp=[Math]::Floor($pad/2); $rp=$pad-$lp
    Write-Host ""; Write-Host "$($FX.TL)$l$($FX.TR)" -ForegroundColor $FC.Border
    Write-Host "$($FX.V)  $(' '*$lp)$T$(' '*$rp)  $($FX.V)" -ForegroundColor $Color
    Write-Host "$($FX.BL)$l$($FX.BR)" -ForegroundColor $FC.Border; Write-Host ""
}
function fx-sec  { param([string]$T,[string]$C=$FC.Accent); Write-Host ""; Write-Host "  $($FX.STL)$($FX.SH*68)$($FX.STR)" -ForegroundColor $FC.Border; Write-Host "  $($FX.SV)  $($FX.Arrow) $T" -ForegroundColor $C; Write-Host "  $($FX.SBL)$($FX.SH*68)$($FX.SBR)" -ForegroundColor $FC.Border }
function fx-ok   { param([string]$m,[string]$v=""); Write-Host "  $($FX.Check)  $([string]::Format('{0,-36}',$m))$v" -ForegroundColor $FC.Ok;   fx-log "$m $v" "OK"   }
function fx-fail { param([string]$m,[string]$v=""); Write-Host "  $($FX.Cross)  $([string]::Format('{0,-36}',$m))$v" -ForegroundColor $FC.Fail; fx-log "$m $v" "FAIL" }
function fx-warn { param([string]$m,[string]$v=""); Write-Host "  $($FX.Warn)   $([string]::Format('{0,-36}',$m))$v" -ForegroundColor $FC.Warn; fx-log "$m $v" "WARN" }
function fx-info { param([string]$m,[string]$v=""); Write-Host "  $($FX.Info)   $([string]::Format('{0,-36}',$m))$v" -ForegroundColor $FC.Info; fx-log "$m $v" "INFO" }
function fx-step { param([string]$m);               Write-Host "  $($FX.Step)   $m" -ForegroundColor $FC.Step; fx-log $m "STEP" }
function fx-new  { param([string]$m,[string]$v=""); Write-Host "  $($FX.New)    $([string]::Format('{0,-36}',$m))$v" -ForegroundColor $FC.New  }
function fx-line { Write-Host "  $($FX.SH*68)" -ForegroundColor $FC.Border }

function Read-Inp {
    param([string]$P,[string]$D="",[string[]]$V=@())
    do {
        Write-Host "  $($FX.Arrow) " -NoNewline -ForegroundColor $FC.Accent
        Write-Host $P -NoNewline -ForegroundColor $FC.Menu
        if ($D) { Write-Host " [$D]" -NoNewline -ForegroundColor $FC.Muted }
        Write-Host " : " -NoNewline -ForegroundColor $FC.Muted
        $r = Read-Host
        if (-not $r -and $D) { $r = $D }
    } until (-not $V -or $r -in $V)
    return $r.Trim()
}
function Read-YN {
    param([string]$P,[bool]$D=$true)
    $opts = if ($D) {"[Y/n]"} else {"[y/N]"}
    $def  = if ($D) {"Y"} else {"N"}
    $r = Read-Inp "$P $opts" -D $def -V @("Y","y","N","n","")
    return ($r -in @("Y","y"))
}
function Pause-Screen { Write-Host ""; Read-Host "  Press Enter to continue" }

# Load config if present (shared with AIO script)
$Script:Cfg = @{ Prefix="bdf-poc"; Location="eastus"; Environment="POC"
    RG=@{AVD=@{Name="rg-avd-bdf-poc"}; Storage=@{Name="rg-avd-storage-bdf-poc"}; Monitoring=@{Name="rg-avd-monitoring-bdf-poc"}}
    Storage=@{Account=@{Name="stavdbdfpoc"; ResourceGroup="rg-avd-storage-bdf-poc"}; Shares=@{E3="profiles-e3"; F1="profiles-f1"; ODFC="odfc-e3"}}
    HostPools=@{E3=@{Name="hp-avd-e3-office-bdf-poc"}; F1=@{Name="hp-avd-f1-frontline-bdf-poc"}}
    KeyVault=@{Name="kv-avd-bdf-poc"}
    LogAnalytics=@{Name="law-avd-bdf-poc"; ResourceGroup="rg-avd-monitoring-bdf-poc"; Id=""}
    Gallery=@{Name="acg_avd_bdf_poc"; ResourceGroup="rg-avd-bdf-poc"}
    SubscriptionId=""; TenantId=""
    FSLogix=@{
        Deployed=$false
        E3VHDLocation=""; F1VHDLocation=""; ODFCLocation=""
        E3SizeGB=10; F1SizeGB=2; ODFCSizeGB=25
        CloudCache=@{Enabled=$false; Locations=@()}
        ExclusionsDeployed=$false
    }
    AppMasking=@{ Deployed=$false; RuleFiles=@() }
    AppAttach=@{ Enabled=$false; ShareName="appattach"; Packages=@() }
    OneDrive=@{
        Deployed            = $false
        TenantId            = ""
        SilentSignIn        = $true
        KFMEnabled          = $true
        KFMFolders          = @("Desktop","Documents","Pictures")
        KFMBlockOptOut      = $true
        KFMSilentOptIn      = $true
        KFMNotification     = $false
        FilesOnDemand       = $true
        BlockPersonalSync   = $true
        PerMachineInstall   = $true
        DisableFirstRun     = $true
        FSLogixExclAdded    = $false
        InstallVersion      = ""
    }
}
if (Test-Path $ConfigFile) {
    try {
        $loaded = Get-Content $ConfigFile -Raw | ConvertFrom-Json -AsHashtable
        foreach ($k in $loaded.Keys) { if ($Script:Cfg.ContainsKey($k)) { $Script:Cfg[$k] = $loaded[$k] } }
        fx-info "Loaded config" $ConfigFile
    } catch { fx-warn "Could not load config" $ConfigFile }
}
#endregion

#region ════════════════════════════════════════════════════════════════════
# FSLOGIX CONSTANTS & REGISTRY PATHS
#════════════════════════════════════════════════════════════════════════════
$FXReg = @{
    ProfileRoot  = 'HKLM:\SOFTWARE\FSLogix\Profiles'
    ODFCRoot     = 'HKLM:\SOFTWARE\Policies\FSLogix\ODFC'
    LogRoot      = 'HKLM:\SOFTWARE\FSLogix\Logging'
    AppMaskRoot  = 'HKLM:\SOFTWARE\FSLogix\AppMasking'
    CCDRoot      = 'HKLM:\SOFTWARE\FSLogix\Profiles'
}

# Categorized best-practice settings
$FXSettings = [ordered]@{
    "PROFILE CONTAINER — CORE" = @(
        @{Key="Enabled";               Val=1;    Type="DWord";  Desc="Enable FSLogix profile containers"}
        @{Key="SizeInMBs";             Val=10240; Type="DWord"; Desc="Profile VHDX max size (E3: 10GB, F1: 2GB)"}
        @{Key="VolumeType";            Val="VHDX"; Type="String"; Desc="Use VHDX format (required for modern features)"}
        @{Key="IsDynamic";             Val=1;    Type="DWord";  Desc="Dynamic VHDX — only uses actual space"}
        @{Key="FlipFlopProfileDirectoryName"; Val=1; Type="DWord"; Desc="USERNAME_SID naming (readable folder names)"}
        @{Key="DeleteLocalProfileWhenVHDShouldApply"; Val=1; Type="DWord"; Desc="Remove stale local profile when VHD available"}
        @{Key="ConcurrentUserSessions"; Val=1;  Type="DWord";  Desc="Allow same user on multiple pooled VMs"}
        @{Key="ProfileType";           Val=0;   Type="DWord";  Desc="0=Normal, 3=Read-only (use 0 for pooled)"}
        @{Key="LockedRetryCount";      Val=3;   Type="DWord";  Desc="Retry count if VHDX locked (concurrent access)"}
        @{Key="LockedRetryInterval";   Val=15;  Type="DWord";  Desc="Seconds between lock retries"}
        @{Key="ReAttachIntervalSeconds";Val=15; Type="DWord";  Desc="Re-attach interval on network interruption"}
        @{Key="ReAttachRetryCount";    Val=3;   Type="DWord";  Desc="Re-attach retry attempts"}
    )
    "PROFILE CONTAINER — ACCESS & SECURITY" = @(
        @{Key="AccessNetworkAsComputerObject"; Val=1; Type="DWord"; Desc="Use computer object for SMB auth (Azure AD Kerberos)"}
        @{Key="RequireNetworkConnectivity"; Val=0; Type="DWord"; Desc="0=Fall back to local if VHD unavailable; 1=Block logon"}
        @{Key="PreventLoginWithFailure"; Val=0; Type="DWord"; Desc="0=Allow local profile fallback; 1=Block login on VHD fail"}
        @{Key="PreventLoginWithTempProfile"; Val=0; Type="DWord"; Desc="0=Allow temp profile; 1=Block (safer but disruptive)"}
        @{Key="RoamSearch";            Val=1;   Type="DWord";  Desc="Enable Windows Search index roaming"}
    )
    "OFFICE CONTAINER (ODFC)" = @(
        @{Key="Enabled";               Val=1;   Type="DWord";  Desc="Enable Office 365 Container"}
        @{Key="VolumeType";            Val="VHDX"; Type="String"; Desc="VHDX format for ODFC"}
        @{Key="IsDynamic";             Val=1;   Type="DWord";  Desc="Dynamic VHDX for ODFC"}
        @{Key="SizeInMBs";             Val=25600; Type="DWord"; Desc="ODFC size: 25GB (Teams+Outlook+OneDrive cache)"}
        @{Key="IncludeTeams";          Val=1;   Type="DWord";  Desc="Route Teams cache to ODFC (CRITICAL — keeps main profile small)"}
        @{Key="IncludeOneDrive";       Val=1;   Type="DWord";  Desc="Route OneDrive cache to ODFC"}
        @{Key="IncludeOutlook";        Val=1;   Type="DWord";  Desc="Route Outlook OST/NST files to ODFC"}
        @{Key="IncludeOutlookPersonalization"; Val=1; Type="DWord"; Desc="Include Outlook signatures and templates"}
        @{Key="IncludeSharepoint";     Val=0;   Type="DWord";  Desc="SharePoint offline files (disable if using OneDrive KFM)"}
        @{Key="FlipFlopProfileDirectoryName"; Val=1; Type="DWord"; Desc="Readable folder naming for ODFC shares"}
    )
    "LOGGING" = @(
        @{Key="LogDir";    Val='C:\ProgramData\FSLogix\Logs'; Type="String"; Desc="Log directory path"}
        @{Key="LogLevel";  Val=2;   Type="DWord"; Desc="0=None, 1=Error, 2=Info+Error, 3=Debug (use 2 production)"}
        @{Key="LogMaxSize"; Val=100; Type="DWord"; Desc="Max log file size in MB before rotation"}
    )
}

# Common exclusion catalog
$FXExclusions = [ordered]@{
    "BROWSER CACHE" = @(
        @{Path='%LOCALAPPDATA%\Microsoft\Edge\User Data\Default\Cache'; Type="Directory"; Desc="Edge browser cache (can reach 1GB+)"}
        @{Path='%LOCALAPPDATA%\Microsoft\Edge\User Data\Default\Code Cache'; Type="Directory"; Desc="Edge code cache"}
        @{Path='%LOCALAPPDATA%\Google\Chrome\User Data\Default\Cache'; Type="Directory"; Desc="Chrome cache"}
        @{Path='%LOCALAPPDATA%\Mozilla\Firefox\Profiles\*\cache2'; Type="Directory"; Desc="Firefox cache"}
    )
    "WINDOWS TEMP & SYSTEM" = @(
        @{Path='%LOCALAPPDATA%\Temp';                              Type="Directory"; Desc="User temp files"}
        @{Path='%LOCALAPPDATA%\Microsoft\Windows\INetCache';       Type="Directory"; Desc="IE/WebView cache"}
        @{Path='%LOCALAPPDATA%\Microsoft\Windows\WebCache';        Type="Directory"; Desc="Windows WebCache database"}
        @{Path='%LOCALAPPDATA%\CrashDumps';                        Type="Directory"; Desc="Application crash dumps"}
        @{Path='%LOCALAPPDATA%\Microsoft\Windows\Explorer\thumbcache*'; Type="File"; Desc="Thumbnail cache files"}
    )
    "WINDOWS SEARCH" = @(
        @{Path='%LOCALAPPDATA%\Microsoft\Windows\Caches';         Type="Directory"; Desc="Windows Search/Cortana caches"}
        @{Path='%APPDATA%\Microsoft\Search\Data\Applications\Windows'; Type="Directory"; Desc="Search app data"}
    )
    "MICROSOFT OFFICE TEMP" = @(
        @{Path='%APPDATA%\Microsoft\Teams\Service Worker\CacheStorage'; Type="Directory"; Desc="Teams service worker cache"}
        @{Path='%APPDATA%\Microsoft\Teams\Cache';                  Type="Directory"; Desc="Teams main cache (in ODFC — exclude from profile)"}
        @{Path='%APPDATA%\Microsoft\Teams\blob_storage';           Type="Directory"; Desc="Teams blob storage"}
        @{Path='%APPDATA%\Microsoft\Teams\databases';              Type="Directory"; Desc="Teams IndexedDB"}
        @{Path='%APPDATA%\Microsoft\Teams\GPUCache';               Type="Directory"; Desc="Teams GPU cache"}
        @{Path='%LOCALAPPDATA%\Microsoft\Office\16.0\Wef';         Type="Directory"; Desc="Office web extension cache"}
    )
    "APPLICATION CACHES" = @(
        @{Path='%LOCALAPPDATA%\Adobe\*\Cache';                     Type="Directory"; Desc="Adobe product caches"}
        @{Path='%LOCALAPPDATA%\Packages\*\TempState';              Type="Directory"; Desc="UWP app temp state"}
        @{Path='%LOCALAPPDATA%\Packages\*\AC\Temp';               Type="Directory"; Desc="UWP app temp files"}
        @{Path='%LOCALAPPDATA%\Microsoft\CLR_v*';                  Type="Directory"; Desc=".NET runtime caches"}
        @{Path='%LOCALAPPDATA%\assembly\dl3';                      Type="Directory"; Desc=".NET assembly cache"}
    )
    "ONEDRIVE (if using KFM)" = @(
        @{Path='%USERPROFILE%\OneDrive*';                          Type="Directory"; Desc="OneDrive sync folder (use KFM instead of profile)"}
        @{Path='%LOCALAPPDATA%\Microsoft\OneDrive\logs';           Type="Directory"; Desc="OneDrive log files"}
        @{Path='%LOCALAPPDATA%\Microsoft\OneDrive\setup';          Type="Directory"; Desc="OneDrive setup cache"}
    )
}
#endregion

#region ════════════════════════════════════════════════════════════════════
# FSLOGIX PROFILE CONTAINER WIZARD
#════════════════════════════════════════════════════════════════════════════

function Invoke-FSLogixProfileWizard {
    fx-hdr "FSLOGIX PROFILE CONTAINER WIZARD"

    fx-sec "Storage Configuration"
    # Detect existing Azure Files
    fx-step "Scanning for Azure Files Premium storage accounts..."
    $storageAccounts = Get-AzStorageAccount | Where-Object { $_.Kind -eq "FileStorage" } |
        Select-Object StorageAccountName, ResourceGroupName,
            @{N="Sku";E={$_.Sku.Name}} | Sort-Object StorageAccountName

    if (@($storageAccounts).Count -gt 0) {
        $i=1; foreach ($s in $storageAccounts) {
            Write-Host ("  {0}. {1,-35} RG:{2,-25} SKU:{3}" -f $i,$s.StorageAccountName,$s.ResourceGroupName,$s.Sku) -ForegroundColor $FC.Menu; $i++
        }
        Write-Host "     N. Enter custom UNC path manually" -ForegroundColor $FC.New
        $opts = (1..@($storageAccounts).Count | ForEach-Object {"$_"}) + @("N","n")
        $ch = Read-Inp "Select storage account or N for manual" -V $opts -D "1"
        if ($ch -notmatch "^[Nn]$") {
            $sel = @($storageAccounts)[[int]$ch-1]
            $fqdn = "$($sel.StorageAccountName).file.core.windows.net"
            $Script:Cfg.FSLogix.E3VHDLocation  = "\\$fqdn\$($Script:Cfg.Storage.Shares.E3)"
            $Script:Cfg.FSLogix.F1VHDLocation  = "\\$fqdn\$($Script:Cfg.Storage.Shares.F1)"
            $Script:Cfg.FSLogix.ODFCLocation   = "\\$fqdn\$($Script:Cfg.Storage.Shares.ODFC)"
            fx-ok "VHD Locations set" $fqdn
        }
    }

    # Allow manual override
    if (-not $Script:Cfg.FSLogix.E3VHDLocation) {
        $Script:Cfg.FSLogix.E3VHDLocation  = Read-Inp "E3 Profile VHD UNC path" -D "\\storageaccount.file.core.windows.net\profiles-e3"
        $Script:Cfg.FSLogix.F1VHDLocation  = Read-Inp "F1 Profile VHD UNC path" -D "\\storageaccount.file.core.windows.net\profiles-f1"
        $Script:Cfg.FSLogix.ODFCLocation   = Read-Inp "ODFC VHD UNC path"        -D "\\storageaccount.file.core.windows.net\odfc-e3"
    }

    # Sizes
    fx-sec "Profile VHDX Size Configuration"
    $e3size  = Read-Inp "E3 Office Worker profile size (GB)"    -D "10"
    $f1size  = Read-Inp "F1 Frontline profile size (GB)"        -D "2"
    $odfc    = Read-Inp "ODFC (Office Container) size (GB)"     -D "25"
    $Script:Cfg.FSLogix.E3SizeGB   = [int]$e3size
    $Script:Cfg.FSLogix.F1SizeGB   = [int]$f1size
    $Script:Cfg.FSLogix.ODFCSizeGB = [int]$odfc

    # Cloud Cache
    fx-sec "Cloud Cache (High Availability)"
    fx-info "Cloud Cache writes profiles to multiple Azure Files locations simultaneously"
    fx-info "Provides near-zero RPO for profile data — recommended for Production"
    $Script:Cfg.FSLogix.CloudCache.Enabled = Read-YN "Enable Cloud Cache?" ($Script:Cfg.Environment -eq "Production")
    if ($Script:Cfg.FSLogix.CloudCache.Enabled) {
        fx-info "Add multiple Azure Files UNC paths (secondary regions for HA)"
        $locations = @()
        do {
            $loc = Read-Inp "Cloud Cache location UNC (blank to finish)" -D ""
            if ($loc) { $locations += $loc; fx-ok "Added" $loc }
        } until (-not $loc -or $locations.Count -ge 4)
        $Script:Cfg.FSLogix.CloudCache.Locations = $locations
        fx-ok "Cloud Cache locations" "$($locations.Count) configured"
    }

    # Deployment method
    fx-sec "Deployment Method"
    Write-Host "  1.  Intune Settings Catalog (JSON — recommended for Entra ID-joined AVD)" -ForegroundColor $FC.Menu
    Write-Host "  2.  Group Policy (ADMX — for Hybrid AD-joined)" -ForegroundColor $FC.Menu
    Write-Host "  3.  Registry Script (PowerShell — direct session host deployment)" -ForegroundColor $FC.Menu
    Write-Host "  4.  All three formats" -ForegroundColor $FC.New
    $method = Read-Inp "Deployment method" -V @("1","2","3","4") -D "4"

    if ($method -in @("1","4")) { Export-FSLogixIntunePolicy }
    if ($method -in @("2","4")) { Export-FSLogixGPOSettings }
    if ($method -in @("3","4")) { Export-FSLogixRegistryScript }

    $Script:Cfg.FSLogix.Deployed = $true
    Save-FXConfig
    Pause-Screen
}

function Export-FSLogixIntunePolicy {
    fx-sec "Generating Intune Settings Catalog JSON"
    $cfg = $Script:Cfg.FSLogix
    $e3Mb = $cfg.E3SizeGB * 1024
    $f1Mb = $cfg.F1SizeGB * 1024
    $ofMb = $cfg.ODFCSizeGB * 1024

    $buildSettingsBlock = {
        param([string]$vhdLoc, [int]$sizeMb, [string]$label)
        @{
            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSetting"
            "settingInstance" = @{
                "@odata.type" = "#microsoft.graph.deviceManagementConfigurationGroupSettingCollectionInstance"
                "settingDefinitionId" = "device_vendor_msft_policy_config_fslogix~policy~fslogix~profiles"
                "groupSettingCollectionValue" = @(
                    @{ "children" = @(
                        @{ "@odata.type"="#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"; "settingDefinitionId"="fslogix_profiles_enabled"; "choiceSettingValue"=@{"value"="fslogix_profiles_enabled_1"} }
                        @{ "@odata.type"="#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"; "settingDefinitionId"="fslogix_profiles_vhdlocations"; "simpleSettingValue"=@{"@odata.type"="#microsoft.graph.deviceManagementConfigurationStringSettingValue";"value"=$vhdLoc} }
                        @{ "@odata.type"="#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"; "settingDefinitionId"="fslogix_profiles_sizeinmbs"; "simpleSettingValue"=@{"@odata.type"="#microsoft.graph.deviceManagementConfigurationIntegerSettingValue";"value"=$sizeMb} }
                        @{ "@odata.type"="#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"; "settingDefinitionId"="fslogix_profiles_volumetype"; "choiceSettingValue"=@{"value"="fslogix_profiles_volumetype_vhdx"} }
                        @{ "@odata.type"="#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"; "settingDefinitionId"="fslogix_profiles_isdynamic"; "choiceSettingValue"=@{"value"="fslogix_profiles_isdynamic_1"} }
                        @{ "@odata.type"="#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"; "settingDefinitionId"="fslogix_profiles_flipflopprofiledirectoryname"; "choiceSettingValue"=@{"value"="fslogix_profiles_flipflopprofiledirectoryname_1"} }
                        @{ "@odata.type"="#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"; "settingDefinitionId"="fslogix_profiles_deletelocalprofilewhenvhdshouldapply"; "choiceSettingValue"=@{"value"="fslogix_profiles_deletelocalprofilewhenvhdshouldapply_1"} }
                        @{ "@odata.type"="#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"; "settingDefinitionId"="fslogix_profiles_concurrentusersessions"; "choiceSettingValue"=@{"value"="fslogix_profiles_concurrentusersessions_1"} }
                        @{ "@odata.type"="#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"; "settingDefinitionId"="fslogix_profiles_accessnetworkascomputerobject"; "choiceSettingValue"=@{"value"="fslogix_profiles_accessnetworkascomputerobject_1"} }
                    ) }
                )
            }
        }
    }

    $policies = @(
        @{
            Name        = "BDF-AVD-FSLogix-E3-Office"
            Description = "FSLogix Profile Container settings for M365 E3 Office Workers"
            Settings    = @(& $buildSettingsBlock $cfg.E3VHDLocation ($cfg.E3SizeGB * 1024) "E3")
        },
        @{
            Name        = "BDF-AVD-FSLogix-F1-Frontline"
            Description = "FSLogix Profile Container settings for M365 F1 Frontline Workers"
            Settings    = @(& $buildSettingsBlock $cfg.F1VHDLocation ($cfg.F1SizeGB * 1024) "F1")
        },
        @{
            Name        = "BDF-AVD-FSLogix-ODFC-Office"
            Description = "FSLogix Office Container (ODFC) for E3 — Teams/Outlook/OneDrive cache"
            Settings    = @(
                @{
                    "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSetting"
                    "settingInstance" = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationGroupSettingCollectionInstance"
                        "settingDefinitionId" = "device_vendor_msft_policy_config_fslogix~policy~fslogix~odfc"
                        "groupSettingCollectionValue" = @(
                            @{ "children" = @(
                                @{ "@odata.type"="#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"; "settingDefinitionId"="fslogix_odfc_enabled"; "choiceSettingValue"=@{"value"="fslogix_odfc_enabled_1"} }
                                @{ "@odata.type"="#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"; "settingDefinitionId"="fslogix_odfc_vhdlocations"; "simpleSettingValue"=@{"@odata.type"="#microsoft.graph.deviceManagementConfigurationStringSettingValue";"value"=$cfg.ODFCLocation} }
                                @{ "@odata.type"="#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"; "settingDefinitionId"="fslogix_odfc_sizeinmbs"; "simpleSettingValue"=@{"@odata.type"="#microsoft.graph.deviceManagementConfigurationIntegerSettingValue";"value"=$ofMb} }
                                @{ "@odata.type"="#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"; "settingDefinitionId"="fslogix_odfc_includeteams"; "choiceSettingValue"=@{"value"="fslogix_odfc_includeteams_1"} }
                                @{ "@odata.type"="#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"; "settingDefinitionId"="fslogix_odfc_includeonedrive"; "choiceSettingValue"=@{"value"="fslogix_odfc_includeonedrive_1"} }
                                @{ "@odata.type"="#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"; "settingDefinitionId"="fslogix_odfc_includeoutlook"; "choiceSettingValue"=@{"value"="fslogix_odfc_includeoutlook_1"} }
                            ) }
                        )
                    }
                }
            )
        }
    )

    foreach ($pol in $policies) {
        $file = ".\FSLogix-Intune-$($pol.Name).json"
        $pol | ConvertTo-Json -Depth 15 | Set-Content $file -Encoding UTF8
        fx-ok "Intune policy JSON exported" $file
    }
}

function Export-FSLogixGPOSettings {
    fx-sec "Generating Group Policy (ADMX) Settings Reference"
    $cfg  = $Script:Cfg.FSLogix
    $file = ".\FSLogix-GPO-Settings.txt"
    $lines = @(
        "# ═══════════════════════════════════════════════════════════════════════",
        "# BDF FSLogix Group Policy Settings Reference",
        "# Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        "# Apply via: GPMC → [Your GPO] → Computer Configuration → Administrative Templates → FSLogix",
        "# ═══════════════════════════════════════════════════════════════════════",
        "",
        "# PROFILE CONTAINERS",
        "# GPO Path: FSLogix → Profile Containers",
        "Enabled                              = 1",
        "VHDLocations (E3 — apply to E3 OU) = $($cfg.E3VHDLocation)",
        "VHDLocations (F1 — apply to F1 OU) = $($cfg.F1VHDLocation)",
        "SizeInMBs (E3)                       = $($cfg.E3SizeGB * 1024)",
        "SizeInMBs (F1)                       = $($cfg.F1SizeGB * 1024)",
        "VolumeType                           = VHDX",
        "IsDynamic                            = 1",
        "FlipFlopProfileDirectoryName         = 1",
        "DeleteLocalProfileWhenVHDShouldApply = 1",
        "ConcurrentUserSessions               = 1",
        "AccessNetworkAsComputerObject        = 1   ← REQUIRED for Azure AD Kerberos",
        "LockedRetryCount                     = 3",
        "LockedRetryInterval                  = 15",
        "",
        "# OFFICE CONTAINERS (ODFC)",
        "# GPO Path: FSLogix → Office 365 Container",
        "Enabled (ODFC)   = 1",
        "VHDLocations     = $($cfg.ODFCLocation)",
        "SizeInMBs        = $($cfg.ODFCSizeGB * 1024)",
        "IncludeTeams     = 1   ← CRITICAL — routes Teams cache to ODFC",
        "IncludeOneDrive  = 1",
        "IncludeOutlook   = 1",
        "",
        "# LOGGING",
        "LogDir   = C:\ProgramData\FSLogix\Logs",
        "LogLevel = 2   (1=Errors only in production; 2=Info for troubleshooting)"
    )
    if ($cfg.CloudCache.Enabled -and $cfg.CloudCache.Locations.Count -gt 0) {
        $lines += ""
        $lines += "# CLOUD CACHE (HIGH AVAILABILITY)"
        $lines += "# Replace VHDLocations with CCDLocations when using Cloud Cache:"
        foreach ($loc in $cfg.CloudCache.Locations) { $lines += "CCDLocations += type=smb,connectionString=$loc" }
    }
    $lines | Set-Content $file -Encoding UTF8
    fx-ok "GPO settings reference exported" $file
}

function Export-FSLogixRegistryScript {
    fx-sec "Generating Registry Deployment Script"
    $cfg  = $Script:Cfg.FSLogix
    $file = ".\Deploy-FSLogix-Registry.ps1"
    $script = @"
<#
.SYNOPSIS  Deploy FSLogix settings via registry — run on each session host during image build or via Intune Win32 App
.NOTES     Generated by BDF AVD AIO — $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
           Run as: SYSTEM (GPO startup script, Intune Management Extension, or image bake-in)
#>

`$E3VHDLocation  = '$($cfg.E3VHDLocation)'
`$F1VHDLocation  = '$($cfg.F1VHDLocation)'
`$ODFCLocation   = '$($cfg.ODFCLocation)'
`$E3SizeMB       = $($cfg.E3SizeGB * 1024)
`$F1SizeMB       = $($cfg.F1SizeGB * 1024)
`$ODFCSizeMB     = $($cfg.ODFCSizeGB * 1024)

function Set-FXReg { param([string]`$Path,[string]`$Name,[object]`$Value,[string]`$Type="DWord")
    if (-not (Test-Path `$Path)) { New-Item -Path `$Path -Force | Out-Null }
    Set-ItemProperty -Path `$Path -Name `$Name -Value `$Value -Type `$Type -Force
    Write-Host "  SET: `$(`$Path -split '\\' | Select-Object -Last 1) -> `$Name = `$Value" -ForegroundColor Cyan
}

Write-Host "" ; Write-Host "  BDF FSLogix Registry Configuration" -ForegroundColor Blue ; Write-Host ""

# ── Determine user type (E3 vs F1) from group membership ──────────────────
# NOTE: For single golden image serving both E3 and F1, configure a SEPARATE
# GPO/Intune profile per group and let Windows apply the correct VHDLocations.
# The registry script below configures one target — adjust `$VHDLocations per OU/group.

`$ProfileReg = 'HKLM:\SOFTWARE\FSLogix\Profiles'
`$ODFCReg    = 'HKLM:\SOFTWARE\Policies\FSLogix\ODFC'
`$LogReg     = 'HKLM:\SOFTWARE\FSLogix\Logging'

# ── PROFILE CONTAINER ─────────────────────────────────────────────────────
Write-Host "  ── Profile Container" -ForegroundColor Magenta
Set-FXReg `$ProfileReg "Enabled"                              1
Set-FXReg `$ProfileReg "VHDLocations"                        `$E3VHDLocation      "String"
Set-FXReg `$ProfileReg "SizeInMBs"                           `$E3SizeMB
Set-FXReg `$ProfileReg "VolumeType"                          "VHDX"               "String"
Set-FXReg `$ProfileReg "IsDynamic"                           1
Set-FXReg `$ProfileReg "FlipFlopProfileDirectoryName"         1
Set-FXReg `$ProfileReg "DeleteLocalProfileWhenVHDShouldApply" 1
Set-FXReg `$ProfileReg "ConcurrentUserSessions"              1
Set-FXReg `$ProfileReg "AccessNetworkAsComputerObject"        1
Set-FXReg `$ProfileReg "LockedRetryCount"                    3
Set-FXReg `$ProfileReg "LockedRetryInterval"                 15
Set-FXReg `$ProfileReg "ReAttachIntervalSeconds"             15
Set-FXReg `$ProfileReg "ReAttachRetryCount"                  3
Set-FXReg `$ProfileReg "PreventLoginWithFailure"             0
Set-FXReg `$ProfileReg "PreventLoginWithTempProfile"         0

# ── OFFICE CONTAINER (ODFC) ───────────────────────────────────────────────
Write-Host "  ── Office Container (ODFC)" -ForegroundColor Magenta
Set-FXReg `$ODFCReg "Enabled"                   1
Set-FXReg `$ODFCReg "VHDLocations"              `$ODFCLocation   "String"
Set-FXReg `$ODFCReg "SizeInMBs"                 `$ODFCSizeMB
Set-FXReg `$ODFCReg "VolumeType"                "VHDX"           "String"
Set-FXReg `$ODFCReg "IsDynamic"                 1
Set-FXReg `$ODFCReg "IncludeTeams"              1
Set-FXReg `$ODFCReg "IncludeOneDrive"           1
Set-FXReg `$ODFCReg "IncludeOutlook"            1
Set-FXReg `$ODFCReg "IncludeOutlookPersonalization" 1
Set-FXReg `$ODFCReg "FlipFlopProfileDirectoryName" 1
Set-FXReg `$ODFCReg "AccessNetworkAsComputerObject" 1

$(if ($cfg.CloudCache.Enabled) {
    "# ── CLOUD CACHE ────────────────────────────────────────────────────────
`$CCDVal = '$($cfg.CloudCache.Locations | ForEach-Object {"type=smb,connectionString=$_"} | Join-String -Separator '|')'
Set-FXReg `$ProfileReg 'CCDLocations' `$CCDVal 'String'
Write-Host '  Cloud Cache configured: $($cfg.CloudCache.Locations.Count) locations' -ForegroundColor Green"
} else { "# Cloud Cache disabled — using direct VHDLocations" })

# ── LOGGING ───────────────────────────────────────────────────────────────
Write-Host "  ── Logging" -ForegroundColor Magenta
Set-FXReg `$LogReg "LogDir"    "C:\ProgramData\FSLogix\Logs" "String"
Set-FXReg `$LogReg "LogLevel"  2
Set-FXReg `$LogReg "LogMaxSize" 100

Write-Host ""
Write-Host "  FSLogix registry configuration complete." -ForegroundColor Green
Write-Host "  Log directory: C:\ProgramData\FSLogix\Logs" -ForegroundColor Cyan
Write-Host "  Validate: Check Event Viewer → Applications and Services → Microsoft-FSLogix-Apps → Operational" -ForegroundColor Cyan
"@
    $script | Set-Content $file -Encoding UTF8
    fx-ok "Registry deployment script exported" $file
}
#endregion

#region ════════════════════════════════════════════════════════════════════
# REDIRECTIONS.XML BUILDER
#════════════════════════════════════════════════════════════════════════════

function Build-RedirectionsXml {
    fx-hdr "REDIRECTIONS.XML BUILDER"
    fx-info "Redirections.xml controls what is excluded from the profile container"
    fx-info "or redirected to alternate local paths (not roamed)"
    Write-Host ""

    $selected = [System.Collections.ArrayList]@()

    foreach ($category in $FXExclusions.GetEnumerator()) {
        fx-sec $category.Key
        foreach ($excl in $category.Value) {
            Write-Host ("    {0,-55} [{1}]" -f $excl.Path, $excl.Type) -ForegroundColor $FC.Menu
            Write-Host ("    {0}" -f $excl.Desc) -ForegroundColor $FC.Muted
        }
        if (Read-YN "Include ALL $($category.Key) exclusions?" $true) {
            $selected.AddRange($category.Value)
            fx-ok "Added $($category.Value.Count) exclusions" $category.Key
        } else {
            # Allow selecting individual exclusions
            $i = 1
            $items = $category.Value
            foreach ($ex in $items) {
                Write-Host ("    {0}. {1}" -f $i, ($ex.Path -split '\\'|Select-Object -Last 1)) -ForegroundColor $FC.Muted
                $i++
            }
            Write-Host "    Enter numbers to include (e.g. 1,3,4) or blank to skip all:"
            $picks = (Read-Inp "Selection (blank=skip all)" -D "").Split(",")
            foreach ($p in $picks) {
                $idx = ([int]$p.Trim()) - 1
                if ($idx -ge 0 -and $idx -lt $items.Count) { $selected.Add($items[$idx]) | Out-Null }
            }
        }
    }

    # Custom exclusion entry
    while (Read-YN "Add a custom exclusion path?" $false) {
        $cPath = Read-Inp "Exclusion path (use %LOCALAPPDATA%, %APPDATA%, %USERPROFILE% etc.)"
        $cType = Read-Inp "Type (Directory or File)" -V @("Directory","File","directory","file") -D "Directory"
        $cDesc = Read-Inp "Description"
        $selected.Add(@{Path=$cPath; Type=$cType; Desc=$cDesc}) | Out-Null
        fx-ok "Custom exclusion added" $cPath
    }

    # Build XML
    $xmlLines = @(
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<!--',
        "  BDF AVD FSLogix Redirections.xml",
        "  Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        "  Deploy to: \\$($Script:Cfg.Storage.Account.Name).file.core.windows.net\$($Script:Cfg.Storage.Shares.E3) (beside VHDX files)",
        "  OR bake into session host golden image: C:\Program Files\FSLogix\Apps\Rules\",
        '-->',
        '<FrxProfileFolderRedirection ExcludeCommonFolders="49">',
        '  <Excludes>'
    )
    foreach ($excl in $selected) {
        $type = if ($excl.Type -like "Dir*") { "Directory" } else { "File" }
        $xmlLines += "    <!-- $($excl.Desc) -->"
        $xmlLines += "    <Exclude Copy=""0"">$($excl.Path)</Exclude>"
    }
    $xmlLines += '  </Excludes>'
    $xmlLines += '  <Redirections>'
    $xmlLines += '    <!-- Redirect Downloads to local temp (not persisted to profile) -->'
    $xmlLines += '    <!-- Uncomment if needed: <Redirect Copy="0">%USERPROFILE%\Downloads</Redirect> -->'
    $xmlLines += '  </Redirections>'
    $xmlLines += '</FrxProfileFolderRedirection>'

    $file = ".\Redirections.xml"
    $xmlLines | Set-Content $file -Encoding UTF8
    fx-ok "Redirections.xml created" $file
    fx-info "Total exclusions" "$($selected.Count)"
    fx-info "Deploy location" "Profile share root (\\storageaccount.file.core.windows.net\profiles-e3\)"
    fx-info "Alternative" "C:\Program Files\FSLogix\Apps\Rules\ on session hosts"

    # Also generate deployment script
    $deployScript = @"
# Deploy Redirections.xml to Azure Files profile shares
# Run this on a machine with access to Azure Files (or from Azure Cloud Shell)

`$StorageAccountName = '$($Script:Cfg.Storage.Account.Name)'
`$ResourceGroupName  = '$($Script:Cfg.Storage.Account.ResourceGroup)'
`$Shares = @('$($Script:Cfg.Storage.Shares.E3)', '$($Script:Cfg.Storage.Shares.F1)')

`$sa  = Get-AzStorageAccount -Name `$StorageAccountName -ResourceGroupName `$ResourceGroupName
`$ctx = `$sa.Context

foreach (`$share in `$Shares) {
    Write-Host "Uploading Redirections.xml to `$share..."
    Set-AzStorageFileContent -ShareName `$share -Source '.\Redirections.xml' -Path 'Redirections.xml' -Context `$ctx -Force
    Write-Host "  Uploaded to: \\`$StorageAccountName.file.core.windows.net\`$share\Redirections.xml"
}
Write-Host "Done. FSLogix will apply exclusions on next user profile mount."
"@
    $deployScript | Set-Content ".\Deploy-Redirections.ps1" -Encoding UTF8
    fx-ok "Deployment script" ".\Deploy-Redirections.ps1"
    Pause-Screen
}
#endregion

#region ════════════════════════════════════════════════════════════════════
# FSLOGIX APP MASKING
#════════════════════════════════════════════════════════════════════════════

function Show-AppMaskingMenu {
    fx-hdr "FSLOGIX APP MASKING MANAGER" $FC.Purple

    while ($true) {
        fx-sec "App Masking Options"
        fx-info "App Masking hides applications, registry keys, and files based on group membership"
        fx-info "One golden image — different app visibility per user type (E3 vs F1)"
        Write-Host ""
        Write-Host "  1.  View / apply BDF preset masking rules" -ForegroundColor $FC.Menu
        Write-Host "  2.  Build custom masking rule (interactive)" -ForegroundColor $FC.Menu
        Write-Host "  3.  Generate rule file (.fxr) for upload" -ForegroundColor $FC.New
        Write-Host "  4.  Deploy rules to session host (direct registry)" -ForegroundColor $FC.Azure
        Write-Host "  5.  Export Intune Win32 app package (rule deployment)" -ForegroundColor $FC.Menu
        Write-Host "  6.  View current rule assignments" -ForegroundColor $FC.Menu
        Write-Host "  0.  Back" -ForegroundColor $FC.Muted
        $ch = Read-Inp "Choice" -V @("0","1","2","3","4","5","6")
        switch ($ch) {
            "1" { Apply-AppMaskingPresets }
            "2" { Build-CustomMaskingRule }
            "3" { Export-AppMaskingRuleFile }
            "4" { Deploy-AppMaskingRules }
            "5" { Export-AppMaskingIntunePackage }
            "6" { Show-MaskingRuleAssignments }
            "0" { return }
        }
    }
}

function Apply-AppMaskingPresets {
    fx-sec "BDF App Masking Presets"

    $presets = @(
        @{
            ID = "hide-office-f1"
            Name = "Hide Microsoft 365 Desktop Apps from F1 Frontline"
            Desc = "F1 users have web-only Office entitlement — hide Word/Excel/PowerPoint/OneNote desktop apps"
            Rules = @(
                @{Type="File";       Path='%ProgramFiles%\Microsoft Office\root\Office16\WINWORD.EXE';   Action="Hide"; App="Word"}
                @{Type="File";       Path='%ProgramFiles%\Microsoft Office\root\Office16\EXCEL.EXE';     Action="Hide"; App="Excel"}
                @{Type="File";       Path='%ProgramFiles%\Microsoft Office\root\Office16\POWERPNT.EXE';  Action="Hide"; App="PowerPoint"}
                @{Type="File";       Path='%ProgramFiles%\Microsoft Office\root\Office16\ONENOTE.EXE';   Action="Hide"; App="OneNote"}
                @{Type="File";       Path='%ProgramFiles%\Microsoft Office\root\Office16\OUTLOOK.EXE';   Action="Hide"; App="Outlook"}
                @{Type="File";       Path='%ProgramFiles%\Microsoft Office\root\Office16\MSACCESS.EXE';  Action="Hide"; App="Access"}
                @{Type="File";       Path='%ProgramFiles%\Microsoft Office\root\Office16\MSPUB.EXE';     Action="Hide"; App="Publisher"}
                @{Type="ShortcutFolder"; Path='%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Microsoft Office'; Action="Hide"; App="Office Start Menu Folder"}
            )
            HideFrom  = "F1 Frontline Users"
            ShowFor   = "E3 Office Users"
        },
        @{
            ID = "hide-admin-tools-standard"
            Name = "Hide Admin/IT Tools from Standard Users"
            Desc = "Hide PowerShell ISE, Registry Editor, Disk Management from non-admin users"
            Rules = @(
                @{Type="File";   Path='%SystemRoot%\System32\regedt32.exe';           Action="Hide"; App="Registry Editor"}
                @{Type="File";   Path='%SystemRoot%\System32\mmc.exe';                Action="Hide"; App="MMC Console"}
                @{Type="File";   Path='%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell_ise.exe'; Action="Hide"; App="PowerShell ISE"}
                @{Type="File";   Path='%SystemRoot%\System32\cmd.exe';                Action="Hide"; App="Command Prompt"}
                @{Type="File";   Path='%SystemRoot%\System32\eventvwr.exe';           Action="Hide"; App="Event Viewer"}
                @{Type="File";   Path='%SystemRoot%\System32\compmgmt.msc';           Action="Hide"; App="Computer Management"}
                @{Type="File";   Path='%SystemRoot%\System32\diskmgmt.msc';           Action="Hide"; App="Disk Management"}
                @{Type="File";   Path='%SystemRoot%\System32\taskmgr.exe';            Action="Hide"; App="Task Manager"}
            )
            HideFrom  = "Standard Users (non-admin)"
            ShowFor   = "IT Admins"
        },
        @{
            ID = "hide-visio-project-unlicensed"
            Name = "Hide Visio and Project from Unlicensed Users"
            Desc = "Visio and Project require separate licenses — hide from users without those licenses"
            Rules = @(
                @{Type="File";   Path='%ProgramFiles%\Microsoft Office\root\Office16\VISIO.EXE';     Action="Hide"; App="Visio"}
                @{Type="File";   Path='%ProgramFiles%\Microsoft Office\root\Office16\WINPROJ.EXE';   Action="Hide"; App="Project"}
                @{Type="File";   Path='%ProgramFiles%\Microsoft Office\root\Office16\MSPROJECT.EXE'; Action="Hide"; App="Project (alt)"}
                @{Type="ShortcutFolder"; Path='%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016\Visio*'; Action="Hide"; App="Visio Shortcuts"}
                @{Type="ShortcutFolder"; Path='%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016\Project*'; Action="Hide"; App="Project Shortcuts"}
            )
            HideFrom  = "Users without Visio/Project license"
            ShowFor   = "Licensed Visio/Project users"
        },
        @{
            ID = "hide-store-apps-office-workers"
            Name = "Hide Store/POS Apps from E3 Office Workers"
            Desc = "Retail POS or store-specific apps not needed by HQ staff"
            Rules = @(
                @{Type="File";   Path='%ProgramFiles%\RetailPOS\pos.exe';             Action="Hide"; App="Retail POS App"}
                @{Type="File";   Path='%ProgramFiles%\InventoryManager\inventory.exe'; Action="Hide"; App="Inventory Manager"}
            )
            HideFrom  = "E3 Office Workers (HQ)"
            ShowFor   = "F1 Frontline Store Workers"
        }
    )

    $i = 1
    foreach ($p in $presets) {
        Write-Host ""
        Write-Host ("  {0}. {1}" -f $i, $p.Name) -ForegroundColor $FC.Accent
        Write-Host ("     {0}" -f $p.Desc) -ForegroundColor $FC.Muted
        Write-Host ("     Hides from: {0}  |  Visible to: {1}" -f $p.HideFrom, $p.ShowFor) -ForegroundColor $FC.Info
        $i++
    }
    Write-Host ""
    Write-Host "  Select presets to apply (comma-separated numbers, e.g. 1,2,3):" -ForegroundColor $FC.Menu
    $picks = (Read-Inp "Selection (Enter=all, blank=skip)").Split(",") | Where-Object { $_ -match "^\d+$" }
    if (-not $picks) { $picks = 1..($presets.Count) }

    $allRules = [System.Collections.ArrayList]@()
    foreach ($p in $picks) {
        $idx = ([int]$p.Trim()) - 1
        if ($idx -ge 0 -and $idx -lt $presets.Count) {
            $preset = $presets[$idx]
            $allRules.AddRange($preset.Rules)
            fx-ok "Preset added" $preset.Name
            $Script:Cfg.AppMasking.RuleFiles += @{ PresetId=$preset.ID; Name=$preset.Name; RuleCount=$preset.Rules.Count }
        }
    }
    if ($allRules.Count -gt 0) {
        $Script:Cfg.AppMasking.PendingRules = $allRules
        fx-ok "Total masking rules queued" "$($allRules.Count)"
        Save-FXConfig
        if (Read-YN "Generate rule files now?" $true) { Export-AppMaskingRuleFile }
    }
    Pause-Screen
}

function Build-CustomMaskingRule {
    fx-sec "Custom App Masking Rule Builder"
    $rules = [System.Collections.ArrayList]@()

    do {
        Write-Host ""
        Write-Host "  Rule Types:" -ForegroundColor $FC.Accent
        Write-Host "  1. File (hide a specific executable or file)" -ForegroundColor $FC.Menu
        Write-Host "  2. Folder (hide a directory and all contents)" -ForegroundColor $FC.Menu
        Write-Host "  3. Registry Key (hide a registry key from applications)" -ForegroundColor $FC.Menu
        Write-Host "  4. Shortcut (hide Start Menu or Desktop shortcut)" -ForegroundColor $FC.Menu
        $type = Read-Inp "Rule type" -V @("1","2","3","4") -D "1"
        $typeMap = @{"1"="File";"2"="Folder";"3"="Registry";"4"="Shortcut"}
        $ruleType = $typeMap[$type]

        $path = Read-Inp "Path (use %ProgramFiles%, %SystemRoot%, %APPDATA% etc.)"
        $app  = Read-Inp "Application name (friendly label)"
        $action = Read-Inp "Action (Hide/Redirect)" -V @("Hide","hide","Redirect","redirect") -D "Hide"

        $rule = @{ Type=$ruleType; Path=$path; Action=$action; App=$app }
        $rules.Add($rule) | Out-Null
        fx-ok "Rule added" "$action $app ($ruleType)"
    } until (-not (Read-YN "Add another rule?" $false))

    $ruleName = Read-Inp "Rule set name" -D "Custom-BDF-Rules"
    $Script:Cfg.AppMasking.RuleFiles += @{ PresetId="custom"; Name=$ruleName; RuleCount=$rules.Count }
    $Script:Cfg.AppMasking.PendingRules += $rules
    Save-FXConfig
    fx-ok "Custom rule set created" "$ruleName ($($rules.Count) rules)"
    Pause-Screen
}

function Export-AppMaskingRuleFile {
    fx-sec "Exporting FSLogix App Masking Rule Files (.fxr / PowerShell)"
    $pending = $Script:Cfg.AppMasking.PendingRules
    if (-not $pending -or @($pending).Count -eq 0) {
        fx-warn "No rules configured — run presets or custom builder first"; Pause-Screen; return
    }

    $fileName = Read-Inp "Output rule set name" -D "BDF-AVD-Rules"

    # Generate PowerShell-based rule creation (uses FSLogix Rule Editor COM or direct XML)
    $script = @"
<#
.SYNOPSIS  Deploy FSLogix App Masking rules via PowerShell
.NOTES     Run on session hosts during image build (as SYSTEM)
           OR deploy via Intune Win32 App with detection rule
           Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
#>

`$RuleDir  = 'C:\Program Files\FSLogix\Apps\Rules'
`$RuleName = '$fileName'

# Ensure FSLogix is installed
if (-not (Test-Path 'C:\Program Files\FSLogix\Apps\frx.exe')) {
    Write-Error "FSLogix not installed. Install FSLogix agent first."
    exit 1
}

# Create rules directory if needed
if (-not (Test-Path `$RuleDir)) { New-Item -Path `$RuleDir -ItemType Directory -Force | Out-Null }

# ── Method A: Use FSLogix frxruleeditor.exe (if GUI available on build host) ──
# frxruleeditor.exe creates .fxr binary rule files
# Uncomment and adjust if running interactively on build workstation:
# & 'C:\Program Files\FSLogix\Apps\frxruleeditor.exe' ...

# ── Method B: Direct XML rule file creation ───────────────────────────────
# FSLogix supports XML-based .fxr files for automation

`$xmlContent = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<!--
  BDF AVD FSLogix App Masking Rules
  Rule: $fileName
  Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
  Documentation: https://docs.microsoft.com/en-us/fslogix/app-masking
-->
<AppMaskingPolicy name="$fileName" version="1">

$(foreach ($rule in $pending) {
"  <!-- $($rule.App) -->
  <Rule type=""$($rule.Type)"" action=""$($rule.Action)"">
    <Path>$($rule.Path)</Path>
  </Rule>"
})

</AppMaskingPolicy>
'@

`$outFile = Join-Path `$RuleDir "`$RuleName.fxr"
`$xmlContent | Set-Content `$outFile -Encoding UTF8
Write-Host "  Rule file created: `$outFile" -ForegroundColor Green

# ── Group Assignment ────────────────────────────────────────────────────
# Create corresponding .fxa (assignment) file to map rules to AD groups:
# NOTE: Object IDs below — replace with actual Entra ID group Object IDs

`$assignContent = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<AppMaskingAssignment>
  <!-- Users in this group CANNOT see the masked apps -->
  <Assignment>
    <UserGroup sid="S-1-12-1-XXXXXXXXX" comment="F1 Frontline Users (replace with actual SID or OID)" />
  </Assignment>
</AppMaskingAssignment>
'@

`$asnFile = Join-Path `$RuleDir "`$RuleName.fxa"
`$assignContent | Set-Content `$asnFile -Encoding UTF8
Write-Host "  Assignment file created: `$asnFile" -ForegroundColor Green
Write-Host ""
Write-Host "  IMPORTANT: Update .fxa file with actual Entra ID group SIDs before deploying." -ForegroundColor Yellow
Write-Host "  Get SID: (Get-AzADGroup -DisplayName 'F1 Frontline Workers').SecurityIdentifier" -ForegroundColor Cyan
"@

    $script | Set-Content ".\Deploy-AppMasking-$fileName.ps1" -Encoding UTF8
    fx-ok "App Masking deployment script" ".\Deploy-AppMasking-$fileName.ps1"

    # Also write a reference XML
    $xmlRef = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<!--
  BDF AVD App Masking Rule Reference — $fileName
  Copy to session hosts: C:\Program Files\FSLogix\Apps\Rules\
  Corresponding .fxa assignment file must also be present
-->
<AppMaskingPolicy name="$fileName" version="1">
$(foreach ($r in $pending) {"  <!-- $($r.App) -->`n  <Rule type=""$($r.Type)"" action=""$($r.Action)""><Path>$($r.Path)</Path></Rule>"})
</AppMaskingPolicy>
"@
    $xmlRef | Set-Content ".\$fileName.fxr.xml" -Encoding UTF8
    fx-ok "Rule XML reference" ".\$fileName.fxr.xml"

    $Script:Cfg.AppMasking.Deployed = $true
    Save-FXConfig
    fx-info "Next steps" "Update .fxa group SIDs, deploy to Rules folder on golden image"
    Pause-Screen
}

function Export-AppMaskingIntunePackage {
    fx-sec "App Masking Intune Win32 Deployment Package"
    $pkgDir = ".\AppMasking-IntunePackage"
    New-Item -Path $pkgDir -ItemType Directory -Force | Out-Null

    # Install script
    @"
# Install.ps1 — Copy FSLogix App Masking rule files to session host
`$dest = 'C:\Program Files\FSLogix\Apps\Rules'
New-Item -Path `$dest -ItemType Directory -Force -EA SilentlyContinue | Out-Null
Copy-Item -Path ".\*.fxr.xml" -Destination `$dest -Force
Copy-Item -Path ".\*.fxa"     -Destination `$dest -Force
Write-Host "FSLogix App Masking rules deployed."
exit 0
"@ | Set-Content "$pkgDir\Install.ps1" -Encoding UTF8

    # Uninstall script
    @"
# Uninstall.ps1
Remove-Item 'C:\Program Files\FSLogix\Apps\Rules\BDF-AVD-Rules*' -Force -EA SilentlyContinue
exit 0
"@ | Set-Content "$pkgDir\Uninstall.ps1" -Encoding UTF8

    # Detection script (Intune Win32 detection rule)
    @"
# Detection.ps1
if (Test-Path 'C:\Program Files\FSLogix\Apps\Rules\BDF-AVD-Rules.fxr.xml') { exit 0 } else { exit 1 }
"@ | Set-Content "$pkgDir\Detection.ps1" -Encoding UTF8

    # IntuneWin packaging instructions
    @"
# HOW TO PACKAGE FOR INTUNE WIN32 APP
# 1. Download Microsoft Win32 Content Prep Tool:
#    https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool
#
# 2. Run: IntuneWinAppUtil.exe -c .\AppMasking-IntunePackage -s Install.ps1 -o .\Output
#
# 3. Upload .intunewin to Intune: Apps > Windows > Add > Windows app (Win32)
#    - Install command: powershell.exe -ExecutionPolicy Bypass -File Install.ps1
#    - Uninstall command: powershell.exe -ExecutionPolicy Bypass -File Uninstall.ps1
#    - Detection: Custom script (Detection.ps1)
#    - Assignment: Required > [E3 Office Workers Device Group] or [F1 Frontline Device Group]
"@ | Set-Content "$pkgDir\PACKAGING-INSTRUCTIONS.txt" -Encoding UTF8

    fx-ok "Intune package files created" $pkgDir
    Pause-Screen
}

function Show-MaskingRuleAssignments {
    fx-sec "App Masking Rule Assignments"
    if (@($Script:Cfg.AppMasking.RuleFiles).Count -eq 0) {
        fx-warn "No masking rules configured yet" ""; Pause-Screen; return
    }
    foreach ($rf in $Script:Cfg.AppMasking.RuleFiles) {
        Write-Host ("  {0,-35} Rules: {1}" -f $rf.Name, $rf.RuleCount) -ForegroundColor $FC.Menu
    }
    fx-info "Rules are stored in" "C:\Program Files\FSLogix\Apps\Rules on session hosts"
    Pause-Screen
}

function Deploy-AppMaskingRules {
    fx-sec "Direct Deploy to Session Hosts via PowerShell Remoting"
    fx-warn "Requires WinRM access to session hosts (or use Intune/GPO instead)" ""
    if (-not (Read-YN "Continue with direct deployment?" $false)) { return }
    $hp = Read-Inp "Host Pool name" -D $Script:Cfg.HostPools.E3.Name
    $rg = Read-Inp "Resource Group" -D $Script:Cfg.RG.AVD.Name
    $hosts = @(Get-AzWvdSessionHost -HostPoolName $hp -ResourceGroupName $rg -EA SilentlyContinue)
    foreach ($h in $hosts) {
        $vmName = ($h.Name -split "/")[-1]
        fx-step "Deploying to $vmName"
        Invoke-AzVMRunCommand -ResourceGroupName $rg -VMName $vmName `
            -CommandId RunPowerShellScript `
            -ScriptPath ".\Deploy-AppMasking-BDF-AVD-Rules.ps1" `
            -EA SilentlyContinue | Out-Null
        fx-ok "Deployed" $vmName
    }
    Pause-Screen
}
#endregion

#region ════════════════════════════════════════════════════════════════════
# APP ATTACH / MSIX APP ATTACH
#════════════════════════════════════════════════════════════════════════════

function Show-AppAttachMenu {
    fx-hdr "APP ATTACH & MSIX APP ATTACH MANAGER" $FC.Azure

    while ($true) {
        fx-sec "App Attach — Current Status"

        # Status summary
        $shareExists = $false
        $sa = Get-AzStorageAccount -Name $Script:Cfg.Storage.Account.Name `
                  -ResourceGroupName $Script:Cfg.Storage.Account.ResourceGroup -EA SilentlyContinue
        if ($sa) {
            $shareExists = $null -ne (Get-AzStorageShare -Name $Script:Cfg.AppAttach.ShareName -Context $sa.Context -EA SilentlyContinue)
        }
        $pkgCount = @(Get-AzWvdAppAttachPackage -ResourceGroupName $Script:Cfg.RG.AVD.Name -EA SilentlyContinue).Count

        fx-info "App Attach Share"     (if ($shareExists) {"✔ $($Script:Cfg.AppAttach.ShareName)"} else {"✖ Not created"})
        fx-info "Packages registered" "$pkgCount"
        fx-info "Host Pools assigned" "$(@($Script:Cfg.AppAttach.Packages).Count)"

        Write-Host ""
        Write-Host "  1.  Setup App Attach Azure Files share" -ForegroundColor $FC.Menu
        Write-Host "  2.  Upload MSIX package (.msix / .msixbundle)" -ForegroundColor $FC.Menu
        Write-Host "  3.  Register App Attach package with AVD" -ForegroundColor $FC.New
        Write-Host "  4.  Manage packages (stage/activate/deregister)" -ForegroundColor $FC.Menu
        Write-Host "  5.  Assign packages to Application Groups" -ForegroundColor $FC.Menu
        Write-Host "  6.  Create MSIX image from installer (packaging guide)" -ForegroundColor $FC.Azure
        Write-Host "  7.  Monitor App Attach health" -ForegroundColor $FC.Menu
        Write-Host "  8.  Troubleshooting & diagnostics" -ForegroundColor $FC.Warn
        Write-Host "  9.  Generate session host readiness script" -ForegroundColor $FC.Menu
        Write-Host "  0.  Back" -ForegroundColor $FC.Muted
        $ch = Read-Inp "Choice" -V @("0","1","2","3","4","5","6","7","8","9")
        switch ($ch) {
            "1" { Setup-AppAttachShare }
            "2" { Upload-MSIXPackage }
            "3" { Register-AppAttachPackage }
            "4" { Manage-AppAttachPackages }
            "5" { Assign-AppAttachToAppGroup }
            "6" { Show-MSIXPackagingGuide }
            "7" { Monitor-AppAttachHealth }
            "8" { Troubleshoot-AppAttach }
            "9" { Export-SessionHostReadinessScript }
            "0" { return }
        }
    }
}

function Setup-AppAttachShare {
    fx-sec "App Attach Azure Files Share Setup"
    fx-info "App Attach packages (.msix, .msixbundle, .vhd, .vhdx, .cim) are stored on Azure Files"
    fx-info "The share must be accessible from session hosts with Kerberos/NTFS permissions"
    Write-Host ""

    $sa = Get-AzStorageAccount -Name $Script:Cfg.Storage.Account.Name `
              -ResourceGroupName $Script:Cfg.Storage.Account.ResourceGroup -EA SilentlyContinue
    if (-not $sa) {
        fx-fail "Storage account not found" $Script:Cfg.Storage.Account.Name
        fx-warn "Deploy Azure Files component first (Main Menu → Component Deployment)" ""
        Pause-Screen; return
    }

    $shareName = Read-Inp "App Attach share name" -D "appattach"
    $Script:Cfg.AppAttach.ShareName = $shareName

    $existing = Get-AzStorageShare -Name $shareName -Context $sa.Context -EA SilentlyContinue
    if ($existing) {
        fx-ok "Share already exists" "$shareName"
    } else {
        $quotaGB  = Read-Inp "Share size (GB)" -D "100"
        New-AzStorageShare -Name $shareName -Context $sa.Context -QuotaGiB ([int]$quotaGB) | Out-Null
        fx-ok "App Attach share created" "$shareName ($quotaGB GB)"
    }

    # RBAC — session hosts need read access
    fx-step "Configuring RBAC for session hosts..."
    fx-info "Session hosts read packages via Storage File Data SMB Share Reader role"
    fx-info "This requires the VM system-assigned managed identity to be assigned this role"

    $shareScope = "$($sa.Id)/fileServices/default/fileshares/$shareName"
    Write-Host ""
    Write-Host "  To grant session host VM identities access, run:" -ForegroundColor $FC.Muted
    Write-Host "  (Get each VM's identity ObjectId and assign Storage File Data SMB Share Reader)" -ForegroundColor $FC.Muted
    Write-Host ""
    Write-Host '  # Example:' -ForegroundColor $FC.Muted
    Write-Host '  $vmId = (Get-AzVM -Name "vm-avd-e3-bdfpoc-1" -ResourceGroupName "rg-avd-bdf-poc").Identity.PrincipalId' -ForegroundColor $FC.Azure
    Write-Host "  New-AzRoleAssignment -ObjectId `$vmId -RoleDefinitionName 'Storage File Data SMB Share Reader' -Scope '$shareScope'" -ForegroundColor $FC.Azure

    # Generate mass-assignment script
    $assignScript = @"
# Assign Storage File Data SMB Share Reader to all AVD session host managed identities
# Run after session hosts are deployed

`$hostPoolNames = @('$($Script:Cfg.HostPools.E3.Name)', '$($Script:Cfg.HostPools.F1.Name)')
`$hpRg          = '$($Script:Cfg.RG.AVD.Name)'
`$shareScope     = '$shareScope'

foreach (`$hp in `$hostPoolNames) {
    `$hosts = Get-AzWvdSessionHost -HostPoolName `$hp -ResourceGroupName `$hpRg -EA SilentlyContinue
    foreach (`$h in `$hosts) {
        `$vmName = (`$h.Name -split '/')[-1]
        `$vm     = Get-AzVM -Name `$vmName -ResourceGroupName `$hpRg -EA SilentlyContinue
        if (`$vm -and `$vm.Identity.PrincipalId) {
            New-AzRoleAssignment -ObjectId `$vm.Identity.PrincipalId `
                -RoleDefinitionName 'Storage File Data SMB Share Reader' `
                -Scope `$shareScope -EA SilentlyContinue | Out-Null
            Write-Host "  Assigned Storage Reader to: `$vmName" -ForegroundColor Green
        }
    }
}
Write-Host "Storage RBAC assignment complete."
"@
    $assignScript | Set-Content ".\Set-AppAttachStoragePermissions.ps1" -Encoding UTF8
    fx-ok "RBAC assignment script" ".\Set-AppAttachStoragePermissions.ps1"

    $Script:Cfg.AppAttach.Enabled = $true
    Save-FXConfig
    Pause-Screen
}

function Upload-MSIXPackage {
    fx-sec "Upload MSIX Package to Azure Files"
    $sa = Get-AzStorageAccount -Name $Script:Cfg.Storage.Account.Name `
              -ResourceGroupName $Script:Cfg.Storage.Account.ResourceGroup -EA SilentlyContinue
    if (-not $sa -or -not $Script:Cfg.AppAttach.ShareName) {
        fx-fail "App Attach share not configured — run Setup first" ""; Pause-Screen; return
    }
    $ctx       = $sa.Context
    $shareName = $Script:Cfg.AppAttach.ShareName

    Write-Host ""
    fx-info "Supported package formats" ".msix, .msixbundle, .vhd, .vhdx, .cim"
    fx-info "Recommended format" ".cim (Composite Image Format) — fastest mount time in AVD"
    Write-Host ""

    $localPath = Read-Inp "Local path to package file (.msix/.vhd/.cim)"
    if (-not (Test-Path $localPath)) { fx-fail "File not found" $localPath; Pause-Screen; return }

    $fileName   = Split-Path $localPath -Leaf
    $appName    = Read-Inp "Application friendly name" -D ($fileName -replace '\.[^.]*$','')
    $appVersion = Read-Inp "Application version" -D "1.0.0.0"
    $appPublisher = Read-Inp "Publisher" -D "BDF Internal"

    # Create subfolder by app name
    $folderPath = "$appName"
    fx-step "Creating folder on share: $folderPath"
    New-AzStorageDirectory -ShareName $shareName -Path $folderPath -Context $ctx -EA SilentlyContinue | Out-Null

    fx-step "Uploading $fileName to Azure Files..."
    $fileSize = (Get-Item $localPath).Length / 1MB
    fx-info "File size" "$([Math]::Round($fileSize, 1)) MB"

    Set-AzStorageFileContent -ShareName $shareName -Source $localPath `
        -Path "$folderPath/$fileName" -Context $ctx -Force | Out-Null

    $fqdn    = "$($Script:Cfg.Storage.Account.Name).file.core.windows.net"
    $uncPath = "\\$fqdn\$shareName\$folderPath\$fileName"
    fx-ok "Package uploaded" $uncPath

    # Store package metadata
    $Script:Cfg.AppAttach.Packages += @{
        AppName    = $appName
        Version    = $appVersion
        Publisher  = $appPublisher
        FileName   = $fileName
        UNCPath    = $uncPath
        Registered = $false
    }
    Save-FXConfig

    if (Read-YN "Register this package with AVD now?" $true) {
        Register-AppAttachPackage -PkgUNC $uncPath -AppName $appName -Version $appVersion
    }
    Pause-Screen
}

function Register-AppAttachPackage {
    param([string]$PkgUNC="",[string]$AppName="",[string]$Version="")
    fx-sec "Register App Attach Package with AVD"

    if (-not $PkgUNC) {
        # List uploaded packages
        $pending = @($Script:Cfg.AppAttach.Packages | Where-Object { -not $_.Registered })
        if ($pending.Count -eq 0) {
            fx-warn "No unregistered packages found — upload a package first" ""; Pause-Screen; return
        }
        $i = 1
        foreach ($p in $pending) {
            Write-Host ("  {0}. {1,-30} v{2}" -f $i, $p.AppName, $p.Version) -ForegroundColor $FC.Menu; $i++
        }
        $ch = Read-Inp "Select package" (1..$pending.Count | ForEach-Object {"$_"})
        $sel = $pending[[int]$ch - 1]
        $PkgUNC  = $sel.UNCPath
        $AppName = $sel.AppName
        $Version = $sel.Version
    }

    fx-step "Registering package: $AppName v$Version"
    fx-info "Package path" $PkgUNC
    fx-info "NOTE: AVD App Attach (New-AzWvdAppAttachPackage) requires the package to be" ""
    fx-info "      accessible from the session host at registration time." ""

    # Select host pools to assign
    $hpNames = @($Script:Cfg.HostPools.E3.Name, $Script:Cfg.HostPools.F1.Name)
    $hpRg    = $Script:Cfg.RG.AVD.Name
    $i = 1
    Write-Host ""
    foreach ($hp in $hpNames) {
        Write-Host ("  {0}. {1}" -f $i, $hp) -ForegroundColor $FC.Menu; $i++
    }
    Write-Host "  A. All host pools" -ForegroundColor $FC.New
    $hpChoice = Read-Inp "Assign to host pool(s)" -V (@("A","a") + (1..$hpNames.Count | ForEach-Object {"$_"})) -D "A"
    $assignHPs = if ($hpChoice -in @("A","a")) { $hpNames } else { @($hpNames[[int]$hpChoice-1]) }

    foreach ($hp in $assignHPs) {
        $hpObj = Get-AzWvdHostPool -Name $hp -ResourceGroupName $hpRg -EA SilentlyContinue
        if (-not $hpObj) { fx-warn "Host pool not found" $hp; continue }

        # Create App Attach package
        $pkgName = "$($AppName -replace '[^a-zA-Z0-9]','_')_v$($Version -replace '\.','_')"
        $existing = Get-AzWvdAppAttachPackage -ResourceGroupName $hpRg -Name $pkgName -EA SilentlyContinue
        if ($existing) {
            fx-ok "Package already registered" $pkgName
        } else {
            try {
                New-AzWvdAppAttachPackage `
                    -Name $pkgName `
                    -ResourceGroupName $hpRg `
                    -Location $Script:Cfg.Location `
                    -ImagePath $PkgUNC `
                    -HostPoolReference @(@{ HostPoolArmPath=$hpObj.Id }) `
                    -IsActive $true `
                    -IsPackageTimestamped Timestamped `
                    -FailHealthCheckOnStagingFailure NeedsAssistance `
                    -EA Stop | Out-Null
                fx-ok "Package registered" "$pkgName → $hp"

                # Update stored metadata
                $pkg = $Script:Cfg.AppAttach.Packages | Where-Object { $_.UNCPath -eq $PkgUNC }
                if ($pkg) { $pkg.Registered = $true }
            } catch {
                fx-warn "Registration requires access to package from Azure — ensure private endpoint and permissions are in place" ""
                fx-warn "Error" $_.Exception.Message
                # Generate manual registration script instead
                @"
# Manual App Attach Registration Script
# Run from Azure Cloud Shell or a machine with Az PowerShell

`$pkgName  = '$pkgName'
`$imgPath  = '$PkgUNC'
`$hpId     = '$(($hpObj).Id)'
`$rg       = '$hpRg'
`$location = '$($Script:Cfg.Location)'

New-AzWvdAppAttachPackage ``
    -Name `$pkgName ``
    -ResourceGroupName `$rg ``
    -Location `$location ``
    -ImagePath `$imgPath ``
    -HostPoolReference @(@{HostPoolArmPath=`$hpId}) ``
    -IsActive `$true ``
    -IsPackageTimestamped Timestamped
"@ | Set-Content ".\Register-$pkgName.ps1" -Encoding UTF8
                fx-ok "Registration script saved" ".\Register-$pkgName.ps1"
            }
        }
    }
    Save-FXConfig
    Pause-Screen
}

function Manage-AppAttachPackages {
    fx-sec "App Attach Package Management"
    $pkgs = @(Get-AzWvdAppAttachPackage -ResourceGroupName $Script:Cfg.RG.AVD.Name -EA SilentlyContinue)
    if ($pkgs.Count -eq 0) {
        fx-warn "No App Attach packages found in resource group" $Script:Cfg.RG.AVD.Name
        Pause-Screen; return
    }

    $i = 1
    foreach ($p in $pkgs) {
        $active = if ($p.IsActive) { "✔ Active" } else { "⊘ Inactive" }
        $col    = if ($p.IsActive) { $FC.Ok } else { $FC.Muted }
        $name   = ($p.Name -split "/")[-1]
        Write-Host ("  {0,3}. {1,-40} {2}" -f $i, $name, $active) -ForegroundColor $col; $i++
    }

    Write-Host ""
    Write-Host "  Actions:  A=Activate  D=Deactivate  R=Remove  V=View Details  Q=Back" -ForegroundColor $FC.Muted
    $ch    = Read-Inp "Select # then action (e.g. '2 A')" -D "Q"
    if ($ch -in @("Q","q","0")) { return }

    if ($ch -match "^(\d+)\s+([ADRVadrv])$") {
        $idx    = [int]$Matches[1] - 1
        $action = $Matches[2].ToUpper()
        $pkg    = $pkgs[$idx]
        $pkgName = ($pkg.Name -split "/")[-1]

        switch ($action) {
            "A" {
                Update-AzWvdAppAttachPackage -Name $pkgName -ResourceGroupName $Script:Cfg.RG.AVD.Name -IsActive $true | Out-Null
                fx-ok "Activated" $pkgName
            }
            "D" {
                Update-AzWvdAppAttachPackage -Name $pkgName -ResourceGroupName $Script:Cfg.RG.AVD.Name -IsActive $false | Out-Null
                fx-ok "Deactivated" $pkgName
            }
            "R" {
                if (Read-YN "Remove package: $pkgName?" $false) {
                    Remove-AzWvdAppAttachPackage -Name $pkgName -ResourceGroupName $Script:Cfg.RG.AVD.Name -Force | Out-Null
                    fx-ok "Removed" $pkgName
                }
            }
            "V" {
                Write-Host ""
                $pkg | Format-List | Out-String | Write-Host -ForegroundColor $FC.Muted
            }
        }
    }
    Pause-Screen
}

function Assign-AppAttachToAppGroup {
    fx-sec "Assign App Attach Applications to Application Groups"
    fx-info "After registering a package, published apps can be added to RemoteApp groups"
    Write-Host ""
    $pkgs = @(Get-AzWvdAppAttachPackage -ResourceGroupName $Script:Cfg.RG.AVD.Name -EA SilentlyContinue)
    if ($pkgs.Count -eq 0) { fx-warn "No packages registered" ""; Pause-Screen; return }

    $i = 1
    foreach ($p in $pkgs) { Write-Host ("  {0}. {1}" -f $i, ($p.Name -split "/")[-1]) -ForegroundColor $FC.Menu; $i++ }
    $pkgChoice = Read-Inp "Select package" (1..$pkgs.Count | ForEach-Object {"$_"})
    $pkg       = $pkgs[[int]$pkgChoice - 1]
    $pkgName   = ($pkg.Name -split "/")[-1]

    # Get app groups in the resource group
    $ags = @(Get-AzWvdApplicationGroup -ResourceGroupName $Script:Cfg.RG.AVD.Name -EA SilentlyContinue)
    $i = 1
    foreach ($ag in $ags) {
        Write-Host ("  {0}. {1,-40} Type: {2}" -f $i, ($ag.Name -split "/")[-1], $ag.ApplicationGroupType) -ForegroundColor $FC.Menu; $i++
    }
    $agChoice  = Read-Inp "Select Application Group" (1..$ags.Count | ForEach-Object {"$_"})
    $appGroup  = $ags[[int]$agChoice - 1]
    $agName    = ($appGroup.Name -split "/")[-1]

    $appName   = Read-Inp "App friendly name (as shown to users)"
    $appId     = Read-Inp "App ID from MSIX manifest (leave blank for default)"

    fx-step "Adding MSIX app to Application Group: $agName"
    $appParams = @{
        Name                 = ($appName -replace '[^a-zA-Z0-9]','_')
        ApplicationGroupName = $agName
        ResourceGroupName    = $Script:Cfg.RG.AVD.Name
        FriendlyName         = $appName
        ApplicationType      = MsixApplication
        MsixPackageFamilyName = $pkgName
        ShowInPortal         = $true
    }
    if ($appId) { $appParams.MsixPackageApplicationId = $appId }
    try {
        New-AzWvdApplication @appParams | Out-Null
        fx-ok "App Attach app added to group" "$appName → $agName"
    } catch {
        fx-warn "Failed to add app — package may need staging" $_.Exception.Message
    }
    Pause-Screen
}

function Show-MSIXPackagingGuide {
    fx-hdr "MSIX PACKAGING GUIDE — CONVERT APPS TO APP ATTACH FORMAT"
    Write-Host @"
  $($FX.STL)$($FX.SH*68)$($FX.STR)
  $($FX.SV)  OVERVIEW                                                              $($FX.SV)
  $($FX.SBL)$($FX.SH*68)$($FX.SBR)

  App Attach supports these package formats:
  • .msix          — Standard MSIX package (preferred for signed apps)
  • .msixbundle    — Multi-arch MSIX bundle
  • .vhd / .vhdx  — Virtual disk containing MSIX packages (supports multiple apps)
  • .cim           — Composite Image Files (FASTEST mount — recommended for AVD)

  $($FX.STL)$($FX.SH*68)$($FX.STR)
  $($FX.SV)  STEP 1: PACKAGE YOUR APP AS MSIX                                     $($FX.SV)
  $($FX.SBL)$($FX.SH*68)$($FX.SBR)

  Tools needed (free from Microsoft):
  A. MSIX Packaging Tool (Microsoft Store or MSIX Packaging Tool Preview)
     → https://aka.ms/MSIXPackagingTool

  B. PSF (Package Support Framework) — for apps needing compatibility fixes
     → https://github.com/microsoft/MSIX-PackageSupportFramework

  Packaging process:
  1. Install MSIX Packaging Tool on a CLEAN reference machine
     (Windows 10/11, same OS version as your session hosts)
  2. Run MSIX Packaging Tool → Create package from installer
  3. Select your installer (.exe or .msi)
  4. Monitor install — capture all registry/file changes
  5. Set Package Name, Publisher, Version in wizard
  6. Test package on the same machine
  7. Export .msix file

  $($FX.STL)$($FX.SH*68)$($FX.STR)
  $($FX.SV)  STEP 2: SIGN THE MSIX PACKAGE                                        $($FX.SV)
  $($FX.SBL)$($FX.SH*68)$($FX.SBR)

  MSIX packages MUST be signed with a trusted certificate. Options:

  A. Internal CA (recommended for enterprise):
     # Create self-signed cert for testing:
     `$cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=BDF Internal" ``
         -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddYears(3)
     `$pwd = ConvertTo-SecureString -String "YourPassword" -Force -AsPlainText
     Export-PfxCertificate -Cert `$cert -FilePath ".\BDF-Sign.pfx" -Password `$pwd
     
     # Sign the package:
     SignTool.exe sign /fd SHA256 /p7ce DetachedSignedData /p7co 1.2.840.113549.1.7.2 ``
         /f ".\BDF-Sign.pfx" /p "YourPassword" ".\YourApp_1.0.0.0_x64.msix"

  B. Azure Trusted Signing (cloud-based, no cert management):
     → Use Azure Trusted Signing service (formerly Azure Code Signing)
     → Integrates with DevOps pipelines

  NOTE: Session hosts must trust the signing certificate (deploy via GPO/Intune)

  $($FX.STL)$($FX.SH*68)$($FX.STR)
  $($FX.SV)  STEP 3: CONVERT TO .CIM (FASTEST for AVD App Attach)                 $($FX.SV)
  $($FX.SBL)$($FX.SH*68)$($FX.SBR)

  CIM format mounts significantly faster than VHD (no full virtual disk overhead):

  # Download MSIXMGR tool:
  # https://aka.ms/msixmgr

  # Convert .msix to .cim:
  msixmgr.exe -Unpack -packagePath ".\YourApp_1.0.0.0_x64.msix" ``
      -destination ".\YourApp_1.0.0.0_x64.cim" -applyacls -create -fileType cim

  # Alternatively, convert to VHD:
  msixmgr.exe -Unpack -packagePath ".\YourApp_1.0.0.0_x64.msix" ``
      -destination ".\YourApp_1.0.0.0_x64.vhd" -applyacls -create ``
      -fileType vhd -rootDirectory apps

  $($FX.STL)$($FX.SH*68)$($FX.STR)
  $($FX.SV)  BDF-SPECIFIC PACKAGING NOTES                                          $($FX.SV)
  $($FX.SBL)$($FX.SH*68)$($FX.SBR)

  SAP GUI (if needed alongside browser):
    • Package SAP GUI for Windows as MSIX — capture all registry entries
    • SAP GUI is complex — use PSF runtime fix for COM registration issues
    • Consider: SAP Browser (RemoteApp) is simpler — avoid SAP GUI in App Attach
      unless explicitly required

  Zscaler Client Connector:
    • DO NOT package as App Attach — install in golden image instead
    • ZCC requires machine-level services and network filters — incompatible with MSIX

  Microsoft 365 Apps:
    • DO NOT package as App Attach — use M365 Apps Machine-Wide Installer on image
    • M365 Apps are explicitly excluded from MSIX app packaging by Microsoft

  Retail/POS Applications:
    • Excellent candidates for App Attach — update without re-imaging session hosts
    • Test: POS applications that use standard Win32/COM registration

"@ -ForegroundColor $FC.Muted
    Pause-Screen
}

function Monitor-AppAttachHealth {
    fx-sec "App Attach Health Monitoring"
    $pkgs = @(Get-AzWvdAppAttachPackage -ResourceGroupName $Script:Cfg.RG.AVD.Name -EA SilentlyContinue)
    if ($pkgs.Count -eq 0) { fx-warn "No packages registered" ""; Pause-Screen; return }

    Write-Host ""
    Write-Host ("  {0,-42} {1,-10} {2,-12} {3}" -f "Package", "Active", "Health", "Host Pools") -ForegroundColor $FC.Muted
    fx-line

    foreach ($p in $pkgs) {
        $name   = ($p.Name -split "/")[-1]
        $active = if ($p.IsActive) { "✔ Active" } else { "⊘ Inactive" }
        $health = $p.ProvisioningState ?? "Unknown"
        $hps    = if ($p.HostPoolReference) { $p.HostPoolReference.Count } else { 0 }
        $col    = if ($p.IsActive -and $health -eq "Succeeded") {$FC.Ok} elseif (-not $p.IsActive) {$FC.Muted} else {$FC.Warn}
        Write-Host ("  {0,-42} {1,-10} {2,-12} {3} pools" -f $name, $active, $health, $hps) -ForegroundColor $col
    }

    # KQL queries for Log Analytics
    if ($Script:Cfg.LogAnalytics.Id) {
        Write-Host ""
        fx-info "Useful KQL Queries for App Attach monitoring (run in Log Analytics):" ""
        Write-Host @"

  -- App Attach staging failures:
  WVDAppAttachPackageActivities
  | where TimeGenerated > ago(1h)
  | where ActivityType == "Stage"
  | where Result != "Success"
  | project TimeGenerated, HostPoolName, PackageFamilyName, ActivityType, Result, Error

  -- App launch events:
  WVDAppAttachPackageActivities  
  | where TimeGenerated > ago(24h)
  | summarize Launches=count() by PackageFamilyName
  | order by Launches desc

"@ -ForegroundColor $FC.Azure
    }
    Pause-Screen
}

function Troubleshoot-AppAttach {
    fx-hdr "APP ATTACH TROUBLESHOOTING GUIDE"
    $issues = @(
        @{
            Issue = "Package fails to stage (error in WVDAppAttachPackageActivities)"
            Causes= @("Session host cannot reach Azure Files share","Certificate not trusted on session host","Insufficient NTFS permissions on share","Package file corrupted or not signed")
            Fixes = @(
                "Verify: Test-NetConnection -ComputerName storageaccount.file.core.windows.net -Port 445"
                "Check: Event Viewer > Microsoft-Windows-AppXDeployment-Server > Operational"
                "Ensure VM managed identity has 'Storage File Data SMB Share Reader' role"
                "Verify package certificate is in session host Trusted Root store"
                "Re-sign package and re-upload: SignTool.exe verify /pa YourApp.cim"
            )
        },
        @{
            Issue = "App not visible in RD Client after App Attach registration"
            Causes= @("User not in assigned app group","Package not active","App Group not associated with Workspace")
            Fixes = @(
                "Verify: user is in Entra ID group assigned to the App Group"
                "Check package IsActive = true in AVD portal"
                "Ensure App Group is linked to the Workspace"
                "Re-publish: New-AzWvdApplication with MsixPackageFamilyName set correctly"
            )
        },
        @{
            Issue = "App launches but fails to run (crashes immediately)"
            Causes= @("Missing dependencies (.NET, VC++ redistributables)","COM registration required outside MSIX container","PSF runtime fix needed","32-bit app on 64-bit OS packaging issue")
            Fixes = @(
                "Add dependencies: include .NET or VC++ runtime in MSIX package"
                "Use PSF (Package Support Framework) for COM/registry compatibility"
                "Check package manifest: AppxManifest.xml > Applications > Application > Executable"
                "Test package locally: Add-AppxPackage -Path .\YourApp.msix"
            )
        },
        @{
            Issue = "FSLogix profile conflicts with App Attach (app data lost between sessions)"
            Causes= @("App writes to %APPDATA% which is in FSLogix profile VHDX","Profile not persisting correctly","ODFC and profile VHDX conflict")
            Fixes = @(
                "App Attach apps write to AppData within the MSIX container — this IS persisted"
                "Verify FSLogix profile is mounting correctly (Event ID 27 in FSLogix log)"
                "Check: FSLogix Redirections.xml is NOT excluding the app's AppData path"
            )
        }
    )

    foreach ($issue in $issues) {
        fx-sec $issue.Issue
        Write-Host "  Possible causes:" -ForegroundColor $FC.Warn
        foreach ($c in $issue.Causes) { Write-Host "    • $c" -ForegroundColor $FC.Muted }
        Write-Host ""
        Write-Host "  Resolution steps:" -ForegroundColor $FC.Ok
        foreach ($f in $issue.Fixes)  { Write-Host "    $($FX.Arrow) $f" -ForegroundColor $FC.Menu }
        Write-Host ""
    }
    Pause-Screen
}

function Export-SessionHostReadinessScript {
    fx-sec "Session Host App Attach Readiness Script"
    $script = @"
<#
.SYNOPSIS  Validate session host readiness for FSLogix + App Attach
.NOTES     Run on each session host or via Invoke-AzVMRunCommand during image build
           Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
#>

`$StorageFQDN    = '$($Script:Cfg.Storage.Account.Name).file.core.windows.net'
`$ProfileShare   = '$($Script:Cfg.Storage.Shares.E3)'
`$AppAttachShare = '$($Script:Cfg.AppAttach.ShareName)'
`$Results        = @{}

Write-Host "" ; Write-Host "  BDF AVD Session Host Readiness Check" -ForegroundColor Cyan

# 1. FSLogix installed
`$fxPath = 'C:\Program Files\FSLogix\Apps\frx.exe'
`$Results['FSLogix Installed'] = Test-Path `$fxPath
if (`$Results['FSLogix Installed']) {
    `$fxVer = (Get-Item `$fxPath).VersionInfo.FileVersion
    Write-Host "  [PASS] FSLogix installed: v`$fxVer" -ForegroundColor Green
} else {
    Write-Host "  [FAIL] FSLogix NOT installed" -ForegroundColor Red
}

# 2. FSLogix enabled in registry
`$fxEnabled = (Get-ItemProperty 'HKLM:\SOFTWARE\FSLogix\Profiles' -Name Enabled -EA SilentlyContinue).Enabled
`$Results['FSLogix Enabled'] = (`$fxEnabled -eq 1)
Write-Host "  [$(if(`$fxEnabled -eq 1){'PASS'}else{'FAIL'})] FSLogix registry Enabled = `$fxEnabled" -ForegroundColor (if(`$fxEnabled -eq 1){'Green'}else{'Red'})

# 3. Azure Files reachability (port 445)
`$smb = Test-NetConnection -ComputerName `$StorageFQDN -Port 445 -WarningAction SilentlyContinue
`$Results['Azure Files Reachable'] = `$smb.TcpTestSucceeded
Write-Host "  [$(if(`$smb.TcpTestSucceeded){'PASS'}else{'FAIL'})] Azure Files SMB (445): `$StorageFQDN" -ForegroundColor (if(`$smb.TcpTestSucceeded){'Green'}else{'Red'})

# 4. AVD Agent running
`$avdAgent = Get-Service -Name 'RDAgentBootLoader' -EA SilentlyContinue
`$Results['AVD Agent Running'] = (`$avdAgent.Status -eq 'Running')
Write-Host "  [$(if(`$avdAgent.Status -eq 'Running'){'PASS'}else{'FAIL'})] AVD Agent: `$(`$avdAgent.Status)" -ForegroundColor (if(`$avdAgent.Status -eq 'Running'){'Green'}else{'Red'})

# 5. Hyper-V FS Filter driver (App Attach dependency)
`$hvFilter = Get-WindowsOptionalFeature -Online -FeatureName Containers-DisposableClientVM -EA SilentlyContinue
`$Results['Hyper-V FS Filter'] = (`$hvFilter.State -eq 'Enabled')
Write-Host "  [$(if(`$hvFilter.State -eq 'Enabled'){'PASS'}else{'WARN'})] Hyper-V FS Filter (Containers feature): `$(`$hvFilter.State)" -ForegroundColor (if(`$hvFilter.State -eq 'Enabled'){'Green'}else{'Yellow'})

# 6. Certificate store check (App Attach signing cert)
`$certs = Get-ChildItem Cert:\LocalMachine\Root | Where-Object {`$_.Subject -like '*BDF*' -or `$_.Subject -like '*Microsoft*'}
`$Results['Signing Certs'] = (`$certs.Count -gt 0)
Write-Host "  [INFO] Trusted Root certs (BDF/Microsoft): `$(`$certs.Count) found" -ForegroundColor Cyan

# 7. Required Windows features for MSIX App Attach
`$msixFeature = Get-WindowsOptionalFeature -Online -FeatureName Containers-SharedPackageContainer -EA SilentlyContinue
`$Results['MSIX Container Feature'] = (`$msixFeature.State -eq 'Enabled')
Write-Host "  [$(if(`$msixFeature.State -eq 'Enabled'){'PASS'}else{'WARN'})] MSIX Shared Package Container: `$(`$msixFeature.State)" -ForegroundColor (if(`$msixFeature.State -eq 'Enabled'){'Green'}else{'Yellow'})

# 8. Profile VHD location configured
`$vhdLoc = (Get-ItemProperty 'HKLM:\SOFTWARE\FSLogix\Profiles' -Name VHDLocations -EA SilentlyContinue).VHDLocations
`$Results['VHD Location Set'] = (`$null -ne `$vhdLoc -and `$vhdLoc -ne '')
Write-Host "  [$(if(`$vhdLoc){'PASS'}else{'FAIL'})] FSLogix VHDLocations: $(if(`$vhdLoc){`$vhdLoc}else{'NOT SET'})" -ForegroundColor (if(`$vhdLoc){'Green'}else{'Red'})

# Summary
`$passed = @(`$Results.Values | Where-Object {`$_}).Count
`$total  = `$Results.Count
Write-Host ""
Write-Host "  Result: `$passed / `$total checks passed" -ForegroundColor (if(`$passed -eq `$total){'Green'}else{'Yellow'})
"@
    $script | Set-Content ".\Test-AVDSessionHostReadiness.ps1" -Encoding UTF8
    fx-ok "Readiness script generated" ".\Test-AVDSessionHostReadiness.ps1"
    fx-info "Deploy via" "Invoke-AzVMRunCommand or Intune Win32 App (detection-only mode)"
    Pause-Screen
}
#endregion

#region ════════════════════════════════════════════════════════════════════
# DIAGNOSTICS & MONITORING
#════════════════════════════════════════════════════════════════════════════

function Show-DiagnosticsMenu {
    fx-hdr "FSLogix & APP ATTACH DIAGNOSTICS" $FC.Warn

    while ($true) {
        fx-sec "Diagnostics Menu"
        Write-Host "  1.  Profile health report (VHDX sizes, last mount, errors)" -ForegroundColor $FC.Menu
        Write-Host "  2.  Generate KQL queries for Log Analytics (all scenarios)" -ForegroundColor $FC.Azure
        Write-Host "  3.  Parse FSLogix logs on session host" -ForegroundColor $FC.Menu
        Write-Host "  4.  Profile size analysis and cleanup recommendations" -ForegroundColor $FC.Menu
        Write-Host "  5.  Export full diagnostic report to HTML" -ForegroundColor $FC.New
        Write-Host "  6.  Common FSLogix fixes (automated remediation)" -ForegroundColor $FC.Warn
        Write-Host "  0.  Back" -ForegroundColor $FC.Muted
        $ch = Read-Inp "Choice" -V @("0","1","2","3","4","5","6")
        switch ($ch) {
            "1" { Show-ProfileHealthReport }
            "2" { Export-KQLQueries }
            "3" { Parse-FSLogixLogs }
            "4" { Analyze-ProfileSizes }
            "5" { Export-DiagnosticHTML }
            "6" { Invoke-CommonFixes }
            "0" { return }
        }
    }
}

function Export-KQLQueries {
    fx-sec "KQL Queries for FSLogix + App Attach Monitoring"
    $kqlFile = ".\BDF-AVD-FSLogix-KQL-Queries.kql"
    $queries = @"
// ═══════════════════════════════════════════════════════════════════════
// BDF AVD — FSLogix & App Attach KQL Query Library
// Workspace: $($Script:Cfg.LogAnalytics.Name)
// Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
// ═══════════════════════════════════════════════════════════════════════

// ── FSLOGIX PROFILE EVENTS ───────────────────────────────────────────────

// [1] Profile mount success rate (last 24h)
Event
| where Source == "Microsoft-FSLogix-Apps"
| where TimeGenerated > ago(24h)
| extend ProfileResult = case(
    EventID == 27, "Success",
    EventID == 33, "Failure - Using Local Profile",
    EventID == 34, "Failure - Profile Locked",
    EventID == 52, "VHDX Full",
    EventID == 35, "VHDX Not Found",
    "Other")
| summarize Count=count() by ProfileResult, Computer
| order by Count desc

// [2] FSLogix profile mount failures (last 7 days)
Event
| where Source == "Microsoft-FSLogix-Apps"
| where EventID in (33, 34, 35, 52, 60, 62)
| where TimeGenerated > ago(7d)
| project TimeGenerated, Computer, EventID, RenderedDescription
| order by TimeGenerated desc

// [3] Average profile mount time
Event
| where Source == "Microsoft-FSLogix-Apps"
| where EventID == 27
| where TimeGenerated > ago(24h)
| extend MountTime = extract("mount time: (\\d+) ms", 1, RenderedDescription)
| where isnotempty(MountTime)
| summarize AvgMountMs=avg(todouble(MountTime)), P95MountMs=percentile(todouble(MountTime),95) by Computer

// [4] Profile VHDX capacity warnings (EventID 52 = VHDX > 90% full)
Event
| where Source == "Microsoft-FSLogix-Apps"
| where EventID == 52
| where TimeGenerated > ago(30d)
| project TimeGenerated, Computer, RenderedDescription
| order by TimeGenerated desc

// [5] Users falling back to local profile (Event 33 — DATA LOSS RISK)
Event
| where Source == "Microsoft-FSLogix-Apps"
| where EventID == 33
| where TimeGenerated > ago(7d)
| extend UserName = extract("for user ([^,]+)", 1, RenderedDescription)
| project TimeGenerated, Computer, UserName, RenderedDescription

// ── AVD SESSION HOST HEALTH ──────────────────────────────────────────────

// [6] Session host availability timeline
WVDAgentHealthStatus
| where TimeGenerated > ago(24h)
| summarize AvailableHosts=countif(Status == "Available"),
            UnavailableHosts=countif(Status == "Unavailable"),
            Total=count()
            by HostPoolName, bin(TimeGenerated, 15m)
| render timechart

// [7] AVD connection latency (RDP RTT)
WVDConnections
| where TimeGenerated > ago(24h)
| summarize AvgRoundTripMs=avg(tolong(EstRoundTripTimeInMs)), Sessions=count() by HostPool=_ResourceId
| order by AvgRoundTripMs desc

// [8] Logon time analysis (Phase breakdown)
WVDConnections
| where TimeGenerated > ago(24h)
| where isnotempty(ConnectTime) and isnotempty(StartTime)
| extend TotalLogonSec=datetime_diff('second', ConnectTime, StartTime)
| where TotalLogonSec > 0
| summarize AvgLogon=avg(TotalLogonSec), P95Logon=percentile(TotalLogonSec,95),
            SlowLogins=countif(TotalLogonSec > 60)
            by HostPool=split(_ResourceId,"/")[-1]

// ── APP ATTACH ────────────────────────────────────────────────────────────

// [9] App Attach staging activity
WVDAppAttachPackageActivities
| where TimeGenerated > ago(24h)
| summarize count() by ActivityType, Result, PackageFamilyName
| order by count_ desc

// [10] App Attach failures detail
WVDAppAttachPackageActivities
| where TimeGenerated > ago(24h)
| where Result != "Success"
| project TimeGenerated, HostPoolName, PackageFamilyName, ActivityType, Result, Error
| order by TimeGenerated desc

// [11] Most used App Attach applications
WVDAppAttachPackageActivities
| where TimeGenerated > ago(7d)
| where ActivityType == "Register"
| where Result == "Success"
| summarize Registrations=count() by PackageFamilyName
| order by Registrations desc

// ── CAPACITY & SCALING ────────────────────────────────────────────────────

// [12] Session host CPU pressure (P95 CPU > 80%)
Perf
| where TimeGenerated > ago(1h)
| where ObjectName == "Processor" and CounterName == "% Processor Time" and InstanceName == "_Total"
| summarize P95CPU=percentile(CounterValue, 95) by Computer
| where P95CPU > 80
| order by P95CPU desc

// [13] Session host memory pressure
Perf
| where TimeGenerated > ago(1h)
| where ObjectName == "Memory" and CounterName == "Available MBytes"
| summarize AvailMB=avg(CounterValue) by Computer
| extend AvailGB=round(AvailMB/1024, 1)
| where AvailGB < 2
| order by AvailGB asc

// [14] Azure Files latency (SMB operations)
StorageFileLogs
| where TimeGenerated > ago(1h)
| where AccountName == "$($Script:Cfg.Storage.Account.Name)"
| summarize AvgLatencyMs=avg(DurationMs), P99LatencyMs=percentile(DurationMs, 99) by OperationName
| order by AvgLatencyMs desc
"@
    $queries | Set-Content $kqlFile -Encoding UTF8
    fx-ok "KQL queries exported" $kqlFile
    fx-info "Import into" "Azure Portal → Log Analytics → Logs → Import query"
    Pause-Screen
}

function Show-ProfileHealthReport {
    fx-sec "Profile Health Check (via Azure Files)"
    $saName = $Script:Cfg.Storage.Account.Name
    $saRG   = $Script:Cfg.Storage.Account.ResourceGroup
    $sa = Get-AzStorageAccount -Name $saName -ResourceGroupName $saRG -EA SilentlyContinue
    if (-not $sa) { fx-warn "Storage account not found" $saName; Pause-Screen; return }
    $ctx = $sa.Context

    foreach ($shareName in @($Script:Cfg.Storage.Shares.E3, $Script:Cfg.Storage.Shares.F1, $Script:Cfg.Storage.Shares.ODFC)) {
        fx-sec "Share: $shareName"
        $share = Get-AzStorageShare -Name $shareName -Context $ctx -EA SilentlyContinue
        if (-not $share) { fx-warn "Share not found" $shareName; continue }

        $usage = $share.Snapshot ?? 0
        $quota = $share.ShareProperties.Quota
        $files = @(Get-AzStorageFile -ShareName $shareName -Context $ctx -EA SilentlyContinue)
        $vhdxFiles = @($files | Where-Object { $_.Name -like "*.vhdx" })
        $totalSizeGB = [Math]::Round(($vhdxFiles | Measure-Object -Property Length -Sum).Sum / 1GB, 2)

        fx-info "Quota"          "$quota GB"
        fx-info "VHDX files"     "$($vhdxFiles.Count)"
        fx-info "Total VHDX size" "$totalSizeGB GB"

        if ($vhdxFiles.Count -gt 0) {
            Write-Host ""
            Write-Host ("  {0,-40} {1,-12} {2}" -f "Filename", "Size (MB)", "Last Modified") -ForegroundColor $FC.Muted
            fx-line
            foreach ($vhd in ($vhdxFiles | Sort-Object LastModified -Descending | Select-Object -First 10)) {
                $sizeMB = [Math]::Round($vhd.Properties.Length / 1MB, 1)
                $col = if ($sizeMB -gt 9000) { $FC.Warn } elseif ($sizeMB -gt 7500) { $FC.Gold } else { $FC.Menu }
                Write-Host ("  {0,-40} {1,-12} {2}" -f $vhd.Name, $sizeMB, $vhd.Properties.LastModified) -ForegroundColor $col
            }
        }
    }
    Pause-Screen
}

function Parse-FSLogixLogs {
    fx-sec "FSLogix Log Parser (Remote)"
    $hp = Read-Inp "Host pool name" -D $Script:Cfg.HostPools.E3.Name
    $rg = Read-Inp "Resource group" -D $Script:Cfg.RG.AVD.Name

    $parseScript = @'
$logDir = "C:\ProgramData\FSLogix\Logs"
$logFiles = Get-ChildItem $logDir -Filter "Profile_*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 5
$errors = @(); $warnings = @(); $successes = 0

foreach ($lf in $logFiles) {
    $content = Get-Content $lf.FullName -EA SilentlyContinue
    $errors   += @($content | Select-String -Pattern "ERROR|FAIL" | Select-Object -Last 5)
    $warnings += @($content | Select-String -Pattern "WARN"       | Select-Object -Last 3)
    $successes += ($content | Select-String "Profile loaded successfully" | Measure-Object).Count
}

Write-Host "FSLogix Log Summary (last 5 log files):"
Write-Host "  Successful mounts: $successes"
Write-Host "  Errors found: $($errors.Count)"
Write-Host "  Warnings: $($warnings.Count)"
if ($errors.Count -gt 0) { Write-Host "`nLast errors:"; $errors | ForEach-Object { Write-Host "  $_" -ForegroundColor Red } }
'@
    fx-step "Parsing FSLogix logs via VM Run Command..."
    $hosts = @(Get-AzWvdSessionHost -HostPoolName $hp -ResourceGroupName $rg -EA SilentlyContinue | Select-Object -First 3)
    foreach ($h in $hosts) {
        $vmName = ($h.Name -split "/")[-1]
        fx-step "Checking $vmName..."
        $result = Invoke-AzVMRunCommand -ResourceGroupName $rg -VMName $vmName `
                      -CommandId RunPowerShellScript -ScriptString $parseScript -EA SilentlyContinue
        if ($result) {
            Write-Host ($result.Value[0].Message) -ForegroundColor $FC.Muted
        }
    }
    Pause-Screen
}

function Analyze-ProfileSizes {
    fx-sec "Profile Size Analysis & Cleanup Recommendations"
    Write-Host @"
  Profile VHDX Growth Causes:
  
  $($FX.Arrow) Teams cache (if ODFC not configured correctly)    → Can reach 5-8 GB
  $($FX.Arrow) Browser cache (if not excluded)                   → Can reach 500MB-2GB
  $($FX.Arrow) Windows Search index                              → 200-500MB
  $($FX.Arrow) Crash dumps accumulation                          → Variable
  $($FX.Arrow) Downloaded files in profile                       → User-dependent
  
  Recommended Actions:
  
  1. Verify ODFC is configured and Teams cache is routing to ODFC (not main profile)
     Registry check: HKLM\SOFTWARE\Policies\FSLogix\ODFC\IncludeTeams = 1
  
  2. Confirm Redirections.xml is present in profile share root
     Check: \\storageaccount.file.core.windows.net\profiles-e3\Redirections.xml
  
  3. Enable OneDrive Known Folder Move to redirect Desktop/Documents/Pictures
     Intune: Device Config > Windows > OneDrive > Silently sign in + KFM
  
  4. Set profile disk compaction schedule (FSLogix 2201+):
     HKLM\SOFTWARE\FSLogix\Profiles\VHDCompactDisk = 1
     HKLM\SOFTWARE\FSLogix\Profiles\OnLogOffEnabled = 1
"@ -ForegroundColor $FC.Muted

    if (Read-YN "Generate profile cleanup runbook for Azure Automation?" $true) {
        @"
# FSLogix Profile Compaction Runbook
# Schedule: Sunday 2 AM — compress large VHDXs to reclaim Azure Files space
# Requires: Az.Storage module + Storage Contributor role

Connect-AzAccount -Identity | Out-Null
Set-AzContext -SubscriptionId (Get-AutomationVariable "AVD-SubscriptionId") | Out-Null

`$saName = '$($Script:Cfg.Storage.Account.Name)'
`$saRg   = '$($Script:Cfg.Storage.Account.ResourceGroup)'
`$sa     = Get-AzStorageAccount -Name `$saName -ResourceGroupName `$saRg
`$ctx    = `$sa.Context

foreach (`$share in @('$($Script:Cfg.Storage.Shares.E3)','$($Script:Cfg.Storage.Shares.F1)')) {
    `$files = @(Get-AzStorageFile -ShareName `$share -Context `$ctx -EA SilentlyContinue | Where-Object {`$_.Name -like "*.vhdx"})
    `$largeFiles = @(`$files | Where-Object { `$_.Properties.Length -gt 5GB })
    Write-Output "Share: `$share | Total VHDXs: `$(`$files.Count) | Large (>5GB): `$(`$largeFiles.Count)"
    # Note: VHDX compaction requires mounting the VHD on a Windows host — cannot be done purely via Azure Storage APIs
    # For automated compaction, consider: Add-AVDSessionHost runbook to mount + compact + unmount during off-hours
}
"@ | Set-Content ".\Invoke-ProfileCompaction.ps1" -Encoding UTF8
        fx-ok "Profile compaction runbook" ".\Invoke-ProfileCompaction.ps1"
    }
    Pause-Screen
}

function Export-DiagnosticHTML {
    fx-sec "Generating Diagnostic HTML Report"
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    $pkgs    = @(Get-AzWvdAppAttachPackage -ResourceGroupName $Script:Cfg.RG.AVD.Name -EA SilentlyContinue)
    $hpE3    = Get-AzWvdHostPool -Name $Script:Cfg.HostPools.E3.Name -ResourceGroupName $Script:Cfg.RG.AVD.Name -EA SilentlyContinue
    $hpF1    = Get-AzWvdHostPool -Name $Script:Cfg.HostPools.F1.Name -ResourceGroupName $Script:Cfg.RG.AVD.Name -EA SilentlyContinue
    $hostsE3 = @(Get-AzWvdSessionHost -HostPoolName $Script:Cfg.HostPools.E3.Name -ResourceGroupName $Script:Cfg.RG.AVD.Name -EA SilentlyContinue)
    $hostsF1 = @(Get-AzWvdSessionHost -HostPoolName $Script:Cfg.HostPools.F1.Name -ResourceGroupName $Script:Cfg.RG.AVD.Name -EA SilentlyContinue)

    $html = @"
<!DOCTYPE html>
<html><head><meta charset='UTF-8'>
<title>BDF AVD FSLogix & App Attach Report — $ts</title>
<style>
body{font-family:Arial,sans-serif;background:#0A1628;color:#E2E8F0;margin:0;padding:30px;}
h1{background:linear-gradient(135deg,#0078D4,#00B4D8);-webkit-background-clip:text;-webkit-text-fill-color:transparent;font-size:28px;}
h2{color:#50ABF1;border-bottom:1px solid #1E3A5F;padding-bottom:8px;margin-top:30px;}
.card{background:#112240;border:1px solid #1E3A5F;border-radius:10px;padding:20px;margin:15px 0;}
table{width:100%;border-collapse:collapse;font-size:13px;}
th{background:#0078D4;color:white;padding:10px 14px;text-align:left;}
td{padding:9px 14px;border-bottom:1px solid #1E3A5F;}
tr:hover td{background:#162D4E;}
.ok{color:#6EE7B7;font-weight:bold;} .fail{color:#FCA5A5;font-weight:bold;} .warn{color:#FCD34D;font-weight:bold;}
.badge{display:inline-block;padding:3px 10px;border-radius:12px;font-size:11px;font-weight:700;}
.badge-ok{background:rgba(16,185,129,0.2);color:#6EE7B7;} .badge-fail{background:rgba(239,68,68,0.2);color:#FCA5A5;}
</style></head><body>
<h1>BDF Azure Virtual Desktop<br>FSLogix &amp; App Attach Diagnostic Report</h1>
<p style='color:#64748B'>Generated: $ts | Environment: $($Script:Cfg.Environment) | Region: $($Script:Cfg.Location)</p>

<h2>FSLogix Configuration</h2>
<div class='card'>
<table><tr><th>Setting</th><th>E3 Office</th><th>F1 Frontline</th></tr>
<tr><td>Profile VHD Location</td><td>$($Script:Cfg.FSLogix.E3VHDLocation)</td><td>$($Script:Cfg.FSLogix.F1VHDLocation)</td></tr>
<tr><td>VHDX Max Size</td><td>$($Script:Cfg.FSLogix.E3SizeGB) GB</td><td>$($Script:Cfg.FSLogix.F1SizeGB) GB</td></tr>
<tr><td>ODFC Location</td><td colspan='2'>$($Script:Cfg.FSLogix.ODFCLocation)</td></tr>
<tr><td>ODFC Size</td><td colspan='2'>$($Script:Cfg.FSLogix.ODFCSizeGB) GB</td></tr>
<tr><td>Cloud Cache</td><td colspan='2'>$(if($Script:Cfg.FSLogix.CloudCache.Enabled){"Enabled — $($Script:Cfg.FSLogix.CloudCache.Locations.Count) locations"}else{"Disabled"})</td></tr>
</table></div>

<h2>Session Host Status</h2>
<div class='card'>
<table><tr><th>Pool</th><th>Host</th><th>Status</th><th>Sessions</th><th>Allow New</th></tr>
$(foreach ($h in ($hostsE3 + $hostsF1)) {
    $vn     = ($h.Name -split "/")[-1]
    $pool   = if ($h.Name -like "*e3*") {"E3"} else {"F1"}
    $stcol  = if ($h.Status -eq "Available") {"ok"} else {"fail"}
    $allow  = if ($h.AllowNewSession) {"Yes"} else {"No (Drain)"}
    "<tr><td>$pool</td><td>$vn</td><td><span class='$stcol'>$($h.Status)</span></td><td>$($h.Session)</td><td>$allow</td></tr>"
})
</table></div>

<h2>App Attach Packages ($($pkgs.Count))</h2>
<div class='card'>
$(if ($pkgs.Count -eq 0) {"<p class='warn'>No App Attach packages registered</p>"} else {
"<table><tr><th>Package</th><th>Active</th><th>Health</th></tr>
$(foreach ($p in $pkgs) {
    $pn = ($p.Name -split "/")[-1]
    $ab = if ($p.IsActive) {"<span class='badge badge-ok'>Active</span>"} else {"<span class='badge badge-fail'>Inactive</span>"}
    "<tr><td>$pn</td><td>$ab</td><td>$($p.ProvisioningState)</td></tr>"
})
</table>"
})
</div>

<h2>App Masking Rules</h2>
<div class='card'>
$(if (@($Script:Cfg.AppMasking.RuleFiles).Count -eq 0) {"<p class='warn'>No masking rules configured</p>"} else {
"<table><tr><th>Rule Set</th><th>Rules</th></tr>
$(foreach ($r in $Script:Cfg.AppMasking.RuleFiles) { "<tr><td>$($r.Name)</td><td>$($r.RuleCount)</td></tr>" })
</table>"
})
</div>
</body></html>
"@
    $file = ".\BDF-AVD-FSLogix-AppAttach-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
    $html | Set-Content $file -Encoding UTF8
    fx-ok "Diagnostic HTML report exported" $file
    Pause-Screen
}

function Invoke-CommonFixes {
    fx-sec "Common FSLogix Automated Fixes"
    Write-Host "  1.  Reset stuck/locked VHDX (clears .lock file from Azure Files)" -ForegroundColor $FC.Menu
    Write-Host "  2.  Re-enable session host drain mode (maintenance mode)" -ForegroundColor $FC.Menu
    Write-Host "  3.  Force FSLogix registry re-apply on all session hosts" -ForegroundColor $FC.Menu
    Write-Host "  4.  Clear FSLogix temp VHDXs (.tmp files on profile share)" -ForegroundColor $FC.Warn
    Write-Host "  0.  Back" -ForegroundColor $FC.Muted
    $ch = Read-Inp "Fix" -V @("0","1","2","3","4")
    switch ($ch) {
        "1" {
            fx-sec "Reset Locked VHDX Files"
            $saName = $Script:Cfg.Storage.Account.Name
            $saRg   = $Script:Cfg.Storage.Account.ResourceGroup
            $sa = Get-AzStorageAccount -Name $saName -ResourceGroupName $saRg -EA SilentlyContinue
            if ($sa) {
                $ctx = $sa.Context
                foreach ($share in @($Script:Cfg.Storage.Shares.E3, $Script:Cfg.Storage.Shares.F1)) {
                    $locks = @(Get-AzStorageFile -ShareName $share -Context $ctx -EA SilentlyContinue | Where-Object { $_.Name -like "*.vhdx.lock" })
                    foreach ($lock in $locks) {
                        fx-step "Removing lock file: $($lock.Name)"
                        Remove-AzStorageFile -ShareName $share -Path $lock.Name -Context $ctx -EA SilentlyContinue | Out-Null
                        fx-ok "Lock removed" $lock.Name
                    }
                    if ($locks.Count -eq 0) { fx-ok "No lock files found on" $share }
                }
            }
        }
        "3" {
            fx-sec "Force Registry Re-apply via VM Run Command"
            fx-step "Pushing FSLogix registry settings to all session hosts..."
            $regCmd = "Set-ItemProperty 'HKLM:\SOFTWARE\FSLogix\Profiles' -Name Enabled -Value 1 -Force; Restart-Service -Name 'frxsvc' -Force -EA SilentlyContinue"
            foreach ($pool in @("E3","F1")) {
                $hosts = @(Get-AzWvdSessionHost -HostPoolName $Script:Cfg.HostPools[$pool].Name -ResourceGroupName $Script:Cfg.RG.AVD.Name -EA SilentlyContinue)
                foreach ($h in $hosts) {
                    $vn = ($h.Name -split "/")[-1]
                    Invoke-AzVMRunCommand -ResourceGroupName $Script:Cfg.RG.AVD.Name -VMName $vn `
                        -CommandId RunPowerShellScript -ScriptString $regCmd -EA SilentlyContinue | Out-Null
                    fx-ok "Registry refreshed" $vn
                }
            }
        }
        "4" {
            fx-sec "Remove .tmp VHDXs from Profile Shares"
            $sa = Get-AzStorageAccount -Name $Script:Cfg.Storage.Account.Name -ResourceGroupName $Script:Cfg.Storage.Account.ResourceGroup -EA SilentlyContinue
            if ($sa) {
                $ctx = $sa.Context
                foreach ($share in @($Script:Cfg.Storage.Shares.E3, $Script:Cfg.Storage.Shares.F1)) {
                    $tmps = @(Get-AzStorageFile -ShareName $share -Context $ctx -EA SilentlyContinue | Where-Object { $_.Name -like "*.tmp" })
                    if ($tmps.Count -gt 0 -and (Read-YN "Delete $($tmps.Count) .tmp files from $share?" $false)) {
                        foreach ($t in $tmps) {
                            Remove-AzStorageFile -ShareName $share -Path $t.Name -Context $ctx -EA SilentlyContinue | Out-Null
                            fx-ok "Removed" $t.Name
                        }
                    } else { fx-ok "No .tmp files found" $share }
                }
            }
        }
    }
    Pause-Screen
}
#endregion

#region ════════════════════════════════════════════════════════════════════
# ONEDRIVE — SILENT SIGN-IN, KNOWN FOLDER MOVE, AVD MULTI-SESSION
#════════════════════════════════════════════════════════════════════════════

# OneDrive registry path constants
$ODReg = @{
    MachinePolicies  = 'HKLM:\SOFTWARE\Policies\Microsoft\OneDrive'
    MachineConfig    = 'HKLM:\SOFTWARE\Microsoft\OneDrive'
    UserPolicies     = 'HKCU:\SOFTWARE\Policies\Microsoft\OneDrive'
    UserConfig       = 'HKCU:\SOFTWARE\Microsoft\OneDrive'
    RunKey           = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
    StartupApproval  = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run'
}

# Full best-practice setting catalog
$ODSettings = [ordered]@{
    "AUTHENTICATION & SIGN-IN" = @(
        @{ Key="SilentAccountConfig";              Hive="HKLM"; Path="Policies"; Val=1;    Type="DWord";  Desc="Silent sign-in using Windows/Entra ID credentials — NO password prompt" }
        @{ Key="EnableADAL";                       Hive="HKLM"; Path="Config";   Val=1;    Type="DWord";  Desc="Force Modern Authentication (ADAL/MSAL) — required for MFA/Conditional Access" }
        @{ Key="AllowTenantList";                  Hive="HKLM"; Path="Policies"; Val="<TenantID>"; Type="String"; Desc="Whitelist only BDF corporate tenant — blocks other org accounts" }
        @{ Key="DisablePersonalSync";              Hive="HKLM"; Path="Policies"; Val=1;    Type="DWord";  Desc="Block users from signing in with personal Microsoft accounts" }
        @{ Key="PreventNetworkTrafficPreUserSignIn"; Hive="HKLM"; Path="Policies"; Val=1;  Type="DWord";  Desc="Stop OneDrive network traffic until user signs in (reduces session startup overhead)" }
    )
    "KNOWN FOLDER MOVE (KFM)" = @(
        @{ Key="KFMSilentOptIn";                   Hive="HKLM"; Path="Policies"; Val="<TenantID>"; Type="String"; Desc="SILENTLY redirect Desktop/Documents/Pictures to OneDrive — no user prompt" }
        @{ Key="KFMSilentOptInWithNotification";   Hive="HKLM"; Path="Policies"; Val=0;    Type="DWord";  Desc="0=Silent (no toast), 1=Show notification after KFM completes" }
        @{ Key="KFMBlockOptOut";                   Hive="HKLM"; Path="Policies"; Val=1;    Type="DWord";  Desc="Prevent users from moving folders BACK from OneDrive to local — enforces policy" }
        @{ Key="KFMBlockOptIn";                    Hive="HKLM"; Path="Policies"; Val=0;    Type="DWord";  Desc="0=Allow KFM (required for silent opt-in to work), 1=Block KFM entirely" }
        @{ Key="KFMOptInWithWizard";               Hive="HKLM"; Path="Policies"; Val=0;    Type="DWord";  Desc="0=No wizard (silent), 1=Show KFM wizard at next sign-in" }
    )
    "FILES ON DEMAND" = @(
        @{ Key="FilesOnDemandEnabled";             Hive="HKLM"; Path="Policies"; Val=1;    Type="DWord";  Desc="CRITICAL for AVD — files stay in cloud, download on access. Prevents disk fill on session hosts." }
        @{ Key="DehydrateSyncedTeamSites";         Hive="HKLM"; Path="Policies"; Val=1;    Type="DWord";  Desc="Files On Demand for SharePoint Team Sites — keeps session host disk usage minimal" }
        @{ Key="AutoMountTeamSites";               Hive="HKLM"; Path="Policies"; Val=1;    Type="DWord";  Desc="Auto-mount SharePoint sites the user has frequent access to" }
    )
    "PERFORMANCE & BANDWIDTH" = @(
        @{ Key="EnableAllOcsiClients";             Hive="HKLM"; Path="Policies"; Val=1;    Type="DWord";  Desc="Enable Office co-authoring (real-time collaboration on Office files in OneDrive)" }
        @{ Key="AutoADGroupGPO";                   Hive="HKLM"; Path="Policies"; Val=0;    Type="DWord";  Desc="0=Normal, 1=Auto-expand all shared folders (use 0 to reduce initial sync load)" }
        @{ Key="UploadBandwidthLimit";             Hive="HKLM"; Path="Policies"; Val=0;    Type="DWord";  Desc="Upload bandwidth limit KB/s (0=unlimited). Set to 1024 for shared AVD hosts if needed." }
        @{ Key="DownloadBandwidthLimit";           Hive="HKLM"; Path="Policies"; Val=0;    Type="DWord";  Desc="Download bandwidth limit KB/s (0=unlimited)" }
        @{ Key="WarningMinDiskSpaceMB";            Hive="HKLM"; Path="Policies"; Val=2048; Type="DWord";  Desc="Warn when free disk on session host < 2GB (important for shared AVD VMs)" }
    )
    "USER EXPERIENCE" = @(
        @{ Key="DisableFirstDeleteDialog";         Hive="HKLM"; Path="Policies"; Val=1;    Type="DWord";  Desc="Suppress 'files will be deleted from cloud' confirmation dialog" }
        @{ Key="DisableFRETutorial";               Hive="HKLM"; Path="Policies"; Val=1;    Type="DWord";  Desc="Skip first-run tutorial (cleaner first login for AVD users)" }
        @{ Key="DisableTutorial";                  Hive="HKLM"; Path="Policies"; Val=1;    Type="DWord";  Desc="Disable onboarding tutorial entirely" }
        @{ Key="EnableHoldTheFile";                Hive="HKLM"; Path="Policies"; Val=1;    Type="DWord";  Desc="Keep local copy during sync conflicts until resolved (prevents silent data loss)" }
        @{ Key="OpenAtLogin";                      Hive="HKLM"; Path="Config";   Val=1;    Type="DWord";  Desc="Start OneDrive at user login (per-machine install sets this automatically)" }
        @{ Key="Tier1Extensions";                  Hive="HKLM"; Path="Policies"; Val="";   Type="String"; Desc="File extensions to always keep local (e.g. .pst — leave blank for default)" }
    )
    "AVD MULTI-SESSION SPECIFIC" = @(
        @{ Key="PerMachineInstall";                Hive="N/A";  Path="Install";  Val=1;    Type="Setup"; Desc="CRITICAL: Install OneDrive with /allusers flag — shared host, per-user sync scope" }
        @{ Key="ShellIntegratorEnabled";           Hive="HKLM"; Path="Config";   Val=0;    Type="DWord";  Desc="Disable File Explorer shell integration overlays (reduces CPU in multi-session)" }
        @{ Key="DefaultRootDir";                   Hive="HKCU"; Path="Config";   Val="";   Type="String"; Desc="Optional: override default sync root from %USERPROFILE%\OneDrive to custom path" }
        @{ Key="BlockExternalSync";                Hive="HKLM"; Path="Policies"; Val=1;    Type="DWord";  Desc="Block syncing external/guest SharePoint libraries (security — retail environment)" }
        @{ Key="DisablePauseOnBatterySaver";       Hive="HKLM"; Path="Policies"; Val=1;    Type="DWord";  Desc="Disable pause-on-battery-saver (session hosts don't have batteries)" }
        @{ Key="DisablePauseOnMeteredNetwork";     Hive="HKLM"; Path="Policies"; Val=0;    Type="DWord";  Desc="0=Pause on metered network (default), set 1 only if session hosts on unmetered only" }
    )
}

function Show-OneDriveMenu {
    fx-hdr "ONEDRIVE — SILENT SIGN-IN & KNOWN FOLDER MOVE" $FC.Azure

    while ($true) {
        $od = $Script:Cfg.OneDrive
        fx-sec "OneDrive Configuration Status"

        $checks = @(
            @{ Label="Tenant ID configured";      Ok=($od.TenantId -ne "")       }
            @{ Label="Silent Sign-In enabled";    Ok=$od.SilentSignIn             }
            @{ Label="KFM enabled";               Ok=$od.KFMEnabled               }
            @{ Label="KFM Block Opt-Out";         Ok=$od.KFMBlockOptOut           }
            @{ Label="Files On Demand";           Ok=$od.FilesOnDemand            }
            @{ Label="Personal sync blocked";     Ok=$od.BlockPersonalSync        }
            @{ Label="Per-machine install";       Ok=$od.PerMachineInstall        }
            @{ Label="FSLogix exclusions added";  Ok=$od.FSLogixExclAdded         }
            @{ Label="Configs deployed";          Ok=$od.Deployed                 }
        )
        foreach ($c in $checks) {
            $icon = if ($c.Ok) {"✔"} else {"○"}
            $col  = if ($c.Ok) {$FC.Ok} else {$FC.Muted}
            Write-Host ("  $icon  {0}" -f $c.Label) -ForegroundColor $col
        }
        Write-Host ""
        if ($od.TenantId) { fx-info "Tenant ID" $od.TenantId }

        Write-Host "  1.  Run OneDrive Configuration Wizard" -ForegroundColor $FC.New
        Write-Host "  2.  Edit individual OneDrive settings" -ForegroundColor $FC.Menu
        Write-Host "  3.  Export — Intune Settings Catalog (JSON)" -ForegroundColor $FC.Menu
        Write-Host "  4.  Export — Group Policy / ADMX reference" -ForegroundColor $FC.Menu
        Write-Host "  5.  Export — Registry script (session host direct deploy)" -ForegroundColor $FC.Menu
        Write-Host "  6.  Generate per-machine install script (golden image)" -ForegroundColor $FC.Menu
        Write-Host "  7.  Update FSLogix Redirections.xml for OneDrive" -ForegroundColor $FC.Menu
        Write-Host "  8.  Validate KFM status on live session hosts" -ForegroundColor $FC.Azure
        Write-Host "  9.  OneDrive diagnostics & troubleshooting" -ForegroundColor $FC.Warn
        Write-Host " 10.  View full settings reference" -ForegroundColor $FC.Menu
        Write-Host "  0.  Back" -ForegroundColor $FC.Muted

        $ch = Read-Inp "Choice" -V @("0","1","2","3","4","5","6","7","8","9","10")
        switch ($ch) {
            "1"  { Invoke-OneDriveWizard }
            "2"  { Edit-OneDriveSettings }
            "3"  { Export-OneDriveIntunePolicy }
            "4"  { Export-OneDriveGPOSettings }
            "5"  { Export-OneDriveRegistryScript }
            "6"  { Export-OneDriveGoldenImageScript }
            "7"  { Update-FSLogixForOneDrive }
            "8"  { Test-OneDriveKFMStatus }
            "9"  { Show-OneDriveDiagnostics }
            "10" { Show-OneDriveSettingsReference }
            "0"  { return }
        }
    }
}

# ── WIZARD ─────────────────────────────────────────────────────────────────

function Invoke-OneDriveWizard {
    fx-hdr "ONEDRIVE CONFIGURATION WIZARD"
    $od = $Script:Cfg.OneDrive

    # Step 1 — Tenant ID
    fx-sec "Step 1 of 6 — Tenant ID"
    fx-info "The Tenant ID is required for both Silent Sign-In and KFM."
    fx-info "It tells OneDrive which Entra ID tenant to authenticate against."

    # Auto-detect from connected Azure session
    $ctx = Get-AzContext -EA SilentlyContinue
    if ($ctx -and $ctx.Tenant.Id) {
        fx-info "Detected Tenant ID from Azure context" $ctx.Tenant.Id
        if (Read-YN "Use this Tenant ID?" $true) {
            $od.TenantId = $ctx.Tenant.Id
        }
    }
    if (-not $od.TenantId) {
        $od.TenantId = Read-Inp "Enter your Entra ID Tenant ID (GUID)" -D $Script:Cfg.TenantId
    }
    fx-ok "Tenant ID" $od.TenantId

    # Step 2 — Silent Sign-In
    fx-sec "Step 2 of 6 — Silent Sign-In (SilentAccountConfig)"
    Write-Host @"
  Silent Sign-In automatically signs the user into OneDrive using their
  Windows/Entra ID session — ZERO prompts. When the user logs into AVD,
  OneDrive starts syncing immediately in the background.

  Requirements:
  • Device must be Azure AD Joined (Entra ID) OR Hybrid AD Joined
  • User must have a Microsoft 365 E3 or F1 license with OneDrive entitlement
  • Modern Authentication (ADAL) must be enabled

"@ -ForegroundColor $FC.Muted
    $od.SilentSignIn = Read-YN "Enable Silent Sign-In? (strongly recommended)" $true
    fx-ok "Silent Sign-In" (if ($od.SilentSignIn) {"Enabled"} else {"Disabled"})

    # Step 3 — KFM
    fx-sec "Step 3 of 6 — Known Folder Move (KFM)"
    Write-Host @"
  Known Folder Move silently redirects Desktop, Documents, and Pictures
  from the local user profile into OneDrive cloud storage.

  CRITICAL BENEFIT for AVD + FSLogix:
  • These folders are typically the biggest items in a user profile VHDX
  • With KFM, they live in OneDrive — NOT in the FSLogix container
  • Result: smaller profile VHDXs, faster logon times, zero data loss risk
  • Files accessible from ANY device (laptop, iPad, thin client, phone)

  KFM Folders:  Desktop  •  Documents  •  Pictures

"@ -ForegroundColor $FC.Muted
    $od.KFMEnabled = Read-YN "Enable KFM (silent redirect Desktop/Documents/Pictures to OneDrive)?" $true

    if ($od.KFMEnabled) {
        $od.KFMSilentOptIn  = Read-YN "Silent opt-in (no user prompt)?" $true
        $od.KFMNotification = Read-YN "Show a notification toast after KFM completes?" $false
        $od.KFMBlockOptOut  = Read-YN "Block users from moving folders back to local? (enforce policy)" $true

        # Folder selection
        fx-sec "Which folders to redirect to OneDrive?"
        $allFolders = @("Desktop","Documents","Pictures")
        $od.KFMFolders = @()
        foreach ($f in $allFolders) {
            if (Read-YN "  Redirect $f to OneDrive?" $true) { $od.KFMFolders += $f }
        }
        fx-ok "KFM folders" ($od.KFMFolders -join ", ")
    }

    # Step 4 — Files on Demand
    fx-sec "Step 4 of 6 — Files On Demand"
    Write-Host @"
  Files On Demand keeps files in OneDrive cloud storage and downloads them
  only when accessed. This is ESSENTIAL for AVD:
  • Session host disk is shared — you cannot let OneDrive download GB of files
  • Files appear in File Explorer with a cloud icon — open and they download instantly
  • Placeholder files use <1 KB of disk space regardless of actual file size

"@ -ForegroundColor $FC.Muted
    $od.FilesOnDemand = Read-YN "Enable Files On Demand? (REQUIRED for AVD — strongly recommended)" $true
    if (-not $od.FilesOnDemand) {
        fx-warn "WARNING" "Disabling Files On Demand on AVD session hosts can cause disk space exhaustion on shared VMs"
    }

    # Step 5 — Security settings
    fx-sec "Step 5 of 6 — Security & Tenant Restrictions"
    $od.BlockPersonalSync = Read-YN "Block personal OneDrive accounts (only allow BDF corporate account)?" $true
    fx-info "This prevents store workers from syncing personal files on corporate session hosts" ""

    $od.PerMachineInstall = Read-YN "Install OneDrive per-machine (REQUIRED for AVD multi-session)?" $true
    fx-info "Per-machine install puts OneDrive in C:\\Program Files — all users on the shared VM get their own sync" ""

    # Step 6 — FSLogix integration
    fx-sec "Step 6 of 6 — FSLogix Integration"
    Write-Host @"
  With OneDrive KFM active, several FSLogix profile exclusions become
  available or necessary:

  ADD to Redirections.xml (exclude from FSLogix VHDX — now in OneDrive):
    %USERPROFILE%\Desktop        ← now synced via KFM
    %USERPROFILE%\Documents      ← now synced via KFM
    %USERPROFILE%\Pictures       ← now synced via KFM

  KEEP in FSLogix profile (needed for OneDrive state persistence):
    %APPDATA%\Microsoft\OneDrive   (sync state, account config)
    %LOCALAPPDATA%\Microsoft\OneDrive\*.ldb  (local database)

  The wizard will update your Redirections.xml automatically.

"@ -ForegroundColor $FC.Muted
    if (Read-YN "Update FSLogix Redirections.xml to add OneDrive exclusions?" $true) {
        $od.FSLogixExclAdded = $true
    }

    # Export everything
    fx-sec "Generating All OneDrive Configuration Files"
    $od.DisableFirstRun = $true
    $od.Deployed = $true
    Save-FXConfig

    Export-OneDriveIntunePolicy
    Export-OneDriveGPOSettings
    Export-OneDriveRegistryScript
    Export-OneDriveGoldenImageScript
    if ($od.FSLogixExclAdded) { Update-FSLogixForOneDrive }

    fx-ok "OneDrive wizard complete" "All configuration files generated"
    Write-Host ""
    fx-info "Next steps" ""
    Write-Host "  1. Upload Intune JSON policies to Intune → Devices → Configuration" -ForegroundColor $FC.Muted
    Write-Host "  2. Assign policies to E3/F1 device groups" -ForegroundColor $FC.Muted
    Write-Host "  3. Run golden image script on next image build" -ForegroundColor $FC.Muted
    Write-Host "  4. Upload updated Redirections.xml to Azure Files shares" -ForegroundColor $FC.Muted
    Write-Host "  5. Validate KFM on pilot users after next session host boot" -ForegroundColor $FC.Muted
    Pause-Screen
}

# ── EDIT INDIVIDUAL SETTINGS ───────────────────────────────────────────────

function Edit-OneDriveSettings {
    fx-sec "Edit Individual OneDrive Settings"
    $od = $Script:Cfg.OneDrive
    $props = @(
        @{ Key="TenantId";          Label="Tenant ID (GUID)";            Type="String" }
        @{ Key="SilentSignIn";      Label="Silent Sign-In";              Type="Bool"   }
        @{ Key="KFMEnabled";        Label="KFM Enabled";                 Type="Bool"   }
        @{ Key="KFMSilentOptIn";    Label="KFM Silent (no prompt)";      Type="Bool"   }
        @{ Key="KFMNotification";   Label="KFM Show Notification";       Type="Bool"   }
        @{ Key="KFMBlockOptOut";    Label="KFM Block Opt-Out";           Type="Bool"   }
        @{ Key="FilesOnDemand";     Label="Files On Demand";             Type="Bool"   }
        @{ Key="BlockPersonalSync"; Label="Block Personal Sync";         Type="Bool"   }
        @{ Key="PerMachineInstall"; Label="Per-Machine Install";         Type="Bool"   }
        @{ Key="DisableFirstRun";   Label="Disable First Run Tutorial";  Type="Bool"   }
    )
    $i = 1
    foreach ($p in $props) {
        $cur = $od[$p.Key]
        Write-Host ("  {0,3}. {1,-38} = {2}" -f $i, $p.Label, $cur) -ForegroundColor $FC.Menu; $i++
    }
    $ch = Read-Inp "Edit setting #" (1..$props.Count | ForEach-Object {"$_"})
    $prop = $props[[int]$ch - 1]
    if ($prop.Type -eq "Bool") {
        $od[$prop.Key] = Read-YN "$($prop.Label)?" $od[$prop.Key]
    } else {
        $od[$prop.Key] = Read-Inp $prop.Label -D $od[$prop.Key]
    }
    fx-ok "Updated" "$($prop.Label) = $($od[$prop.Key])"
    Save-FXConfig
    Pause-Screen
}

# ── INTUNE EXPORT ─────────────────────────────────────────────────────────

function Export-OneDriveIntunePolicy {
    fx-sec "Exporting Intune Settings Catalog JSON Policies"
    $od    = $Script:Cfg.OneDrive
    $tenId = $od.TenantId

    # Policy 1: OneDrive Core (silent sign-in, tenant restrictions)
    $policy1 = @{
        "@odata.type" = "#microsoft.graph.windows10CustomConfiguration"
        "displayName" = "BDF-AVD-OneDrive-Core-Config"
        "description" = "OneDrive per-machine silent sign-in, tenant restrictions, Files On Demand"
        "omaSettings"  = @(
            @{ "@odata.type"="#microsoft.graph.omaSettingInteger";  displayName="Silent Account Config";        omaUri="./Device/Vendor/MSFT/Policy/Config/OneDrive~Policy~OneDriveNGSC/SilentAccountConfig";        value=1 }
            @{ "@odata.type"="#microsoft.graph.omaSettingString";   displayName="Allow Tenant List";            omaUri="./Device/Vendor/MSFT/Policy/Config/OneDrive~Policy~OneDriveNGSC/AllowTenantList";            value="<enabled/><data id=`"AllowTenantList_Prompt`" value=`"$tenId`"/>" }
            @{ "@odata.type"="#microsoft.graph.omaSettingInteger";  displayName="Enable ADAL (Modern Auth)";   omaUri="./Device/Vendor/MSFT/Policy/Config/OneDrive~Policy~OneDriveNGSC/EnableADAL";                  value=1 }
            @{ "@odata.type"="#microsoft.graph.omaSettingInteger";  displayName="Files On Demand";             omaUri="./Device/Vendor/MSFT/Policy/Config/OneDrive~Policy~OneDriveNGSC/FilesOnDemandEnabled";       value=1 }
            @{ "@odata.type"="#microsoft.graph.omaSettingInteger";  displayName="Dehydrate Team Sites";        omaUri="./Device/Vendor/MSFT/Policy/Config/OneDrive~Policy~OneDriveNGSC/DehydrateSyncedTeamSites";   value=1 }
            @{ "@odata.type"="#microsoft.graph.omaSettingInteger";  displayName="Disable Personal Sync";       omaUri="./Device/Vendor/MSFT/Policy/Config/OneDrive~Policy~OneDriveNGSC/DisablePersonalSync";        value=$(if($od.BlockPersonalSync){1}else{0}) }
            @{ "@odata.type"="#microsoft.graph.omaSettingInteger";  displayName="Disable First Delete Dialog"; omaUri="./Device/Vendor/MSFT/Policy/Config/OneDrive~Policy~OneDriveNGSC/DisableFirstDeleteDialog";   value=1 }
            @{ "@odata.type"="#microsoft.graph.omaSettingInteger";  displayName="Disable FRE Tutorial";       omaUri="./Device/Vendor/MSFT/Policy/Config/OneDrive~Policy~OneDriveNGSC/DisableFRETutorial";          value=1 }
            @{ "@odata.type"="#microsoft.graph.omaSettingInteger";  displayName="Prevent Traffic Pre Sign-In"; omaUri="./Device/Vendor/MSFT/Policy/Config/OneDrive~Policy~OneDriveNGSC/PreventNetworkTrafficPreUserSignIn"; value=1 }
            @{ "@odata.type"="#microsoft.graph.omaSettingInteger";  displayName="Enable Hold The File";       omaUri="./Device/Vendor/MSFT/Policy/Config/OneDrive~Policy~OneDriveNGSC/EnableHoldTheFile";           value=1 }
            @{ "@odata.type"="#microsoft.graph.omaSettingInteger";  displayName="Disable Battery Saver Pause"; omaUri="./Device/Vendor/MSFT/Policy/Config/OneDrive~Policy~OneDriveNGSC/DisablePauseOnBatterySaver"; value=1 }
            @{ "@odata.type"="#microsoft.graph.omaSettingInteger";  displayName="Warn Low Disk Space (2GB)";  omaUri="./Device/Vendor/MSFT/Policy/Config/OneDrive~Policy~OneDriveNGSC/WarningMinDiskSpaceMB";       value=2048 }
            @{ "@odata.type"="#microsoft.graph.omaSettingInteger";  displayName="Block External Sync";        omaUri="./Device/Vendor/MSFT/Policy/Config/OneDrive~Policy~OneDriveNGSC/BlockExternalSync";           value=1 }
        )
    }

    # Policy 2: KFM (Known Folder Move)
    $policy2 = @{
        "@odata.type" = "#microsoft.graph.windows10CustomConfiguration"
        "displayName" = "BDF-AVD-OneDrive-KFM-Policy"
        "description" = "OneDrive Known Folder Move — silently redirects Desktop/Documents/Pictures to OneDrive cloud storage"
        "omaSettings"  = @(
            @{ "@odata.type"="#microsoft.graph.omaSettingString";  displayName="KFM Silent Opt-In";       omaUri="./Device/Vendor/MSFT/Policy/Config/OneDrive~Policy~OneDriveNGSC~KFM/KFMSilentOptIn";             value="<enabled/><data id=`"KFMSilentOptIn_Prompt`" value=`"$tenId`"/>" }
            @{ "@odata.type"="#microsoft.graph.omaSettingInteger"; displayName="KFM Silent No Notify";    omaUri="./Device/Vendor/MSFT/Policy/Config/OneDrive~Policy~OneDriveNGSC~KFM/KFMSilentOptInWithNotification"; value=$(if($od.KFMNotification){1}else{0}) }
            @{ "@odata.type"="#microsoft.graph.omaSettingInteger"; displayName="KFM Block Opt-Out";       omaUri="./Device/Vendor/MSFT/Policy/Config/OneDrive~Policy~OneDriveNGSC~KFM/KFMBlockOptOut";              value=$(if($od.KFMBlockOptOut){1}else{0}) }
            @{ "@odata.type"="#microsoft.graph.omaSettingInteger"; displayName="KFM Block Opt-In wizard"; omaUri="./Device/Vendor/MSFT/Policy/Config/OneDrive~Policy~OneDriveNGSC~KFM/KFMBlockOptIn";               value=0 }
            @{ "@odata.type"="#microsoft.graph.omaSettingInteger"; displayName="KFM No Wizard";           omaUri="./Device/Vendor/MSFT/Policy/Config/OneDrive~Policy~OneDriveNGSC~KFM/KFMOptInWithWizard";          value=0 }
        )
    }

    # Policy 3: AVD Multi-Session specific
    $policy3 = @{
        "@odata.type" = "#microsoft.graph.windows10CustomConfiguration"
        "displayName" = "BDF-AVD-OneDrive-MultiSession-Tuning"
        "description" = "OneDrive settings optimized for AVD pooled multi-session hosts — shell integration off, Office co-auth on"
        "omaSettings"  = @(
            @{ "@odata.type"="#microsoft.graph.omaSettingInteger"; displayName="Enable Office Co-authoring"; omaUri="./Device/Vendor/MSFT/Policy/Config/OneDrive~Policy~OneDriveNGSC/EnableAllOcsiClients";       value=1 }
            @{ "@odata.type"="#microsoft.graph.omaSettingInteger"; displayName="Shell Integrator Off (AVD)"; omaUri="./Device/Vendor/MSFT/Policy/Config/OneDrive~Policy~OneDriveNGSC/DisableNewFileExperience";    value=1 }
        )
    }

    foreach ($pol in @($policy1, $policy2, $policy3)) {
        $file = ".\OneDrive-Intune-$($pol.displayName).json"
        $pol | ConvertTo-Json -Depth 10 | Set-Content $file -Encoding UTF8
        fx-ok "Intune policy exported" $file
    }
    fx-info "Import into" "Intune → Devices → Configuration → + Create → Windows 10 and later → Custom"
    fx-info "Assign to"   "E3 Device Group AND F1 Device Group"
    Pause-Screen
}

# ── GPO EXPORT ────────────────────────────────────────────────────────────

function Export-OneDriveGPOSettings {
    fx-sec "Exporting Group Policy Reference"
    $od   = $Script:Cfg.OneDrive
    $file = ".\OneDrive-GPO-Settings.txt"
    $lines = @(
        "# ═══════════════════════════════════════════════════════════════════════",
        "# BDF OneDrive Group Policy Settings Reference",
        "# Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        "# Tenant ID: $($od.TenantId)",
        "# ADMX location: https://aka.ms/OneDriveAdmx",
        "# GPO Path: Computer Configuration → Policies → Admin Templates → OneDrive",
        "# ═══════════════════════════════════════════════════════════════════════",
        "",
        "# ── AUTHENTICATION ──────────────────────────────────────────────────",
        "Use OneDrive Files On Demand                                  = Enabled",
        "Silently sign in users to the OneDrive sync app               = Enabled",
        "Allow syncing OneDrive accounts for only specific orgs        = Enabled  | Value: $($od.TenantId)",
        "Prevent users from syncing personal OneDrive accounts         = $(if($od.BlockPersonalSync){'Enabled'}else{'Disabled'})",
        "Enable Modern Authentication (ADAL)                           = Enabled",
        "Prevent OneDrive from generating network traffic              = Enabled",
        "  until the user signs in to OneDrive  (startup performance)",
        "",
        "# ── KNOWN FOLDER MOVE ───────────────────────────────────────────────",
        "Silently move Windows known folders to OneDrive               = Enabled  | Tenant ID: $($od.TenantId)",
        "  Show notification to users after folders redirected         = $(if($od.KFMNotification){'Show (value:1)'}else{'Silent (value:0)'})",
        "Prevent users from redirecting their Windows known folders    = $(if($od.KFMBlockOptOut){'Enabled (blocks opt-out)'}else{'Disabled'})",
        "Prevent users from moving their Windows known folders to      = Disabled  (must be disabled for KFM to work)",
        "  OneDrive",
        "Prompt users to move Windows known folders to OneDrive        = Disabled  (we use Silent, not Prompt)",
        "",
        "# ── FILES ON DEMAND ─────────────────────────────────────────────────",
        "Use OneDrive Files On Demand                                  = $(if($od.FilesOnDemand){'Enabled'}else{'Disabled (NOT RECOMMENDED for AVD)'})",
        "Dehydrate team sites when users access them on demand         = Enabled",
        "",
        "# ── PERFORMANCE & EXPERIENCE ────────────────────────────────────────",
        "Enable Office co-authoring for OneDrive files                 = Enabled",
        "Disable the tutorial that appears at end of OneDrive Setup    = Enabled",
        "Silently remove the Personal tab in OneDrive for Business     = Enabled",
        "Prevent users from seeing the 'Manage Storage' page           = Enabled",
        "Warn users who are low on disk space                          = Enabled  | 2048 MB",
        "Block syncing SharePoint Online libraries with labels         = Enabled  (data governance)",
        "Block external sync                                           = Enabled  (block external org SharePoint)",
        "",
        "# ── AVD MULTI-SESSION NOTES ─────────────────────────────────────────",
        "# OneDrive MUST be installed per-machine for AVD pooled hosts:",
        "# Run: OneDriveSetup.exe /allusers  (during golden image build)",
        "# Per-machine install path: C:\Program Files\Microsoft OneDrive\",
        "# Each user gets their own sync scope — config above applies machine-wide",
        "",
        "# ── DEPLOYMENT TARGET ───────────────────────────────────────────────",
        "# Apply to: OU=AVD Session Hosts,DC=bdf,DC=internal",
        "# Or via Intune: Assigned to E3 + F1 Device Groups",
        "# Note: OneDrive machine-wide ADMX must be imported:",
        "#   Get from: C:\Program Files\Microsoft OneDrive\<version>\adm\"
    )
    $lines | Set-Content $file -Encoding UTF8
    fx-ok "GPO settings reference" $file
    Pause-Screen
}

# ── REGISTRY SCRIPT ───────────────────────────────────────────────────────

function Export-OneDriveRegistryScript {
    fx-sec "Generating Registry Deployment Script"
    $od   = $Script:Cfg.OneDrive
    $file = ".\Deploy-OneDrive-Registry.ps1"

    $script = @"
<#
.SYNOPSIS  Deploy OneDrive policy settings via registry on AVD session hosts.
           Run as SYSTEM during image build OR via Intune Management Extension.
           For hybrid AD or any scenario where Intune OMA-URI doesn't cover everything.
.NOTES     Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
           Tenant ID: $($od.TenantId)
           Run as: SYSTEM or local administrator
#>

`$TenantId = '$($od.TenantId)'

function Set-ODReg {
    param([string]`$Path,[string]`$Name,[object]`$Value,[string]`$Type="DWord")
    if (-not (Test-Path `$Path)) { New-Item -Path `$Path -Force | Out-Null }
    Set-ItemProperty -Path `$Path -Name `$Name -Value `$Value -Type `$Type -Force
    Write-Host ("  SET [{0}] {1} = {2}" -f (Split-Path `$Path -Leaf), `$Name, `$Value) -ForegroundColor Cyan
}

Write-Host "" ; Write-Host "  BDF OneDrive Registry Configuration" -ForegroundColor Blue
Write-Host "  Tenant ID: `$TenantId" -ForegroundColor Cyan ; Write-Host ""

`$polPath = 'HKLM:\SOFTWARE\Policies\Microsoft\OneDrive'
`$cfgPath = 'HKLM:\SOFTWARE\Microsoft\OneDrive'

# ── AUTHENTICATION ────────────────────────────────────────────────────────
Write-Host "  ── Authentication & Sign-In" -ForegroundColor Magenta
Set-ODReg `$polPath "SilentAccountConfig"                1
Set-ODReg `$cfgPath "EnableADAL"                         1
Set-ODReg `$polPath "PreventNetworkTrafficPreUserSignIn" 1
$(if ($od.BlockPersonalSync) {"Set-ODReg `$polPath 'DisablePersonalSync'                1"})

# Allow only BDF tenant
Set-ODReg "`$polPath\AllowTenantList" `$TenantId         `$TenantId "String"

# ── KNOWN FOLDER MOVE ────────────────────────────────────────────────────
Write-Host "  ── Known Folder Move" -ForegroundColor Magenta
$(if ($od.KFMEnabled) {
"Set-ODReg `$polPath 'KFMSilentOptIn'                    `$TenantId 'String'
Set-ODReg `$polPath 'KFMSilentOptInWithNotification'    $(if($od.KFMNotification){1}else{0})
Set-ODReg `$polPath 'KFMBlockOptIn'                     0   # MUST be 0 for silent opt-in to work
Set-ODReg `$polPath 'KFMBlockOptOut'                    $(if($od.KFMBlockOptOut){1}else{0})
Set-ODReg `$polPath 'KFMOptInWithWizard'                0   # No wizard — silent only"
} else {
"Write-Host '  KFM disabled in config — skipping' -ForegroundColor Yellow"
})

# ── FILES ON DEMAND ───────────────────────────────────────────────────────
Write-Host "  ── Files On Demand" -ForegroundColor Magenta
Set-ODReg `$polPath "FilesOnDemandEnabled"               $(if($od.FilesOnDemand){1}else{0})
Set-ODReg `$polPath "DehydrateSyncedTeamSites"           1

# ── PERFORMANCE & UX ─────────────────────────────────────────────────────
Write-Host "  ── Performance & User Experience" -ForegroundColor Magenta
Set-ODReg `$polPath "EnableAllOcsiClients"               1   # Office co-authoring
Set-ODReg `$polPath "DisableFirstDeleteDialog"           1
Set-ODReg `$polPath "DisableFRETutorial"                 1
Set-ODReg `$polPath "EnableHoldTheFile"                  1
Set-ODReg `$polPath "BlockExternalSync"                  1
Set-ODReg `$polPath "DisablePauseOnBatterySaver"         1   # No batteries on session hosts
Set-ODReg `$polPath "WarningMinDiskSpaceMB"              2048
Set-ODReg `$cfgPath "ShellIntegratorEnabled"             0   # Reduce CPU in multi-session

# ── AVD MULTI-SESSION: Startup Registration ───────────────────────────────
Write-Host "  ── AVD Startup Configuration" -ForegroundColor Magenta
# For per-machine install, OneDrive adds itself to HKLM Run automatically.
# Verify / force it here as a safety net:
`$odExe = "`$env:ProgramFiles\Microsoft OneDrive\OneDrive.exe"
if (Test-Path `$odExe) {
    Set-ODReg 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run' `
        "OneDrive" "`"``$odExe``" /background" "String"
    Write-Host "  OneDrive startup entry set (per-machine)" -ForegroundColor Green
} else {
    Write-Host "  OneDrive.exe not found at `$odExe — run golden image install script first" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "  OneDrive registry configuration complete." -ForegroundColor Green
Write-Host "  KFM will trigger on next user sign-in to their OneDrive account." -ForegroundColor Cyan
Write-Host "  Validate: Run 'Test-OneDriveKFMStatus' after deploying to pilot users." -ForegroundColor Cyan
"@
    $script | Set-Content $file -Encoding UTF8
    fx-ok "Registry deployment script" $file
    Pause-Screen
}

# ── GOLDEN IMAGE INSTALL SCRIPT ────────────────────────────────────────────

function Export-OneDriveGoldenImageScript {
    fx-sec "Generating Golden Image Install Script"
    $od   = $Script:Cfg.OneDrive
    $file = ".\Install-OneDrive-GoldenImage.ps1"

    $script = @"
<#
.SYNOPSIS  Install and configure OneDrive per-machine on AVD golden image.
           Run ONCE during image build (before Sysprep capture).
           Handles download, per-machine install, policy pre-seeding.
.NOTES     Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
           MUST run before Sysprep — sets machine-wide config, not user-specific.
           Run as: SYSTEM or local administrator with internet access.
#>

Set-StrictMode -Version Latest
`$ErrorActionPreference = "Stop"

Write-Host "" ; Write-Host "  BDF OneDrive Golden Image Installer" -ForegroundColor Blue ; Write-Host ""

# ── STEP 1: Remove existing per-user OneDrive install ─────────────────────
Write-Host "  Step 1: Removing per-user OneDrive install (if present)..." -ForegroundColor Magenta
`$perUserPaths = @(
    "`$env:LOCALAPPDATA\Microsoft\OneDrive\OneDriveSetup.exe",
    "`$env:USERPROFILE\AppData\Local\Microsoft\OneDrive\OneDriveSetup.exe"
)
foreach (`$p in `$perUserPaths) {
    if (Test-Path `$p) {
        Write-Host "  Uninstalling per-user OneDrive: `$p" -ForegroundColor Yellow
        Start-Process `$p -ArgumentList "/uninstall" -Wait -EA SilentlyContinue
    }
}

# Stop OneDrive processes
Stop-Process -Name "OneDrive" -Force -EA SilentlyContinue
Stop-Process -Name "OneDriveSetup" -Force -EA SilentlyContinue
Start-Sleep -Seconds 3

# ── STEP 2: Download latest OneDrive per-machine installer ─────────────────
Write-Host "  Step 2: Downloading OneDrive per-machine installer..." -ForegroundColor Magenta
`$dlPath = "`$env:TEMP\OneDriveSetup.exe"

# Enterprise ring (stable, tested for enterprise AVD)
`$dlUrl = "https://go.microsoft.com/fwlink/?linkid=844652"  # Per-machine / AllUsers installer

try {
    Invoke-WebRequest -Uri `$dlUrl -OutFile `$dlPath -UseBasicParsing
    Write-Host "  Downloaded: `$dlPath" -ForegroundColor Green
} catch {
    Write-Host "  Download failed — trying alternate URL" -ForegroundColor Yellow
    `$dlUrl2 = "https://oneclient.sfx.ms/Win/Installers/OneDriveSetup.exe"
    Invoke-WebRequest -Uri `$dlUrl2 -OutFile `$dlPath -UseBasicParsing
}

# Verify download
`$fileInfo = Get-Item `$dlPath
Write-Host "  File size: `$([Math]::Round(`$fileInfo.Length/1MB,1)) MB" -ForegroundColor Cyan

# ── STEP 3: Install per-machine (/allusers flag) ──────────────────────────
Write-Host "  Step 3: Installing OneDrive per-machine (/allusers)..." -ForegroundColor Magenta
`$installArgs = "/allusers"
`$proc = Start-Process `$dlPath -ArgumentList `$installArgs -Wait -PassThru
if (`$proc.ExitCode -eq 0 -or `$proc.ExitCode -eq 3010) {
    Write-Host "  OneDrive per-machine install complete (exit: `$(`$proc.ExitCode))" -ForegroundColor Green
} else {
    Write-Host "  Install may have issues (exit: `$(`$proc.ExitCode)) — check C:\Windows\Logs\CBS\" -ForegroundColor Yellow
}

# Verify install location
`$perMachinePath = "`$env:ProgramFiles\Microsoft OneDrive"
if (Test-Path `$perMachinePath) {
    `$ver = (Get-Item "`$perMachinePath\OneDrive.exe" -EA SilentlyContinue).VersionInfo.FileVersion
    Write-Host "  Installed at: `$perMachinePath (v`$ver)" -ForegroundColor Green
} else {
    Write-Host "  WARNING: Expected path not found: `$perMachinePath" -ForegroundColor Red
}

# ── STEP 4: Apply registry policy pre-seed ────────────────────────────────
Write-Host "  Step 4: Pre-seeding registry policies..." -ForegroundColor Magenta
`$TenantId = '$($od.TenantId)'
`$polPath  = 'HKLM:\SOFTWARE\Policies\Microsoft\OneDrive'

function Preset-ODReg { param(`$Path,`$Name,`$Value,`$Type="DWord")
    if (-not (Test-Path `$Path)) { New-Item -Path `$Path -Force | Out-Null }
    Set-ItemProperty -Path `$Path -Name `$Name -Value `$Value -Type `$Type -Force
}

Preset-ODReg `$polPath "SilentAccountConfig"                1
Preset-ODReg `$polPath "EnableADAL"                         1  -Path "HKLM:\SOFTWARE\Microsoft\OneDrive"
Preset-ODReg `$polPath "FilesOnDemandEnabled"               1
Preset-ODReg `$polPath "DehydrateSyncedTeamSites"           1
Preset-ODReg `$polPath "DisablePersonalSync"                1
Preset-ODReg `$polPath "PreventNetworkTrafficPreUserSignIn" 1
Preset-ODReg `$polPath "DisableFRETutorial"                 1
Preset-ODReg `$polPath "DisableFirstDeleteDialog"           1
Preset-ODReg `$polPath "EnableHoldTheFile"                  1
Preset-ODReg `$polPath "BlockExternalSync"                  1
Preset-ODReg `$polPath "EnableAllOcsiClients"               1
Preset-ODReg `$polPath "DisablePauseOnBatterySaver"         1
Preset-ODReg `$polPath "WarningMinDiskSpaceMB"              2048
Preset-ODReg `$polPath "KFMSilentOptIn"                     `$TenantId "String"
Preset-ODReg `$polPath "KFMSilentOptInWithNotification"     0
Preset-ODReg `$polPath "KFMBlockOptIn"                      0
Preset-ODReg `$polPath "KFMBlockOptOut"                     1
Preset-ODReg `$polPath "KFMOptInWithWizard"                 0
Preset-ODReg "`$polPath\AllowTenantList" `$TenantId         `$TenantId "String"
Preset-ODReg "HKLM:\SOFTWARE\Microsoft\OneDrive" "ShellIntegratorEnabled" 0

Write-Host "  Registry pre-seed complete." -ForegroundColor Green

# ── STEP 5: Configure startup (per-machine) ────────────────────────────────
Write-Host "  Step 5: Configuring per-machine startup..." -ForegroundColor Magenta
`$oneDriveExe = "`$env:ProgramFiles\Microsoft OneDrive\OneDrive.exe"
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" `
    -Name "OneDrive" -Value "`"`$oneDriveExe`" /background" -Type String -Force
Write-Host "  OneDrive startup registered under HKLM (all users)" -ForegroundColor Green

# ── STEP 6: Exclude OneDrive Update tasks from image ─────────────────────
Write-Host "  Step 6: Disabling OneDrive auto-update task (managed via Intune)..." -ForegroundColor Magenta
Disable-ScheduledTask -TaskName "OneDrive Standalone Update Task v2" -EA SilentlyContinue | Out-Null
Write-Host "  Update task disabled." -ForegroundColor Green

# ── STEP 7: Validation ────────────────────────────────────────────────────
Write-Host "" ; Write-Host "  ── Post-Install Validation" -ForegroundColor Magenta
`$checks = @(
    @{ Name="OneDrive.exe exists";       Pass=Test-Path "`$env:ProgramFiles\Microsoft OneDrive\OneDrive.exe" }
    @{ Name="SilentAccountConfig=1";     Pass=((Get-ItemProperty `$polPath -Name SilentAccountConfig -EA SilentlyContinue).SilentAccountConfig -eq 1) }
    @{ Name="KFMSilentOptIn=TenantId";   Pass=((Get-ItemProperty `$polPath -Name KFMSilentOptIn -EA SilentlyContinue).KFMSilentOptIn -eq `$TenantId) }
    @{ Name="FilesOnDemandEnabled=1";    Pass=((Get-ItemProperty `$polPath -Name FilesOnDemandEnabled -EA SilentlyContinue).FilesOnDemandEnabled -eq 1) }
    @{ Name="DisablePersonalSync=1";     Pass=((Get-ItemProperty `$polPath -Name DisablePersonalSync -EA SilentlyContinue).DisablePersonalSync -eq 1) }
    @{ Name="HKLM Run key exists";       Pass=`$null -ne (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -Name OneDrive -EA SilentlyContinue) }
)
`$pass=0; `$fail=0
foreach (`$c in `$checks) {
    if (`$c.Pass) { Write-Host "  [PASS] `$(`$c.Name)" -ForegroundColor Green; `$pass++ }
    else          { Write-Host "  [FAIL] `$(`$c.Name)" -ForegroundColor Red;   `$fail++ }
}
Write-Host ""
Write-Host "  Result: `$pass/`$(`$checks.Count) checks passed" -ForegroundColor (if(`$fail -eq 0){'Green'}else{'Yellow'})
Write-Host ""
Write-Host "  IMPORTANT: Do NOT run Sysprep while OneDrive is running." -ForegroundColor Yellow
Write-Host "  Stop OneDrive before Sysprep: Stop-Process -Name OneDrive -Force" -ForegroundColor Cyan
Write-Host "  After Sysprep + capture, OneDrive will start on first user login and KFM will activate." -ForegroundColor Cyan
"@
    $script | Set-Content $file -Encoding UTF8
    fx-ok "Golden image install script" $file
    fx-info "Run order" "1) Run this script during image build  2) Stop OneDrive  3) Sysprep  4) Capture"
    Pause-Screen
}

# ── FSLOGIX INTEGRATION ───────────────────────────────────────────────────

function Update-FSLogixForOneDrive {
    fx-sec "Updating FSLogix Redirections.xml for OneDrive KFM"
    $od = $Script:Cfg.OneDrive

    $newExclusions = @()
    foreach ($f in $od.KFMFolders) {
        $newExclusions += @{
            Path = "%USERPROFILE%\$f"
            Type = "Directory"
            Desc = "$f folder — redirected to OneDrive via KFM (do not include in FSLogix VHDX)"
        }
    }
    $newExclusions += @(
        @{ Path="%LOCALAPPDATA%\Microsoft\OneDrive\logs";     Type="Directory"; Desc="OneDrive log files (not needed in profile)" }
        @{ Path="%LOCALAPPDATA%\Microsoft\OneDrive\setup";    Type="Directory"; Desc="OneDrive setup cache" }
        @{ Path="%LOCALAPPDATA%\OneDrive";                    Type="Directory"; Desc="OneDrive local state cache (rebuilt on sign-in)" }
    )

    # Check if Redirections.xml exists
    $existing = ""
    if (Test-Path ".\Redirections.xml") {
        $existing = Get-Content ".\Redirections.xml" -Raw
        fx-info "Found existing Redirections.xml" ".\Redirections.xml"
    }

    # Inject new exclusions before </FrxProfileFolderRedirection>
    $newXmlBlock = ""
    foreach ($excl in $newExclusions) {
        $newXmlBlock += "`n    <!-- $($excl.Desc) -->"
        $newXmlBlock += "`n    <Exclude Copy=""0"">$($excl.Path)</Exclude>"
    }

    if ($existing -and $existing -like "*</FrxProfileFolderRedirection>*") {
        $updated = $existing -replace "</Excludes>", "$newXmlBlock`n  </Excludes>"
        $updated | Set-Content ".\Redirections.xml" -Encoding UTF8
        fx-ok "Redirections.xml updated" "$($newExclusions.Count) OneDrive exclusions added"
    } else {
        # Create from scratch with OneDrive exclusions
        @(
            '<?xml version="1.0" encoding="UTF-8"?>',
            '<!-- BDF AVD FSLogix Redirections.xml — OneDrive KFM Integration -->',
            '<!-- Generated: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + ' -->',
            '<FrxProfileFolderRedirection ExcludeCommonFolders="49">',
            '  <Excludes>'
        ) | Set-Content ".\Redirections.xml" -Encoding UTF8
        foreach ($excl in $newExclusions) {
            "    <!-- $($excl.Desc) -->" | Add-Content ".\Redirections.xml" -Encoding UTF8
            "    <Exclude Copy=""0"">$($excl.Path)</Exclude>" | Add-Content ".\Redirections.xml" -Encoding UTF8
        }
        @('  </Excludes>', '  <Redirections/>', '</FrxProfileFolderRedirection>') |
            Add-Content ".\Redirections.xml" -Encoding UTF8
        fx-ok "Redirections.xml created with OneDrive exclusions" ""
    }

    fx-info "IMPORTANT" "Upload updated Redirections.xml to ALL profile shares:"
    Write-Host "  \\$($Script:Cfg.Storage.Account.Name).file.core.windows.net\$($Script:Cfg.Storage.Shares.E3)\Redirections.xml" -ForegroundColor $FC.Azure
    Write-Host "  \\$($Script:Cfg.Storage.Account.Name).file.core.windows.net\$($Script:Cfg.Storage.Shares.F1)\Redirections.xml" -ForegroundColor $FC.Azure

    # Generate upload script
    $uploadScript = @"
# Upload Redirections.xml to all FSLogix profile shares
`$saName = '$($Script:Cfg.Storage.Account.Name)'
`$saRg   = '$($Script:Cfg.Storage.Account.ResourceGroup)'
`$sa     = Get-AzStorageAccount -Name `$saName -ResourceGroupName `$saRg
`$ctx    = `$sa.Context
foreach (`$share in @('$($Script:Cfg.Storage.Shares.E3)','$($Script:Cfg.Storage.Shares.F1)')) {
    Set-AzStorageFileContent -ShareName `$share -Source '.\Redirections.xml' -Path 'Redirections.xml' -Context `$ctx -Force
    Write-Host "Uploaded to: `$share" -ForegroundColor Green
}
"@
    $uploadScript | Set-Content ".\Deploy-Redirections.ps1" -Encoding UTF8
    fx-ok "Upload script updated" ".\Deploy-Redirections.ps1"

    $Script:Cfg.OneDrive.FSLogixExclAdded = $true
    Save-FXConfig
    Pause-Screen
}

# ── LIVE KFM VALIDATION ────────────────────────────────────────────────────

function Test-OneDriveKFMStatus {
    fx-sec "KFM Status Check on Live Session Hosts"

    $checkScript = @'
$polPath  = 'HKLM:\SOFTWARE\Policies\Microsoft\OneDrive'
$results  = @{}

# Check registry policies
$results['SilentAccountConfig']   = (Get-ItemProperty $polPath -Name SilentAccountConfig   -EA SilentlyContinue).SilentAccountConfig   -eq 1
$results['KFMSilentOptIn set']    = (Get-ItemProperty $polPath -Name KFMSilentOptIn         -EA SilentlyContinue).KFMSilentOptIn -ne ""
$results['KFMBlockOptOut']        = (Get-ItemProperty $polPath -Name KFMBlockOptOut         -EA SilentlyContinue).KFMBlockOptOut         -eq 1
$results['FilesOnDemand']         = (Get-ItemProperty $polPath -Name FilesOnDemandEnabled   -EA SilentlyContinue).FilesOnDemandEnabled   -eq 1
$results['OneDrive.exe present']  = Test-Path "$env:ProgramFiles\Microsoft OneDrive\OneDrive.exe"
$results['OneDrive running']      = ($null -ne (Get-Process -Name OneDrive -EA SilentlyContinue))

# Check per-user KFM completion (run for each user profile)
$kfmDone = @()
$profiles = Get-ChildItem "C:\Users" -Directory | Where-Object { $_.Name -notmatch "^(Default|Public|All Users|desktop.ini)$" }
foreach ($prof in $profiles) {
    $userPolPath = "Registry::HKEY_USERS\$(((New-Object System.Security.Principal.NTAccount($prof.Name)).Translate([System.Security.Principal.SecurityIdentifier]).Value))\SOFTWARE\Microsoft\OneDrive"
    $kfmStatus = (Get-ItemProperty $userPolPath -Name "KFMState" -EA SilentlyContinue).KFMState
    $syncRoot  = (Get-ItemProperty "Registry::HKEY_USERS\*\SOFTWARE\Microsoft\OneDrive\Accounts\Business1" -Name "UserFolder" -EA SilentlyContinue).UserFolder
    $kfmDone += @{ User=$prof.Name; KFMState=$kfmStatus; SyncRoot=$syncRoot }
}

Write-Host "Machine Policy Status:"
foreach ($k in $results.GetEnumerator()) {
    $icon = if ($k.Value) {"[PASS]"} else {"[FAIL]"}
    $col  = if ($k.Value) {"Green"} else {"Red"}
    Write-Host ("  $icon {0}" -f $k.Key) -ForegroundColor $col
}
Write-Host "`nPer-User KFM State:"
foreach ($u in $kfmDone) {
    Write-Host ("  User: {0,-20} KFMState: {1}  SyncRoot: {2}" -f $u.User, $u.KFMState, $u.SyncRoot) -ForegroundColor Cyan
}
'@
    $hp = $Script:Cfg.HostPools.E3.Name
    $rg = $Script:Cfg.RG.AVD.Name
    $hosts = @(Get-AzWvdSessionHost -HostPoolName $hp -ResourceGroupName $rg -EA SilentlyContinue | Select-Object -First 2)
    if ($hosts.Count -eq 0) { fx-warn "No session hosts found — validate manually" ""; Pause-Screen; return }

    foreach ($h in $hosts) {
        $vmName = ($h.Name -split "/")[-1]
        fx-step "Checking $vmName..."
        $result = Invoke-AzVMRunCommand -ResourceGroupName $rg -VMName $vmName `
                      -CommandId RunPowerShellScript -ScriptString $checkScript -EA SilentlyContinue
        if ($result?.Value[0].Message) {
            Write-Host $result.Value[0].Message -ForegroundColor $FC.Muted
        }
    }
    Pause-Screen
}

# ── DIAGNOSTICS ───────────────────────────────────────────────────────────

function Show-OneDriveDiagnostics {
    fx-hdr "ONEDRIVE DIAGNOSTICS & TROUBLESHOOTING"

    while ($true) {
        fx-sec "Diagnostics Menu"
        Write-Host "  1.  Common OneDrive + AVD issues & fixes" -ForegroundColor $FC.Menu
        Write-Host "  2.  OneDrive KQL queries for Log Analytics" -ForegroundColor $FC.Azure
        Write-Host "  3.  Generate ODDiag collection script" -ForegroundColor $FC.Menu
        Write-Host "  0.  Back" -ForegroundColor $FC.Muted
        $ch = Read-Inp "Choice" -V @("0","1","2","3")
        switch ($ch) {
            "1" { Show-OneDriveIssues }
            "2" { Export-OneDriveKQL }
            "3" { Export-ODDiagScript }
            "0" { return }
        }
    }
}

function Show-OneDriveIssues {
    fx-sec "Common OneDrive Issues in AVD Multi-Session"
    Write-Host @"
  ┌─ ISSUE 1: OneDrive not starting / not signing in ────────────────────┐
  │  Symptom: OneDrive icon missing from system tray on login            │
  │  Cause:   Per-user install replacing per-machine, or startup key     │
  │           missing                                                    │
  │  Fix:                                                                │
  │  1. Verify per-machine install: ls 'C:\Program Files\Microsoft       │
  │     OneDrive\OneDrive.exe'                                           │
  │  2. Check HKLM Run key: Get-ItemProperty HKLM:\SOFTWARE\Microsoft\  │
  │     Windows\CurrentVersion\Run -Name OneDrive                        │
  │  3. Re-run Install-OneDrive-GoldenImage.ps1 on next image build      │
  └───────────────────────────────────────────────────────────────────────┘

  ┌─ ISSUE 2: KFM not completing / Desktop still local ──────────────────┐
  │  Symptom: User's Desktop/Documents still in C:\Users\name\           │
  │  Cause:   KFMSilentOptIn policy not applied, or user not yet signed  │
  │           into OneDrive, or KFMBlockOptIn=1 (must be 0)              │
  │  Fix:                                                                │
  │  1. Verify: (Get-ItemProperty HKLM:\SOFTWARE\Policies\Microsoft\    │
  │     OneDrive -Name KFMSilentOptIn).KFMSilentOptIn = TenantID        │
  │  2. Check KFMBlockOptIn = 0 (not 1 — blocks KFM entirely)           │
  │  3. Verify user is signed into OneDrive (check system tray)          │
  │  4. Run: C:\Program Files\Microsoft OneDrive\OneDrive.exe /reset    │
  └───────────────────────────────────────────────────────────────────────┘

  ┌─ ISSUE 3: Files On Demand not working / disk fills up ───────────────┐
  │  Symptom: Session host disk space critically low                     │
  │  Cause:   FilesOnDemandEnabled not set, or user pinned files         │
  │  Fix:                                                                │
  │  1. Verify: (Get-ItemProperty HKLM:\SOFTWARE\Policies\Microsoft\    │
  │     OneDrive -Name FilesOnDemandEnabled).FilesOnDemandEnabled = 1   │
  │  2. Set WarningMinDiskSpaceMB = 2048 for early warning alerts        │
  │  3. Audit large local files: Get-ChildItem C:\Users -Recurse        │
  │     -File | Sort Length -Desc | Select -First 20                    │
  └───────────────────────────────────────────────────────────────────────┘

  ┌─ ISSUE 4: Multiple OneDrive sign-in prompts ─────────────────────────┐
  │  Symptom: Users get prompted to sign in every session                │
  │  Cause:   SilentAccountConfig not applied, or AllowTenantList        │
  │           mismatch, or FSLogix not persisting OneDrive account data  │
  │  Fix:                                                                │
  │  1. Verify SilentAccountConfig = 1 (HKLM Policies)                  │
  │  2. Ensure FSLogix profile is mounting correctly (Event ID 27)       │
  │  3. Do NOT exclude %APPDATA%\Microsoft\OneDrive from FSLogix profile │
  │     (OneDrive account state must roam with the profile)              │
  └───────────────────────────────────────────────────────────────────────┘

  ┌─ ISSUE 5: FSLogix profile slow logon after enabling KFM ────────────┐
  │  Symptom: Logon times increased after KFM rollout                   │
  │  Cause:   OneDrive sync starting during logon, competing with       │
  │           FSLogix VHDX mount                                        │
  │  Fix:                                                               │
  │  1. Set PreventNetworkTrafficPreUserSignIn = 1 (defers OneDrive     │
  │     until after logon completes)                                    │
  │  2. Verify FSLogix VHDLocations exclusions updated (Desktop etc.)  │
  │     removed from profile VHDX since they now live in OneDrive       │
  │  3. Check Redirections.xml has KFM folder exclusions applied        │
  └───────────────────────────────────────────────────────────────────────┘

  ┌─ ISSUE 6: Personal accounts appearing in OneDrive ──────────────────┐
  │  Symptom: Users can add personal Microsoft accounts                 │
  │  Fix:     DisablePersonalSync = 1 + AllowTenantList = TenantID     │
  └───────────────────────────────────────────────────────────────────────┘

"@ -ForegroundColor $FC.Muted
    Pause-Screen
}

function Export-OneDriveKQL {
    fx-sec "OneDrive KQL Queries for Log Analytics"
    $file = ".\BDF-OneDrive-KQL-Queries.kql"
    @"
// ═══════════════════════════════════════════════════════════════════════
// BDF OneDrive KQL Query Library — AVD Multi-Session
// ═══════════════════════════════════════════════════════════════════════

// [1] OneDrive sign-in events (from Entra ID logs)
SigninLogs
| where TimeGenerated > ago(24h)
| where AppDisplayName has "OneDrive"
| summarize SignIns=count(), FailedSignIns=countif(ResultType != 0)
            by UserPrincipalName, ResultType, ResultDescription
| order by FailedSignIns desc

// [2] KFM completion events (Windows Event via AMA)
Event
| where Source == "OneDrive"
| where EventID == 7043   // KFM complete event
| where TimeGenerated > ago(7d)
| project TimeGenerated, Computer, UserName=extract("user (.+?) ", 1, RenderedDescription)
| summarize KFMComplete=count() by Computer

// [3] OneDrive sync errors on session hosts
Event
| where Source == "OneDrive"
| where EventLog == "Application"
| where EventLevelName in ("Error","Warning")
| where TimeGenerated > ago(24h)
| project TimeGenerated, Computer, EventID, RenderedDescription
| order by TimeGenerated desc

// [4] Files On Demand usage — dehydrated vs hydrated files
Perf
| where TimeGenerated > ago(1h)
| where ObjectName == "OneDrive" and CounterName has "Files"
| summarize avg(CounterValue) by Computer, CounterName

// [5] OneDrive disk space impact on session hosts
Perf
| where TimeGenerated > ago(1h)
| where ObjectName == "LogicalDisk" and CounterName == "% Free Space"
| where InstanceName == "C:"
| where CounterValue < 15   // Less than 15% free
| summarize AvgFreeSpace=avg(CounterValue) by Computer
| extend Warning = "LOW DISK — check OneDrive Files On Demand setting"
| order by AvgFreeSpace asc
"@ | Set-Content $file -Encoding UTF8
    fx-ok "OneDrive KQL queries" $file
    Pause-Screen
}

function Export-ODDiagScript {
    fx-sec "Generating OneDrive Diagnostic Collection Script"
    @"
<# OneDrive Diagnostics Collection — run on session host, output to \\share or email #>
`$out = "C:\Temp\ODDiag-`$(Get-Date -Format 'yyyyMMdd-HHmmss')"
New-Item -Path `$out -ItemType Directory -Force | Out-Null

# Registry dump
reg export "HKLM\SOFTWARE\Policies\Microsoft\OneDrive" "`$out\HKLM-OD-Policies.reg" /y 2>&1
reg export "HKCU\SOFTWARE\Microsoft\OneDrive"          "`$out\HKCU-OD-Config.reg"   /y 2>&1

# OneDrive log files
`$odLog = "`$env:LOCALAPPDATA\Microsoft\OneDrive\logs"
if (Test-Path `$odLog) { Copy-Item `$odLog -Destination "`$out\Logs" -Recurse -EA SilentlyContinue }

# Process status
Get-Process -Name OneDrive* -EA SilentlyContinue | Select-Object Name,Id,CPU,WorkingSet |
    Export-Csv "`$out\OD-Processes.csv" -NoTypeInformation

# Sync status via SyncEngine
`$odExe = "`$env:ProgramFiles\Microsoft OneDrive\OneDrive.exe"
if (Test-Path `$odExe) { & `$odExe /help *> "`$out\OD-Version.txt" }

# Known Folder state
`$kfmRegs = @("Desktop","Documents","Pictures") | ForEach-Object {
    `$val = (Get-ItemProperty "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" -Name `$_ -EA SilentlyContinue)
    @{ Folder=`$_; CurrentPath=`$val.`$_ }
}
`$kfmRegs | ConvertTo-Json | Set-Content "`$out\KFM-FolderState.json"

Compress-Archive -Path `$out -DestinationPath "`$out.zip" -Force
Write-Host "Diagnostics collected: `$out.zip" -ForegroundColor Green
"@ | Set-Content ".\Collect-OneDriveDiagnostics.ps1" -Encoding UTF8
    fx-ok "Diagnostic collection script" ".\Collect-OneDriveDiagnostics.ps1"
    Pause-Screen
}

function Show-OneDriveSettingsReference {
    fx-hdr "ONEDRIVE SETTINGS REFERENCE"
    foreach ($cat in $ODSettings.GetEnumerator()) {
        fx-sec $cat.Key
        foreach ($s in $cat.Value) {
            Write-Host ("  [{0,-4}] {1,-45} = {2}" -f $s.Hive, $s.Key, $s.Val) -ForegroundColor $FC.Menu
            Write-Host ("          {0}" -f $s.Desc) -ForegroundColor $FC.Muted
            Write-Host ""
        }
    }
    Pause-Screen
}
#endregion



function Save-FXConfig {
    $Script:Cfg | ConvertTo-Json -Depth 10 | Set-Content $ConfigFile -Encoding UTF8
    fx-log "Config saved to $ConfigFile"
}

function Show-FXLogo {
    if ($NoLogo) { return }
    Clear-Host
    Write-Host @"

`e[34m  ╔═══════════════════════════════════════════════════════════════════╗
  ║  `e[36m FSLogix  │  App Masking  │  App Attach  `e[34m                        ║
  ║  `e[97m BDF Azure Virtual Desktop — Advanced Module v1.0 `e[34m               ║
  ╚═══════════════════════════════════════════════════════════════════╝`e[0m
"@
    Write-Host "  Config : $ConfigFile" -ForegroundColor $FC.Muted
    Write-Host "  Log    : $LogFile" -ForegroundColor $FC.Muted
    Write-Host ""
}

function Show-FXMainMenu {
    while ($true) {
        Show-FXLogo

        $fxState  = if ($Script:Cfg.FSLogix.Deployed)        {"✔ Configured"} else {"○ Not Configured"}
        $amState  = if ($Script:Cfg.AppMasking.Deployed)     {"✔ Configured"} else {"○ Not Configured"}
        $aaState  = if ($Script:Cfg.AppAttach.Enabled)       {"✔ Configured"} else {"○ Not Configured"}
        $odState  = if ($Script:Cfg.OneDrive.Deployed)       {"✔ Configured"} else {"○ Not Configured"}
        $fxCol    = if ($Script:Cfg.FSLogix.Deployed)       {$FC.Ok} else {$FC.Muted}
        $amCol    = if ($Script:Cfg.AppMasking.Deployed)    {$FC.Ok} else {$FC.Muted}
        $aaCol    = if ($Script:Cfg.AppAttach.Enabled)      {$FC.Ok} else {$FC.Muted}
        $odCol    = if ($Script:Cfg.OneDrive.Deployed)      {$FC.Ok} else {$FC.Muted}

        Write-Host "  ┌─ Module Status ──────────────────────────────────────────────────┐" -ForegroundColor $FC.Border
        Write-Host ("  │  FSLogix Profiles    : {0,-43}│" -f $fxState) -ForegroundColor $fxCol
        Write-Host ("  │  App Masking         : {0,-43}│" -f $amState) -ForegroundColor $amCol
        Write-Host ("  │  App Attach          : {0,-43}│" -f $aaState) -ForegroundColor $aaCol
        Write-Host ("  │  OneDrive KFM        : {0,-43}│" -f $odState) -ForegroundColor $odCol
        Write-Host "  └──────────────────────────────────────────────────────────────────┘" -ForegroundColor $FC.Border

        Write-Host ""
        Write-Host "  ┌─ Menu ───────────────────────────────────────────────────────────┐" -ForegroundColor $FC.Border
        Write-Host "  │  1.  FSLogix Profile Container Wizard                           │" -ForegroundColor $FC.Menu
        Write-Host "  │  2.  Redirections.xml Builder (Exclusions)                      │" -ForegroundColor $FC.Menu
        Write-Host "  │  3.  App Masking Manager ▶                                      │" -ForegroundColor $FC.Purple
        Write-Host "  │  4.  App Attach & MSIX App Attach Manager ▶                    │" -ForegroundColor $FC.Azure
        Write-Host "  │  5.  OneDrive — Silent Sign-In & Known Folder Move ▶           │" -ForegroundColor $FC.Azure
        Write-Host "  │  6.  Diagnostics & Health Reporting ▶                           │" -ForegroundColor $FC.Warn
        Write-Host "  │  7.  View FSLogix Best Practice Settings Reference              │" -ForegroundColor $FC.Menu
        Write-Host "  │  0.  Exit                                                       │" -ForegroundColor $FC.Muted
        Write-Host "  └──────────────────────────────────────────────────────────────────┘" -ForegroundColor $FC.Border

        $ch = Read-Inp "Menu" -V @("0","1","2","3","4","5","6","7")
        switch ($ch) {
            "1" { Invoke-FSLogixProfileWizard }
            "2" { Build-RedirectionsXml }
            "3" { Show-AppMaskingMenu }
            "4" { Show-AppAttachMenu }
            "5" { Show-OneDriveMenu }
            "6" { Show-DiagnosticsMenu }
            "7" { Show-BestPracticesReference }
            "0" { Write-Host "`n  Goodbye. Log: $LogFile`n" -ForegroundColor $FC.Muted; exit 0 }
        }
    }
}

function Show-BestPracticesReference {
    fx-hdr "FSLOGIX BEST PRACTICES QUICK REFERENCE"
    foreach ($cat in $FXSettings.GetEnumerator()) {
        fx-sec $cat.Key
        foreach ($s in $cat.Value) {
            Write-Host ("  {0,-48} = {1,-12} [{2}]" -f $s.Key, $s.Val, $s.Type) -ForegroundColor $FC.Menu
            Write-Host ("  {0}" -f $s.Desc) -ForegroundColor $FC.Muted
            Write-Host ""
        }
    }
    Pause-Screen
}
#endregion

# ── Entry Point ──────────────────────────────────────────────────────────
fx-log "BDF AVD FSLogix/AppMasking/AppAttach module started"

# If dot-sourced into AIO script — just load functions, don't show menu
if ($MyInvocation.InvocationName -ne '.') {
    Show-FXMainMenu
}

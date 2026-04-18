#Requires -RunAsAdministrator
#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
    Wipes selected Win32 App registry entries and cache to force Intune re-evaluation.
.DESCRIPTION
    Connects to Microsoft Graph, compares local registry with cloud apps, 
    identifies orphaned entries, and allows selective cleanup.
.NOTES
    Author: Optimized version
    Version: 2.0
#>

[CmdletBinding()]
param(
    [switch]$Force,
    [string]$LogPath = "$env:ProgramData\Intune\Logs\Win32Apps-RegistryWipe.log"
)

#region Functions
function Write-Log {
    param([string]$Message, [ValidateSet('Info','Warning','Error')]$Level = 'Info')
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Level] $Message"
    
    $Color = switch ($Level) {
        'Info'    { 'Cyan' }
        'Warning' { 'Yellow' }
        'Error'   { 'Red' }
    }
    Write-Host $LogMessage -ForegroundColor $Color
    
    # Ensure log directory exists
    $LogDir = Split-Path $LogPath -Parent
    if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory -Force | Out-Null }
    Add-Content -Path $LogPath -Value $LogMessage
}

function Stop-IMEService {
    $Service = Get-Service -Name "IntuneManagementExtension" -ErrorAction SilentlyContinue
    if ($null -eq $Service) {
        Write-Log "Intune Management Extension service not found!" -Level Error
        return $false
    }
    
    if ($Service.Status -eq 'Running') {
        Write-Log "Stopping Intune Management Extension service..."
        Stop-Service -Name "IntuneManagementExtension" -Force -ErrorAction Stop
        
        # Wait for service to fully stop
        $Timeout = 30
        $Timer = [Diagnostics.Stopwatch]::StartNew()
        while ((Get-Service "IntuneManagementExtension").Status -ne 'Stopped' -and $Timer.Elapsed.TotalSeconds -lt $Timeout) {
            Start-Sleep -Milliseconds 500
        }
        $Timer.Stop()
        
        if ((Get-Service "IntuneManagementExtension").Status -ne 'Stopped') {
            Write-Log "Failed to stop service within $Timeout seconds" -Level Error
            return $false
        }
    }
    Write-Log "Service stopped successfully"
    return $true
}

function Start-IMEService {
    Write-Log "Starting Intune Management Extension service..."
    Start-Service -Name "IntuneManagementExtension" -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    
    $Status = (Get-Service "IntuneManagementExtension").Status
    if ($Status -eq 'Running') {
        Write-Log "Service started successfully"
    } else {
        Write-Log "Service status: $Status" -Level Warning
    }
}
#endregion

#region Main Script
try {
    Write-Log "=== Win32 Apps Registry Wipe Started ==="
    
    # Connect to Graph
    Write-Log "Connecting to Microsoft Graph..."
    try {
        Connect-MgGraph -Scopes "DeviceManagementApps.Read.All" -NoWelcome -ErrorAction Stop
    } catch {
        Write-Log "Failed to connect to Graph: $_" -Level Error
        return
    }
    
    # Fetch cloud apps
    Write-Log "Fetching Win32 Apps from Intune..."
    $Uri = "https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps?`$filter=isof('microsoft.graph.win32LobApp')"
    
    try {
        $Response = Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop
        $CloudApps = $Response.value | Select-Object displayName, id
        
        # Handle pagination
        while ($Response.'@odata.nextLink') {
            $Response = Invoke-MgGraphRequest -Method GET -Uri $Response.'@odata.nextLink'
            $CloudApps += $Response.value | Select-Object displayName, id
        }
        Write-Log "Found $($CloudApps.Count) Win32 Apps in Intune"
    } catch {
        Write-Log "Failed to fetch apps from Graph: $_" -Level Error
        return
    }
    
    # Check registry
    $Win32RegBase = "HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps"
    
    if (-not (Test-Path $Win32RegBase)) {
        Write-Log "Intune registry path not found: $Win32RegBase" -Level Error
        Write-Log "Is the Intune Management Extension installed?" -Level Warning
        return
    }
    
    # Scan for App IDs (GUIDs under user SIDs)
    $UserSIDs = Get-ChildItem -Path $Win32RegBase -ErrorAction SilentlyContinue | 
                Where-Object { $_.PSChildName -match '^S-1-' -or $_.PSChildName -eq '00000000-0000-0000-0000-000000000000' }
    
    $LocalApps = @()
    foreach ($SID in $UserSIDs) {
        $AppKeys = Get-ChildItem -Path $SID.PSPath -ErrorAction SilentlyContinue | 
                   Where-Object { $_.PSChildName -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$' }
        foreach ($AppKey in $AppKeys) {
            $LocalApps += [PSCustomObject]@{
                SID   = $SID.PSChildName
                AppID = $AppKey.PSChildName
                Path  = $AppKey.PSPath
            }
        }
    }
    
    $UniqueAppIDs = $LocalApps.AppID | Select-Object -Unique
    Write-Log "Found $($UniqueAppIDs.Count) unique App IDs in registry"
    
    if ($UniqueAppIDs.Count -eq 0) {
        Write-Log "No Win32 App IDs found in registry" -Level Warning
        return
    }
    
    # Build report
    $Report = foreach ($AppID in $UniqueAppIDs) {
        $CloudMatch = $CloudApps | Where-Object { $_.id -eq $AppID }
        [PSCustomObject]@{
            AppName    = if ($CloudMatch) { $CloudMatch.displayName } else { "ORPHANED (Deleted from Portal)" }
            AppID      = $AppID
            IsOrphaned = [bool](-not $CloudMatch)
            KeyCount   = ($LocalApps | Where-Object { $_.AppID -eq $AppID }).Count
        }
    }
    
    # Show selection grid
    $SelectedApps = $Report | Sort-Object IsOrphaned -Descending | 
                    Out-GridView -Title "Select Apps to Wipe (Orphaned apps shown first)" -PassThru
    
    if ($null -eq $SelectedApps -or $SelectedApps.Count -eq 0) {
        Write-Log "No apps selected. Exiting."
        return
    }
    
    Write-Log "Selected $($SelectedApps.Count) apps for cleanup"
    
    # Stop service
    if (-not (Stop-IMEService)) {
        if (-not $Force) {
            Write-Log "Cannot proceed without stopping service. Use -Force to override." -Level Error
            return
        }
        Write-Log "Proceeding anyway due to -Force flag" -Level Warning
    }
    
    # Cleanup
    foreach ($App in $SelectedApps) {
        Write-Log "Cleaning: $($App.AppName) [$($App.AppID)]"
        
        # Remove registry keys
        $TargetKeys = Get-ChildItem -Path $Win32RegBase -Recurse -ErrorAction SilentlyContinue | 
                      Where-Object { $_.PSChildName -eq $App.AppID }
        
        foreach ($Key in $TargetKeys) {
            try {
                Remove-Item -Path $Key.PSPath -Recurse -Force -ErrorAction Stop
                Write-Log "  Removed registry: $($Key.PSPath)"
            } catch {
                Write-Log "  Failed to remove registry key: $_" -Level Warning
            }
        }
        
        # Remove cache locations
        $CachePaths = @(
            "C:\Windows\IMECache\$($App.AppID)",
            "$env:ProgramData\Microsoft\IntuneManagementExtension\Cache\$($App.AppID)"
        )
        
        foreach ($CachePath in $CachePaths) {
            if (Test-Path $CachePath) {
                try {
                    Remove-Item -Path $CachePath -Recurse -Force -ErrorAction Stop
                    Write-Log "  Removed cache: $CachePath"
                } catch {
                    Write-Log "  Failed to remove cache: $_" -Level Warning
                }
            }
        }
    }
    
    # Restart service
    Start-IMEService
    
    Write-Log "=== Cleanup Complete ==="
    Write-Log "Intune will re-evaluate apps within 5-15 minutes, or trigger sync from Company Portal"
    
} catch {
    Write-Log "Unexpected error: $_" -Level Error
} finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}
#endregion
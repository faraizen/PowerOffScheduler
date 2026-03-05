<#
.SYNOPSIS
    Advanced system power-off scheduler with countdown, cancellation, and logging.

.DESCRIPTION
    This script schedules a shutdown, restart, logoff, or forces all applications to close after a specified number of minutes.
    It features a real-time countdown with progress bar, interactive cancellation (press C), logging, and the ability to abort any pending action.
    Perfect for automating system maintenance or forcing a cleanup before leaving.

.PARAMETER Minutes
    Number of minutes to wait before executing the action. Default is 5.

.PARAMETER Action
    Action to perform: Shutdown, Restart, Logoff, or CloseApps. Default is Shutdown.

.PARAMETER Force
    If specified, the action will be executed immediately after countdown without additional confirmation (though you can still cancel during countdown).

.PARAMETER Abort
    If specified, cancels any previously scheduled action (by this script) and removes the marker file.

.PARAMETER KillAll
    Caution: When used with -Action CloseApps, forces termination of ALL user processes (including background ones), not just visible windows. Use with extreme care.

.EXAMPLE
    .\PowerOffScheduler.ps1 -Minutes 10 -Action Restart -Force
    Schedules a system restart in 10 minutes and proceeds without asking.

.EXAMPLE
    .\PowerOffScheduler.ps1 -Abort
    Aborts any pending scheduled action.

.EXAMPLE
    .\PowerOffScheduler.ps1 -Minutes 15 -Action CloseApps
    Closes all user applications after 15 minutes (with countdown and cancellation option).

.NOTES
    Author: Your GitHub Username
    Date: 2025
    Requires: Administrator rights for Shutdown, Restart, and Logoff actions. CloseApps may also benefit from admin rights to close elevated applications.
    Marker file: %TEMP%\PowerOffScheduler.marker
    Log file: %TEMP%\PowerOffScheduler.log

.LINK
    https://github.com/yourusername/PowerOffScheduler
#>

[CmdletBinding(DefaultParameterSetName = 'Schedule')]
param(
    [Parameter(ParameterSetName = 'Schedule')]
    [int]$Minutes = 5,

    [Parameter(ParameterSetName = 'Schedule')]
    [ValidateSet('Shutdown', 'Restart', 'Logoff', 'CloseApps')]
    [string]$Action = 'Shutdown',

    [Parameter(ParameterSetName = 'Schedule')]
    [switch]$Force,

    [Parameter(ParameterSetName = 'Schedule')]
    [switch]$KillAll,

    [Parameter(ParameterSetName = 'Abort', Mandatory = $true)]
    [switch]$Abort
)

#region Functions
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "[$timestamp] [$Level] $Message"
    Write-Host $logEntry
    try {
        $logEntry | Out-File -FilePath $logFile -Append -Encoding utf8
    }
    catch {
        Write-Warning "Failed to write to log: $_"
    }
}

function Test-Admin {
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Set-Marker {
    param([datetime]$TargetTime)
    $marker = @{
        TargetTime = $TargetTime
        Action = $Action
        Minutes = $Minutes
        Force = $Force.IsPresent
        KillAll = $KillAll.IsPresent
        User = [Environment]::UserName
        Computer = [Environment]::MachineName
    }
    $marker | ConvertTo-Json | Set-Content -Path $markerFile -Force
    Write-Log "Marker set: Action '$Action' scheduled at $TargetTime"
}

function Remove-Marker {
    if (Test-Path $markerFile) {
        Remove-Item -Path $markerFile -Force
        Write-Log "Marker removed."
    }
}

function Check-Marker {
    if (Test-Path $markerFile) {
        try {
            $marker = Get-Content -Path $markerFile -Raw | ConvertFrom-Json
            return $marker
        }
        catch {
            Write-Log "Marker file corrupted. Removing it." -Level 'WARN'
            Remove-Marker
        }
    }
    return $null
}

function Stop-UserApplications {
    Write-Log "Closing user applications..." -Level 'ACTION'

    if ($KillAll) {
        # Kill all processes owned by the current user (excluding critical system processes)
        Write-Log "KillAll mode: Terminating all user processes (be careful)..." -Level 'WARN'
        $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        Get-Process | Where-Object { $_.SI -eq (Get-Process -Id $pid).SessionId } | ForEach-Object {
            try {
                $_.Kill()
                Write-Log "Killed process: $($_.Name) (PID: $($_.Id))"
            }
            catch {
                Write-Log "Failed to kill $($_.Name): $_" -Level 'WARN'
            }
        }
    }
    else {
        # Only close applications with visible windows (safer)
        $apps = Get-Process | Where-Object { $_.MainWindowTitle -ne '' -and $_.SessionId -eq (Get-Process -Id $pid).SessionId }
        foreach ($app in $apps) {
            try {
                $app.CloseMainWindow() | Out-Null
                Start-Sleep -Milliseconds 200
                if (!$app.HasExited) {
                    $app.Kill()
                    Write-Log "Force killed: $($app.Name) (PID: $($app.Id))"
                }
                else {
                    Write-Log "Closed gracefully: $($app.Name)"
                }
            }
            catch {
                Write-Log "Error closing $($app.Name): $_" -Level 'WARN'
            }
        }
    }
    Write-Log "Application closure completed."
}

function Invoke-Action {
    param(
        [string]$Action,
        [switch]$Force
    )
    Write-Log "Executing action: $Action" -Level 'ACTION'

    # Additional confirmation if not forced
    if (-not $Force) {
        Write-Host "Proceed with $Action? (Y/N) " -NoNewline -ForegroundColor Yellow
        $key = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        if ($key.Character -ne 'y' -and $key.Character -ne 'Y') {
            Write-Log "Action cancelled by user at final confirmation."
            Remove-Marker
            return
        }
    }

    switch ($Action) {
        'Shutdown' {
            if (-not (Test-Admin)) {
                Write-Log "Administrator rights required for shutdown. Restarting as admin..." -Level 'ERROR'
                # Attempt to self-elevate
                Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Minutes $Minutes -Action $Action -Force:$Force -KillAll:$KillAll"
                return
            }
            Stop-Computer -Force
        }
        'Restart' {
            if (-not (Test-Admin)) {
                Write-Log "Administrator rights required for restart. Restarting as admin..." -Level 'ERROR'
                Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Minutes $Minutes -Action $Action -Force:$Force -KillAll:$KillAll"
                return
            }
            Restart-Computer -Force
        }
        'Logoff' {
            # Logoff can be done without admin, but to ensure all processes close we use shutdown.exe
            shutdown /l /f
        }
        'CloseApps' {
            Stop-UserApplications
        }
    }
    Remove-Marker
}

function Show-Countdown {
    param(
        [int]$Seconds,
        [string]$Action
    )
    $endTime = (Get-Date).AddSeconds($Seconds)
    Write-Host "`n⏳ Action: $Action will be performed in $Minutes minute(s)." -ForegroundColor Cyan
    Write-Host "Press 'C' at any time to cancel.`n"

    # Notify user via popup (if possible)
    try {
        $popup = New-Object -ComObject Wscript.Shell
        $popup.Popup("$Action scheduled in $Minutes minutes.`nPress Cancel in console to abort.", 5, "PowerOff Scheduler", 64)
    }
    catch {
        # Ignore if no UI
    }

    # Send message to all sessions (requires admin)
    if (Test-Admin) {
        msg * "System $Action in $Minutes minutes. Please save your work."
    }

    # Countdown loop
    while ($endTime -gt (Get-Date)) {
        $remaining = ($endTime - (Get-Date)).TotalSeconds
        if ($remaining -le 0) { break }

        # Update progress bar
        $percent = [math]::Round((($Seconds - $remaining) / $Seconds) * 100, 2)
        Write-Progress -Activity "Countdown to $Action" -Status "Time left: $([math]::Round($remaining)) seconds" -PercentComplete $percent -SecondsRemaining $remaining

        # Check for user cancellation (press C)
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            if ($key.Key -eq 'C') {
                Write-Host "`n⛔ Cancelled by user." -ForegroundColor Red
                Write-Log "Countdown cancelled by user."
                Remove-Marker
                return $false
            }
        }

        Start-Sleep -Seconds 1
    }

    Write-Progress -Activity "Countdown to $Action" -Completed
    return $true
}
#endregion

#region Main script
# Setup paths
$markerFile = Join-Path -Path $env:TEMP -ChildPath 'PowerOffScheduler.marker'
$logFile = Join-Path -Path $env:TEMP -ChildPath 'PowerOffScheduler.log'

# Handle Abort switch
if ($Abort) {
    $existing = Check-Marker
    if ($existing) {
        Write-Host "Aborting scheduled $($existing.Action) at $($existing.TargetTime)." -ForegroundColor Yellow
        Remove-Marker
        # Also abort any system shutdown if it was triggered via shutdown.exe
        shutdown /a 2>$null
        Write-Log "Abort command issued."
    }
    else {
        Write-Host "No scheduled action found." -ForegroundColor Green
    }
    exit 0
}

# Check for existing marker (another instance may have scheduled something)
$existing = Check-Marker
if ($existing) {
    Write-Host "⚠️  A scheduled action already exists:" -ForegroundColor Yellow
    Write-Host "   Action: $($existing.Action)" -ForegroundColor Cyan
    Write-Host "   Time: $($existing.TargetTime) (local)" -ForegroundColor Cyan
    Write-Host "   Minutes: $($existing.Minutes)" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Do you want to overwrite it with new schedule? (Y/N) " -NoNewline -ForegroundColor Yellow
    $key = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    if ($key.Character -ne 'y' -and $key.Character -ne 'Y') {
        Write-Host "Operation cancelled."
        exit 0
    }
    Remove-Marker
}

# Validate minutes
if ($Minutes -le 0) {
    Write-Error "Minutes must be greater than 0."
    exit 1
}

# Log start
Write-Log "=== PowerOff Scheduler started ==="
Write-Log "Action: $Action, Minutes: $Minutes, Force: $Force, KillAll: $KillAll"

# Check admin rights for actions that need it
if ($Action -in @('Shutdown','Restart') -and -not (Test-Admin)) {
    Write-Warning "Action '$Action' requires administrator privileges. Attempting to elevate..."
    # Re-launch as admin
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Minutes $Minutes -Action $Action"
    if ($Force) { $arguments += " -Force" }
    if ($KillAll) { $arguments += " -KillAll" }
    Start-Process powershell -Verb RunAs -ArgumentList $arguments
    exit 0
}

# Set marker with target time
$targetTime = (Get-Date).AddMinutes($Minutes)
Set-Marker -TargetTime $targetTime

# Run countdown
$continue = Show-Countdown -Seconds ($Minutes * 60) -Action $Action
if (-not $continue) {
    # Cancelled by user
    Remove-Marker
    exit 0
}

# Execute action
Invoke-Action -Action $Action -Force:$Force

Write-Log "Script completed."
#endregion
